from __future__ import annotations

import html
import os
import mimetypes
from calendar import monthrange
from datetime import date
from decimal import Decimal
from http.cookies import SimpleCookie
from pathlib import Path
from urllib.parse import parse_qs, quote, unquote, urlparse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

from .generator import (
    convert_to_pdf,
    date_for_display,
    export_history_csv,
    format_cell_phone,
    format_cpf,
    money,
    parse_decimal,
    require_valid_cell_phone,
    require_valid_cpf,
    render_orcamento,
    today_iso,
    validate_template,
)
from .storage import (
    Item,
    Orcamento,
    add_farmaceutico,
    connect,
    create_session,
    create_user,
    delete_session,
    delete_farmaceutico,
    delete_orcamento,
    get_user_by_session,
    get_itens,
    get_orcamento,
    has_users,
    init_db,
    list_farmaceuticos,
    list_orcamentos,
    save_orcamento,
    verify_login,
)
from .temperature_map import (
    MONTH_OPTIONS,
    convert_temperature_map_to_pdf,
    list_temperature_maps,
    output_file as temperature_output_file,
    parse_temperature_map_input,
    render_temperature_map,
)


HOST = os.environ.get("HOST", "0.0.0.0" if os.environ.get("RENDER") else "127.0.0.1")
PORT = 8000
ASSETS_DIR = Path("assets")
LOGO_FILENAME = "logo_preco_popular.svg"
SESSION_COOKIE = "orcamento_session"
SESSION_MAX_AGE = 12 * 60 * 60


def esc(value: object) -> str:
    return html.escape("" if value is None else str(value), quote=True)


def parse_form(body: bytes) -> dict[str, list[str]]:
    return parse_qs(body.decode("utf-8"), keep_blank_values=True)


def first(data: dict[str, list[str]], key: str, default: str = "") -> str:
    values = data.get(key, [default])
    return values[0].strip()


def build_items(data: dict[str, list[str]]) -> list[Item]:
    codigos = data.get("codigo", [])
    produtos = data.get("produto", [])
    pmcs = data.get("pmc", [])
    descontos = data.get("desconto", [])
    valores = data.get("valor_unitario", [])
    quantidades = data.get("quantidade", [])
    items: list[Item] = []

    for idx, produto in enumerate(produtos):
        produto = produto.strip()
        codigo = codigos[idx].strip() if idx < len(codigos) else ""
        valor_raw = valores[idx].strip() if idx < len(valores) else ""
        qtd_raw = quantidades[idx].strip() if idx < len(quantidades) else ""
        if not any([produto, codigo, valor_raw, qtd_raw]):
            continue
        if not produto:
            raise ValueError("Informe o nome do produto em todas as linhas usadas.")
        item = Item(
            codigo=codigo,
            produto=produto,
            pmc=pmcs[idx].strip() if idx < len(pmcs) else "",
            desconto=descontos[idx].strip() if idx < len(descontos) else "",
            valor_unitario=parse_decimal(valor_raw),
            quantidade=parse_decimal(qtd_raw or "1"),
        )
        items.append(item)

    if not items:
        raise ValueError("Informe pelo menos um produto.")
    return items


def row_to_orcamento(orcamento_id: int, row) -> Orcamento:
    itens = [
        Item(
            codigo=item["codigo"],
            produto=item["produto"],
            pmc=item["pmc"],
            desconto=item["desconto"],
            valor_unitario=Decimal(item["valor_unitario"]),
            quantidade=Decimal(item["quantidade"]),
        )
        for item in get_itens(orcamento_id)
    ]
    return Orcamento(
        cliente_nome=row["cliente_nome"],
        farmaceutico_responsavel=row["farmaceutico_responsavel"],
        cpf=row["cpf"],
        telefone=row["telefone"],
        email=row["email"],
        data_orcamento=row["data_orcamento"],
        localidade=row["localidade"],
        itens=itens,
    )


def existing_path(value: str | None) -> Path | None:
    if not value:
        return None
    candidates = [Path(value)]
    normalized = value.replace("\\", "/")
    if normalized != value:
        candidates.append(Path(normalized))
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return None


def update_orcamento_files(
    orcamento_id: int,
    *,
    docx_path: Path | None = None,
    pdf_path: Path | None = None,
    pdf_status: str | None = None,
) -> None:
    assignments = []
    values: list[str | None | int] = []
    if docx_path is not None:
        assignments.append("docx_path = ?")
        values.append(str(docx_path))
    if pdf_path is not None:
        assignments.append("pdf_path = ?")
        values.append(str(pdf_path))
    if pdf_status is not None:
        assignments.append("pdf_status = ?")
        values.append(pdf_status)
    if not assignments:
        return
    values.append(orcamento_id)
    with connect() as conn:
        conn.execute(
            f"UPDATE orcamentos SET {', '.join(assignments)} WHERE id = ?",
            values,
        )


def ensure_docx_file(orcamento_id: int, row) -> Path:
    existing = existing_path(row["docx_path"])
    if existing is not None:
        return existing
    orcamento = row_to_orcamento(orcamento_id, row)
    docx_path = render_orcamento(orcamento, sequence=orcamento_id)
    update_orcamento_files(orcamento_id, docx_path=docx_path)
    return docx_path


def layout(title: str, content: str, *, show_nav: bool = True) -> bytes:
    nav_html = ""
    if show_nav:
        nav_html = """
    <input class="nav-toggle" id="nav-toggle" type="checkbox" aria-label="Abrir menu">
    <label class="nav-toggle-button" for="nav-toggle">Menu</label>
    <nav>
      <a href="/">Novo orçamento</a>
      <a href="/mapa-temperatura">Mapa temperatura</a>
      <a href="/historico">Histórico</a>
      <a href="/farmaceuticos">Farmacêuticos</a>
      <a href="/validar-modelo">Validar modelo</a>
      <a href="/exportar">Exportar CSV</a>
      <form method="post" action="/logout">
        <button class="nav-logout" type="submit">Sair</button>
      </form>
    </nav>
        """
    page = f"""<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{esc(title)}</title>
  <style>
    :root {{
      --bg: #f5f7f6;
      --panel: #ffffff;
      --ink: #202724;
      --muted: #65716b;
      --line: #dfe6e1;
      --line-strong: #cbd8d0;
      --accent: rgb(7, 143, 71);
      --accent-dark: #056d36;
      --accent-soft: #edf7f1;
      --brand-red: #d71932;
      --brand-red-dark: #b91127;
      --ok: #078f47;
      --warn: #8a5a00;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: "Segoe UI", Arial, Helvetica, sans-serif;
      background: var(--bg);
      color: var(--ink);
    }}
    header {{
      background: #ffffff;
      border-bottom: 1px solid var(--line-strong);
      padding: 14px 22px;
      display: grid;
      grid-template-columns: auto minmax(520px, 1fr);
      align-items: center;
      gap: 14px;
      box-shadow: 0 1px 0 rgba(7, 143, 71, .04);
    }}
    .brand {{
      display: flex;
      align-items: center;
      gap: 14px;
      min-width: 0;
    }}
    .brand-logo {{
      width: 142px;
      height: auto;
      display: block;
      flex: 0 0 auto;
    }}
    header h1 {{
      margin: 0;
      font-size: 18px;
      color: rgb(7, 143, 71);
      letter-spacing: 0;
      line-height: 1.2;
      white-space: nowrap;
    }}
    nav {{
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      justify-content: flex-end;
      gap: 6px;
    }}
    .nav-toggle,
    .nav-toggle-button {{
      display: none;
    }}
    nav form {{
      margin: 0;
    }}
    nav a {{
      white-space: nowrap;
      color: #2e4f3d;
      text-decoration: none;
      font-weight: 700;
      margin-left: 0;
      padding: 7px 8px;
      border-radius: 6px;
      transition: background .16s ease, color .16s ease;
    }}
    nav a:hover {{
      background: var(--accent-soft);
      color: var(--accent-dark);
    }}
    .nav-logout {{
      min-height: 32px;
      height: 32px;
      padding: 6px 10px;
      background: #eef3f0;
      color: #2e4f3d;
      border: 1px solid var(--line);
    }}
    .nav-logout:hover {{
      background: #e1ebe5;
    }}
    main {{
      width: min(1120px, calc(100vw - 32px));
      margin: 24px auto 48px;
    }}
    section {{
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 20px;
      margin-bottom: 18px;
      box-shadow: 0 1px 2px rgba(32, 39, 36, .035);
    }}
    h2 {{
      font-size: 16px;
      margin: 0 0 16px;
      color: #22382d;
    }}
    label {{
      display: block;
      font-size: 13px;
      color: var(--muted);
      margin-bottom: 6px;
      font-weight: 700;
    }}
    input,
    select {{
      width: 100%;
      height: 38px;
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 8px 10px;
      font-size: 14px;
      background: #fff;
      color: var(--ink);
    }}
    input:focus,
    select:focus {{
      outline: 2px solid rgba(7, 143, 71, .16);
      border-color: var(--accent);
    }}
    input.invalid,
    select.invalid {{
      border-color: #b42318;
      background: #fff8f8;
    }}
    .field-error {{
      color: #8a1f1f;
      font-size: 12px;
      font-weight: 700;
      margin-top: 6px;
    }}
    .grid {{
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 14px;
    }}
    .span-2 {{ grid-column: span 2; }}
    .span-4 {{ grid-column: span 4; }}
    table {{
      width: 100%;
      border-collapse: collapse;
      table-layout: fixed;
    }}
    th, td {{
      border-bottom: 1px solid var(--line);
      padding: 8px;
      vertical-align: middle;
    }}
    th {{
      text-align: left;
      font-size: 12px;
      color: var(--muted);
      background: #fafcfb;
    }}
    th:nth-child(1) {{ width: 95px; }}
    th:nth-child(2) {{ width: auto; }}
    th:nth-child(3), th:nth-child(4), th:nth-child(5), th:nth-child(6) {{ width: 105px; }}
    th:nth-child(7) {{ width: 58px; }}
    td input {{ height: 34px; }}
    .actions {{
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      align-items: center;
      justify-content: flex-end;
    }}
    .final-actions {{
      justify-content: flex-start;
      gap: 18px;
      align-items: end;
    }}
    .responsavel-select {{
      flex: 1 1 320px;
      max-width: 430px;
    }}
    .final-actions .primary {{
      height: 38px;
      min-height: 38px;
      padding-top: 0;
      padding-bottom: 0;
      margin-bottom: 0;
    }}
    button, .button {{
      border: 0;
      border-radius: 6px;
      padding: 10px 14px;
      font-weight: 700;
      text-decoration: none;
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-height: 38px;
      transition: background .16s ease, box-shadow .16s ease, transform .16s ease;
    }}
    button.primary, .button.primary {{
      background: var(--accent);
      color: #fff;
    }}
    button.primary:hover, .button.primary:hover {{
      background: var(--accent-dark);
      box-shadow: 0 2px 8px rgba(7, 143, 71, .18);
    }}
    button.secondary, .button.secondary {{
      background: #e9efec;
      color: var(--ink);
    }}
    button.secondary:hover, .button.secondary:hover {{
      background: #dfe8e3;
    }}
    button.danger, .button.danger {{
      background: var(--brand-red);
      color: #fff;
    }}
    button.danger:hover, .button.danger:hover {{
      background: var(--brand-red-dark);
    }}
    button.success, .button.success {{
      background: #2fa35e;
      color: #fff;
    }}
    button.success:hover, .button.success:hover {{
      background: #23894d;
    }}
    button.icon {{
      width: 34px;
      min-height: 34px;
      padding: 0;
      background: #eef3f0;
      color: #374151;
    }}
    .history td, .history th {{ font-size: 13px; }}
    .recent-history th:nth-child(1), .recent-history td:nth-child(1) {{ width: 54px; }}
    .recent-history th:nth-child(2), .recent-history td:nth-child(2) {{ width: auto; }}
    .recent-history th:nth-child(3), .recent-history td:nth-child(3) {{ width: 170px; }}
    .recent-history th:nth-child(4), .recent-history td:nth-child(4) {{ width: 94px; }}
    .recent-history th:nth-child(5), .recent-history td:nth-child(5) {{ width: 106px; }}
    .full-history th:nth-child(1), .full-history td:nth-child(1) {{ width: 54px; }}
    .full-history th:nth-child(2), .full-history td:nth-child(2) {{ width: 126px; }}
    .full-history th:nth-child(3), .full-history td:nth-child(3) {{ width: auto; }}
    .full-history th:nth-child(4), .full-history td:nth-child(4) {{ width: 150px; }}
    .full-history th:nth-child(5), .full-history td:nth-child(5) {{
      width: 124px;
      white-space: nowrap;
    }}
    .full-history th:nth-child(6), .full-history td:nth-child(6) {{ width: 94px; }}
    .full-history th:nth-child(7), .full-history td:nth-child(7) {{ width: 106px; }}
    .temperature-history th:first-child,
    .temperature-history td:first-child {{
      width: auto;
      word-break: break-word;
    }}
    .temperature-history th:last-child,
    .temperature-history td:last-child {{
      width: 330px;
      max-width: 330px;
    }}
    .highlight-row {{
      background: var(--accent-soft);
    }}
    .history th:last-child,
    .history td:last-child {{
      width: 128px;
      max-width: 128px;
    }}
    .temperature-history th:last-child,
    .temperature-history td:last-child {{
      width: 330px;
      max-width: 330px;
    }}
    .row-actions {{
      padding-left: 6px;
      padding-right: 6px;
    }}
    .action-group {{
      display: flex;
      align-items: center;
      justify-content: flex-end;
      gap: 6px;
      white-space: nowrap;
      width: 100%;
      min-width: 0;
    }}
    .action-group form {{
      margin: 0;
    }}
    .action-group a,
    .action-group button {{
      flex: 0 0 auto;
      min-height: 30px;
      padding: 6px 8px;
      font-size: 12px;
    }}
    .muted {{ color: var(--muted); }}
    .ok {{ color: var(--ok); font-weight: 700; }}
    .warn {{ color: var(--warn); font-weight: 700; }}
    .success-title {{ color: var(--ok); }}
    .pdf-status {{
      position: relative;
      width: 30px;
      min-height: 30px;
      padding: 0;
      border-radius: 999px;
      border: 1px solid #f2c36b;
      background: #fff7e6;
      color: var(--warn);
      font-weight: 900;
    }}
    .pdf-status .popover {{
      display: none;
      position: absolute;
      right: 0;
      bottom: calc(100% + 8px);
      z-index: 10;
      width: min(320px, calc(100vw - 40px));
      padding: 10px 12px;
      border: 1px solid #e5bf72;
      border-radius: 6px;
      background: #fffaf0;
      color: #5f4300;
      box-shadow: 0 8px 20px rgba(32, 39, 36, .16);
      font-size: 12px;
      font-weight: 700;
      line-height: 1.35;
      white-space: normal;
      overflow-wrap: anywhere;
      text-align: left;
    }}
    .pdf-status:hover .popover,
    .pdf-status:focus .popover,
    .pdf-status:focus-within .popover {{
      display: block;
    }}
    .calendar-picker {{
      margin-top: 16px;
    }}
    .calendar-toggle {{
      justify-content: flex-start;
      margin-bottom: 0;
    }}
    .calendar-collapse {{
      max-height: 0;
      overflow: hidden;
      opacity: 0;
      transition: max-height .24s ease, opacity .18s ease, margin-top .18s ease;
    }}
    .calendar-collapse.open {{
      max-height: 560px;
      opacity: 1;
      margin-top: 10px;
    }}
    .calendar-shell {{
      max-width: 460px;
      border: 1px solid var(--line);
      border-radius: 8px;
      overflow: hidden;
      background: #fff;
    }}
    .calendar-title {{
      padding: 10px 12px;
      color: var(--accent);
      font-size: 20px;
      font-weight: 800;
      line-height: 1.1;
      text-transform: lowercase;
      background: #fafcfb;
      border-bottom: 1px solid var(--line);
    }}
    .calendar-weekdays,
    .calendar-grid {{
      display: grid;
      grid-template-columns: repeat(7, minmax(0, 1fr));
    }}
    .calendar-weekdays span {{
      padding: 6px 4px;
      background: var(--accent);
      color: #fff;
      font-weight: 800;
      font-size: 11px;
      text-align: center;
      border-right: 1px solid rgba(255,255,255,.25);
    }}
    .calendar-weekdays span:last-child {{
      border-right: 0;
    }}
    .day-check {{
      position: relative;
      display: block;
      min-height: 42px;
      border-right: 1px solid #cfd8d2;
      border-bottom: 1px solid #cfd8d2;
      background: #fff;
    }}
    .day-check input {{
      position: absolute;
      opacity: 0;
      pointer-events: none;
    }}
    .day-check span {{
      display: flex;
      align-items: center;
      justify-content: flex-start;
      height: 100%;
      padding: 6px;
      border: 2px solid transparent;
      border-radius: 0;
      background: transparent;
      color: var(--ink);
      font-size: 13px;
      font-weight: 800;
      cursor: pointer;
    }}
    .day-check input:checked + span {{
      border-color: var(--accent);
      background: var(--accent-soft);
      color: var(--accent-dark);
    }}
    .day-check.sunday span {{
      color: var(--accent);
    }}
    .day-check.saturday span {{
      color: var(--accent);
    }}
    .day-check.blank {{
      background: #f3f5f4;
    }}
    .day-check.blank span {{
      cursor: not-allowed;
    }}
    .calendar-legend {{
      display: flex;
      flex-wrap: wrap;
      gap: 8px 14px;
      margin-top: 10px;
      font-size: 12px;
      color: var(--muted);
      font-weight: 700;
    }}
    .manager-row {{
      display: grid;
      grid-template-columns: minmax(220px, 1fr) auto;
      gap: 10px;
      align-items: end;
    }}
    .manager-row button {{
      height: 38px;
      min-height: 38px;
      padding-top: 0;
      padding-bottom: 0;
    }}
    .manager-row.compact {{
      grid-template-columns: minmax(240px, 1fr) auto;
      align-items: center;
    }}
    .manager-row.compact input {{
      height: 38px;
    }}
    .manager-row.compact button {{
      align-self: center;
    }}
    .manager-row.create-row {{
      grid-template-columns: minmax(320px, 640px) auto;
      justify-content: start;
    }}
    .error {{
      background: #fff6f7;
      border-color: #f3c9d0;
      color: #8a1f2b;
      font-weight: 700;
    }}
    .auth-shell {{
      min-height: calc(100vh - 120px);
      display: grid;
      place-items: center;
    }}
    .auth-card {{
      width: min(430px, 100%);
    }}
    .auth-card .brand-logo {{
      width: 160px;
      margin-bottom: 18px;
    }}
    .auth-card form {{
      display: grid;
      gap: 14px;
    }}
    .total-box {{
      font-size: 18px;
      font-weight: 800;
      text-align: right;
    }}
    @media (max-width: 1320px) {{
      header {{ grid-template-columns: 1fr; align-items: flex-start; }}
      nav {{ justify-content: flex-start; }}
    }}
    @media (max-width: 860px) {{
      header {{
        padding: 12px 16px;
        gap: 10px;
      }}
      .brand {{ align-items: flex-start; flex-direction: column; gap: 8px; }}
      .brand-logo {{ width: 138px; }}
      header h1 {{ white-space: normal; }}
      .nav-toggle-button {{
        display: inline-flex;
        align-items: center;
        justify-content: center;
        min-height: 38px;
        width: 100%;
        border: 1px solid var(--line-strong);
        border-radius: 6px;
        background: #eef3f0;
        color: #2e4f3d;
        font-weight: 800;
        cursor: pointer;
      }}
      nav {{
        display: none;
        width: 100%;
        flex-direction: column;
        align-items: stretch;
        gap: 8px;
        padding-top: 2px;
      }}
      .nav-toggle:checked + .nav-toggle-button + nav {{
        display: flex;
      }}
      nav a,
      nav form,
      .nav-logout {{
        width: 100%;
      }}
      nav a {{
        margin: 0;
        padding: 10px 12px;
        background: #fafcfb;
        border: 1px solid var(--line);
      }}
      .nav-logout {{
        justify-content: center;
        height: 38px;
      }}
      main {{
        width: min(1120px, calc(100vw - 24px));
        margin-top: 16px;
      }}
      section {{
        padding: 16px;
      }}
      .grid {{ grid-template-columns: 1fr; }}
      .span-2, .span-4 {{ grid-column: span 1; }}
      #produtos,
      #produtos thead,
      #produtos tbody,
      #produtos tr,
      #produtos td,
      .history.recent-history,
      .history.temperature-history,
      .history.recent-history thead,
      .history.temperature-history thead,
      .history.recent-history tbody,
      .history.temperature-history tbody,
      .history.recent-history tr,
      .history.temperature-history tr,
      .history.recent-history td {{
        display: block;
        width: 100%;
      }}
      .history.temperature-history td {{
        display: block;
        width: 100%;
      }}
      #produtos,
      .history.recent-history,
      .history.temperature-history {{
        min-width: 0;
      }}
      #produtos thead,
      .history.recent-history thead,
      .history.temperature-history thead {{
        display: none;
      }}
      #produtos tr,
      .history.recent-history tr,
      .history.temperature-history tr {{
        border: 1px solid var(--line);
        border-radius: 8px;
        padding: 10px;
        margin-bottom: 12px;
        background: #fff;
      }}
      #produtos td,
      .history.recent-history td,
      .history.temperature-history td {{
        border-bottom: 0;
        padding: 7px 0;
      }}
      #produtos td::before,
      .history.recent-history td::before,
      .history.temperature-history td::before {{
        display: block;
        margin-bottom: 5px;
        color: var(--muted);
        font-size: 12px;
        font-weight: 800;
      }}
      #produtos td:nth-child(1)::before {{ content: "Código"; }}
      #produtos td:nth-child(2)::before {{ content: "Produto"; }}
      #produtos td:nth-child(3)::before {{ content: "PMC"; }}
      #produtos td:nth-child(4)::before {{ content: "Desc. %"; }}
      #produtos td:nth-child(5)::before {{ content: "Valor un."; }}
      #produtos td:nth-child(6)::before {{ content: "Qtd"; }}
      #produtos td:nth-child(7)::before,
      .history.recent-history td:nth-child(6)::before {{
        content: "";
        display: none;
      }}
      #produtos td:last-child {{
        display: flex;
        justify-content: flex-end;
      }}
      .history.recent-history td:nth-child(1)::before {{ content: "ID"; }}
      .history.recent-history td:nth-child(2)::before {{ content: "Cliente"; }}
      .history.recent-history td:nth-child(3)::before {{ content: "Farm. resp."; }}
      .history.recent-history td:nth-child(4)::before {{ content: "Data"; }}
      .history.recent-history td:nth-child(5)::before {{ content: "Total"; }}
      .history.temperature-history td:nth-child(1)::before {{ content: "Arquivo"; }}
      .history.temperature-history td:nth-child(2)::before {{
        content: "";
        display: none;
      }}
      .history.recent-history td:last-child {{
        max-width: none;
        width: 100%;
        padding-top: 10px;
      }}
      .history.temperature-history td:last-child {{
        max-width: none;
        width: 100%;
        padding-top: 10px;
      }}
      .history.recent-history td.muted::before {{
        display: none;
      }}
      .history.recent-history .action-group {{
        justify-content: stretch;
      }}
      .history.temperature-history .action-group {{
        justify-content: stretch;
        flex-wrap: wrap;
      }}
      .history.recent-history .action-group a,
      .history.recent-history .action-group button,
      .history.temperature-history .action-group a,
      .history.temperature-history .action-group button {{
        flex: 1 1 0;
      }}
      .table-wrap {{ overflow-x: auto; }}
      .final-actions {{
        align-items: stretch;
        gap: 12px;
      }}
      .responsavel-select {{
        flex: 1 1 100%;
        max-width: none;
      }}
      .final-actions .primary {{
        width: 100%;
      }}
      .calendar-title {{
        font-size: 18px;
      }}
      .calendar-weekdays span {{
        font-size: 9px;
        padding-left: 2px;
        padding-right: 2px;
      }}
      .day-check {{
        min-height: 38px;
      }}
      .day-check span {{
        font-size: 12px;
        padding: 5px;
      }}
      .calendar-collapse.open {{
        max-height: 520px;
      }}
      .manager-row,
      .manager-row.compact,
      .manager-row.create-row {{ grid-template-columns: 1fr; }}
    }}
  </style>
</head>
<body>
  <header>
    <div class="brand">
      <img class="brand-logo" src="/assets/{LOGO_FILENAME}" alt="Preço Popular">
      <h1>Gerador de Orçamentos CLAMED</h1>
    </div>
    {nav_html}
  </header>
  <main>{content}</main>
</body>
</html>"""
    return page.encode("utf-8")


def form_value(data: dict[str, list[str]] | None, key: str, default: str = "") -> str:
    if data is None:
        return default
    return first(data, key, default)


def form_list(data: dict[str, list[str]] | None, key: str, default: list[str]) -> list[str]:
    if data is None:
        return default
    return data.get(key, default)


def field_class(field_error: str | None, key: str) -> str:
    return ' class="invalid"' if field_error == key else ""


def field_message(field_error: str | None, key: str, message: str) -> str:
    if field_error != key:
        return ""
    return f'<div class="field-error">{esc(message)}</div>'


def setup_page(error: str = "") -> bytes:
    error_html = f'<section class="error">{esc(error)}</section>' if error else ""
    content = f"""
<div class="auth-shell">
  <section class="auth-card">
    <img class="brand-logo" src="/assets/{LOGO_FILENAME}" alt="Preço Popular">
    <h2>Criar acesso administrador</h2>
    <p class="muted">Cadastre o primeiro usuário para proteger o app.</p>
    {error_html}
    <form method="post" action="/setup" autocomplete="off">
      <div>
        <label for="username">Usuário</label>
        <input id="username" name="username" autocomplete="off" autocapitalize="off" autocorrect="off" spellcheck="false" required>
      </div>
      <div>
        <label for="password">Senha</label>
        <input id="password" name="password" type="password" autocomplete="new-password" required>
      </div>
      <div>
        <label for="password_confirm">Confirmar senha</label>
        <input id="password_confirm" name="password_confirm" type="password" autocomplete="new-password" required>
      </div>
      <button class="primary" type="submit">Criar acesso</button>
    </form>
  </section>
</div>
"""
    return layout("Criar acesso", content, show_nav=False)


def login_page(error: str = "") -> bytes:
    error_html = f'<section class="error">{esc(error)}</section>' if error else ""
    content = f"""
<div class="auth-shell">
  <section class="auth-card">
    <img class="brand-logo" src="/assets/{LOGO_FILENAME}" alt="Preço Popular">
    <h2>Entrar no app</h2>
    {error_html}
    <form method="post" action="/login" autocomplete="off">
      <div>
        <label for="username">Usuário</label>
        <input id="username" name="username" autocomplete="off" autocapitalize="off" autocorrect="off" spellcheck="false" required>
      </div>
      <div>
        <label for="password">Senha</label>
        <input id="password" name="password" type="password" autocomplete="current-password" required>
      </div>
      <button class="primary" type="submit">Entrar</button>
    </form>
  </section>
</div>
"""
    return layout("Login", content, show_nav=False)


def farmaceutico_select(data: dict[str, list[str]] | None, field_error: str | None = None) -> str:
    selected = form_value(data, "farmaceutico_responsavel")
    options = [
        f'<option value=""{" selected" if not selected else ""}>Não informar</option>'
    ]
    found_selected = not selected
    for row in list_farmaceuticos():
        nome = row["nome"]
        is_selected = nome == selected
        found_selected = found_selected or is_selected
        options.append(
            f'<option value="{esc(nome)}"{" selected" if is_selected else ""}>{esc(nome)}</option>'
        )
    if selected and not found_selected:
        options.append(f'<option value="{esc(selected)}" selected>{esc(selected)}</option>')
    css_class = "form-select invalid" if field_error == "farmaceutico_responsavel" else "form-select"
    return f"""
        <select id="farmaceutico_responsavel" name="farmaceutico_responsavel" class="{css_class}" aria-label="Farmacêutico(a) responsável">
          {''.join(options)}
        </select>
        {field_message(field_error, 'farmaceutico_responsavel', 'Selecione um responsável cadastrado.')}
    """


def product_rows_from_form(data: dict[str, list[str]] | None) -> str:
    defaults = {
        "codigo": [""],
        "produto": [""],
        "pmc": [""],
        "desconto": [""],
        "valor_unitario": [""],
        "quantidade": ["1"],
    }
    columns = {key: form_list(data, key, value) for key, value in defaults.items()}
    row_count = max(len(values) for values in columns.values())
    rows = []
    for idx in range(row_count):
        codigo = columns["codigo"][idx] if idx < len(columns["codigo"]) else ""
        produto = columns["produto"][idx] if idx < len(columns["produto"]) else ""
        pmc = columns["pmc"][idx] if idx < len(columns["pmc"]) else ""
        desconto = columns["desconto"][idx] if idx < len(columns["desconto"]) else ""
        valor_unitario = columns["valor_unitario"][idx] if idx < len(columns["valor_unitario"]) else ""
        quantidade = columns["quantidade"][idx] if idx < len(columns["quantidade"]) else "1"
        rows.append(
            f"""
          <tr>
            <td><input name="codigo" value="{esc(codigo)}" autocomplete="off" autocapitalize="off" autocorrect="off" spellcheck="false"></td>
            <td><input name="produto" value="{esc(produto)}" autocomplete="off" autocapitalize="off" autocorrect="off" spellcheck="false" required></td>
            <td><input name="pmc" value="{esc(pmc)}" autocomplete="off" autocapitalize="off" autocorrect="off" spellcheck="false"></td>
            <td><input name="desconto" value="{esc(desconto)}" autocomplete="off" autocapitalize="off" autocorrect="off" spellcheck="false"></td>
            <td><input name="valor_unitario" value="{esc(valor_unitario)}" autocomplete="off" autocapitalize="off" autocorrect="off" spellcheck="false" required></td>
            <td><input name="quantidade" value="{esc(quantidade)}" autocomplete="off" autocapitalize="off" autocorrect="off" spellcheck="false" required></td>
            <td><button class="icon" type="button" onclick="removeRow(this)" title="Remover linha">×</button></td>
          </tr>
            """
        )
    return "".join(rows)


def form_page(
    error: str = "",
    data: dict[str, list[str]] | None = None,
    field_error: str | None = None,
) -> bytes:
    error_html = f'<section class="error">{esc(error)}</section>' if error else ""
    product_rows = product_rows_from_form(data)
    cpf_message = error if field_error == "cpf" else ""
    telefone_message = error if field_error == "telefone" else ""
    histories = list_orcamentos(8)
    rows = "".join(
        f"""
        <tr>
          <td>{row['id']}</td>
          <td>{esc(row['cliente_nome'])}</td>
          <td>{esc(row['farmaceutico_responsavel'])}</td>
          <td>{esc(date_for_display(row['data_orcamento']))}</td>
          <td>R$ {esc(money(Decimal(row['total'])))}</td>
          <td class="row-actions">
            <div class="action-group">
              <a class="button success" href="/orcamentos/{row['id']}">Abrir</a>
              <form method="post" action="/orcamentos/{row['id']}/excluir" onsubmit="return confirm('Excluir este orçamento e seus arquivos gerados?');">
                <button class="danger" type="submit">Excluir</button>
              </form>
            </div>
          </td>
        </tr>
        """
        for row in histories
    )
    if not rows:
        rows = '<tr><td colspan="6" class="muted">Nenhum orçamento salvo ainda.</td></tr>'

    content = f"""
{error_html}
<form method="post" action="/orcamentos" autocomplete="off">
  <section>
    <h2>Dados do cliente</h2>
    <div class="grid">
      <div class="span-2">
        <label for="cliente_nome">Cliente</label>
        <input id="cliente_nome" name="cliente_nome" value="{esc(form_value(data, 'cliente_nome'))}" autocomplete="off" autocapitalize="off" autocorrect="off" spellcheck="false" required>
      </div>
      <div>
        <label for="cpf">CPF</label>
        <input id="cpf" name="cpf" value="{esc(form_value(data, 'cpf'))}" inputmode="numeric" placeholder="000.000.000-00" autocomplete="off" autocapitalize="off" autocorrect="off" spellcheck="false"{field_class(field_error, 'cpf')}>
        {field_message(field_error, 'cpf', cpf_message)}
      </div>
      <div>
        <label for="telefone">Telefone</label>
        <input id="telefone" name="telefone" value="{esc(form_value(data, 'telefone'))}" inputmode="numeric" placeholder="(00) 00000-0000" autocomplete="off" autocapitalize="off" autocorrect="off" spellcheck="false"{field_class(field_error, 'telefone')}>
        {field_message(field_error, 'telefone', telefone_message)}
      </div>
      <div class="span-2">
        <label for="email">E-mail</label>
        <input id="email" name="email" value="{esc(form_value(data, 'email'))}" type="text" inputmode="email" autocomplete="off" autocapitalize="off" autocorrect="off" spellcheck="false">
      </div>
      <div>
        <label for="localidade">Local</label>
        <input id="localidade" name="localidade" value="{esc(form_value(data, 'localidade', 'Campo Grande'))}" autocomplete="off" autocapitalize="off" autocorrect="off" spellcheck="false">
      </div>
      <div>
        <label for="data_orcamento">Data</label>
        <input id="data_orcamento" name="data_orcamento" type="date" value="{esc(form_value(data, 'data_orcamento', today_iso()))}" autocomplete="off">
      </div>
    </div>
  </section>

  <section>
    <h2>Produtos</h2>
    <div class="table-wrap">
      <table id="produtos">
        <thead>
          <tr>
            <th>Código</th>
            <th>Produto</th>
            <th>PMC</th>
            <th>Desc. %</th>
            <th>Valor un.</th>
            <th>Qtd</th>
            <th></th>
          </tr>
        </thead>
        <tbody>
          {product_rows}
        </tbody>
      </table>
    </div>
    <div class="actions" style="margin-top: 14px;">
      <button class="secondary" type="button" onclick="addRow()">Adicionar produto</button>
    </div>
  </section>

  <section>
    <div class="actions final-actions">
      <div class="responsavel-select">
        <label for="farmaceutico_responsavel">Farm. resp.</label>
        {farmaceutico_select(data, field_error)}
      </div>
      <button class="primary" type="submit">Salvar e gerar arquivos</button>
    </div>
  </section>
</form>

<section>
  <h2>Últimos orçamentos</h2>
  <div class="table-wrap">
    <table class="history recent-history">
      <thead><tr><th>ID</th><th>Cliente</th><th>Farm. resp.</th><th>Data</th><th>Total</th><th></th></tr></thead>
      <tbody>{rows}</tbody>
    </table>
  </div>
</section>

<script>
function addRow() {{
  const tbody = document.querySelector("#produtos tbody");
  const first = tbody.querySelector("tr");
  const row = first.cloneNode(true);
  row.querySelectorAll("input").forEach((input) => {{
    input.value = input.name === "quantidade" ? "1" : "";
  }});
  tbody.appendChild(row);
}}
function removeRow(button) {{
  const tbody = document.querySelector("#produtos tbody");
  if (tbody.querySelectorAll("tr").length === 1) {{
    tbody.querySelectorAll("input").forEach((input) => {{
      input.value = input.name === "quantidade" ? "1" : "";
    }});
    return;
  }}
  button.closest("tr").remove();
}}
document.querySelectorAll("form, input, select").forEach((element) => {{
  element.setAttribute("autocomplete", "off");
}});
</script>
"""
    return layout("Novo orçamento", content)


def detail_page(orcamento_id: int) -> bytes:
    row = get_orcamento(orcamento_id)
    if row is None:
        return layout("Não encontrado", '<section class="error">Orçamento não encontrado.</section>')

    itens = get_itens(orcamento_id)
    item_rows = "".join(
        f"""
        <tr>
          <td>{esc(item['codigo'])}</td>
          <td>{esc(item['produto'])}</td>
          <td>{esc(item['pmc'])}</td>
          <td>{esc(item['desconto'])}</td>
          <td>R$ {esc(money(Decimal(item['valor_unitario'])))}</td>
          <td>{esc(item['quantidade'])}</td>
          <td>R$ {esc(money(Decimal(item['total'])))}</td>
        </tr>
        """
        for item in itens
    )
    pdf_exists = bool(row["pdf_path"]) and Path(row["pdf_path"]).exists()
    if pdf_exists:
        pdf_controls = f"""
    <a class="button secondary" href="/download/{orcamento_id}/pdf" target="_blank" rel="noopener">Abrir PDF</a>
    <form method="post" action="/orcamentos/{orcamento_id}/gerar-pdf">
      <button class="secondary" type="submit">Gerar PDF novamente</button>
    </form>
        """
    else:
        pdf_controls = f"""
    <span class="warn">{esc(row["pdf_status"])}</span>
    <form method="post" action="/orcamentos/{orcamento_id}/gerar-pdf">
      <button class="secondary" type="submit">Gerar PDF</button>
    </form>
        """
    content = f"""
<section>
  <h2>Orçamento #{orcamento_id}</h2>
  <div class="grid">
    <div class="span-2"><label>Cliente</label><div>{esc(row['cliente_nome'])}</div></div>
    <div class="span-2"><label>Farm. resp.</label><div>{esc(row['farmaceutico_responsavel'])}</div></div>
    <div><label>CPF</label><div>{esc(format_cpf(row['cpf']))}</div></div>
    <div><label>Telefone</label><div>{esc(format_cell_phone(row['telefone']))}</div></div>
    <div class="span-2"><label>E-mail</label><div>{esc(row['email'])}</div></div>
    <div><label>Data</label><div>{esc(date_for_display(row['data_orcamento']))}</div></div>
    <div><label>Total</label><div class="total-box">R$ {esc(money(Decimal(row['total'])))}</div></div>
  </div>
</section>
<section>
  <h2>Produtos</h2>
  <div class="table-wrap">
    <table>
      <thead>
        <tr><th>Código</th><th>Produto</th><th>PMC</th><th>Desc.</th><th>Valor un.</th><th>Qtd</th><th>Total</th></tr>
      </thead>
      <tbody>{item_rows}</tbody>
    </table>
  </div>
</section>
<section>
  <div class="actions">
    <a class="button primary" href="/download/{orcamento_id}/docx">Baixar DOCX</a>
    {pdf_controls}
    <form method="post" action="/orcamentos/{orcamento_id}/excluir" onsubmit="return confirm('Excluir este orçamento e seus arquivos gerados?');">
      <button class="danger" type="submit">Excluir orçamento</button>
    </form>
  </div>
</section>
"""
    return layout(f"Orçamento #{orcamento_id}", content)


def history_page() -> bytes:
    rows = "".join(
        f"""
        <tr>
          <td>{row['id']}</td>
          <td>{esc(row['criado_em'])}</td>
          <td>{esc(row['cliente_nome'])}</td>
          <td>{esc(row['farmaceutico_responsavel'])}</td>
          <td>{esc(format_cpf(row['cpf']))}</td>
          <td>{esc(date_for_display(row['data_orcamento']))}</td>
          <td>R$ {esc(money(Decimal(row['total'])))}</td>
          <td class="row-actions">
            <div class="action-group">
              <a class="button success" href="/orcamentos/{row['id']}">Abrir</a>
              <form method="post" action="/orcamentos/{row['id']}/excluir" onsubmit="return confirm('Excluir este orçamento e seus arquivos gerados?');">
                <button class="danger" type="submit">Excluir</button>
              </form>
            </div>
          </td>
        </tr>
        """
        for row in list_orcamentos(100)
    )
    if not rows:
        rows = '<tr><td colspan="8" class="muted">Nenhum orçamento salvo ainda.</td></tr>'
    content = f"""
<section>
  <h2>Histórico</h2>
  <div class="table-wrap">
    <table class="history full-history">
      <thead>
        <tr><th>ID</th><th>Criado em</th><th>Cliente</th><th>Farm. resp.</th><th>CPF</th><th>Data</th><th>Total</th><th></th></tr>
      </thead>
      <tbody>{rows}</tbody>
    </table>
  </div>
</section>
"""
    return layout("Histórico", content)


def farmaceuticos_page(error: str = "") -> bytes:
    error_html = f'<section class="error">{esc(error)}</section>' if error else ""
    rows = "".join(
        f"""
        <tr>
          <td>
            <div class="manager-row compact">
              <input id="farmaceutico_{row['id']}" value="{esc(row['nome'])}" readonly>
            </div>
          </td>
          <td class="row-actions">
            <form method="post" action="/farmaceuticos/{row['id']}/excluir" onsubmit="return confirm('Excluir este responsável da lista? Os orçamentos antigos não serão alterados.');">
              <button class="danger" type="submit">Excluir</button>
            </form>
          </td>
        </tr>
        """
        for row in list_farmaceuticos(include_inactive=True)
    )
    if not rows:
        rows = '<tr><td colspan="2" class="muted">Nenhum farmacêutico(a) cadastrado ainda.</td></tr>'
    content = f"""
{error_html}
<section>
  <h2>Novo responsável</h2>
  <form method="post" action="/farmaceuticos" autocomplete="off">
    <div class="manager-row create-row">
      <div>
        <label for="novo_farmaceutico">Farmacêutico(a)</label>
        <input id="novo_farmaceutico" name="nome" autocomplete="off" autocapitalize="off" autocorrect="off" spellcheck="false" required>
      </div>
      <button class="primary" type="submit">Adicionar</button>
    </div>
  </form>
</section>
<section>
  <h2>Responsáveis cadastrados</h2>
  <div class="table-wrap">
    <table class="history">
      <thead><tr><th>Farmacêutico(a)</th><th></th></tr></thead>
      <tbody>{rows}</tbody>
    </table>
  </div>
</section>
<section>
  <div class="actions">
    <a class="button secondary" href="/">Voltar ao formulário</a>
  </div>
</section>
"""
    return layout("Farmacêuticos", content)


def template_validation_page() -> bytes:
    ok, messages = validate_template()
    status = "ok" if ok else "warn"
    title = "Modelo pronto para uso" if ok else "Modelo precisa de ajuste"
    rows = "".join(
        f'<tr><td class="{esc("ok" if message.startswith("OK:") else "warn")}">{esc(message)}</td></tr>'
        for message in messages
    )
    content = f"""
<section>
  <h2>Validação do modelo</h2>
  <p class="{status}">{title}</p>
  <div class="table-wrap">
    <table class="history">
      <tbody>{rows}</tbody>
    </table>
  </div>
</section>
<section>
  <div class="actions">
    <a class="button secondary" href="/">Voltar ao formulário</a>
  </div>
</section>
"""
    return layout("Validar modelo", content)


def month_select(data: dict[str, list[str]] | None = None) -> str:
    current_month = str(int(today_iso()[5:7]))
    selected = form_value(data, "mes", current_month)
    options = []
    for value, label in MONTH_OPTIONS:
        value_raw = str(value)
        options.append(
            f'<option value="{value_raw}"{" selected" if selected == value_raw else ""}>{esc(label)}</option>'
        )
    return "".join(options)


def extra_days_calendar(data: dict[str, list[str]] | None = None) -> str:
    selected = set(form_list(data, "dias_extras", []))
    try:
        current_month = int(form_value(data, "mes", str(int(today_iso()[5:7]))) or today_iso()[5:7])
        current_year = int(form_value(data, "ano", today_iso()[:4]) or today_iso()[:4])
        days_in_month = monthrange(current_year, current_month)[1]
        first_weekday = date(current_year, current_month, 1).weekday()
        leading_blanks = (first_weekday + 1) % 7
        sunday_days = {
            day
            for day in range(1, days_in_month + 1)
            if date(current_year, current_month, day).weekday() == 6
        }
    except Exception:
        current_month = int(today_iso()[5:7])
        current_year = int(today_iso()[:4])
        days_in_month = 31
        leading_blanks = 0
        sunday_days = set()
    days_html = []
    for _ in range(leading_blanks):
        days_html.append('<div class="day-check blank"><span></span></div>')
    for day in range(1, 32):
        if day > days_in_month:
            break
        value = str(day)
        classes = ["day-check"]
        if day in sunday_days:
            classes.append("sunday")
        if date(current_year, current_month, day).weekday() == 5:
            classes.append("saturday")
        days_html.append(
            f"""
        <label class="{' '.join(classes)}">
          <input type="checkbox" name="dias_extras" value="{value}"{" checked" if value in selected else ""}>
          <span>{day}</span>
        </label>
            """
        )
    trailing_blanks = (7 - (len(days_html) % 7)) % 7
    for _ in range(trailing_blanks):
        days_html.append('<div class="day-check blank"><span></span></div>')
    month_label = MONTH_OPTIONS[current_month - 1][1].lower() if 1 <= current_month <= 12 else ""
    return f"""
      <div class="calendar-picker span-4">
        <div class="actions calendar-toggle">
          <button class="secondary" id="calendar-toggle-button" type="button" aria-expanded="false" aria-controls="extra-days-calendar">
            Selecionar dias
          </button>
        </div>
        <div class="calendar-collapse" id="extra-days-calendar" hidden>
          <div class="calendar-shell">
            <div class="calendar-title" id="calendar-title">{esc(month_label)} {current_year}</div>
            <div class="calendar-weekdays">
              <span>domingo</span>
              <span>segunda-feira</span>
              <span>terça-feira</span>
              <span>quarta-feira</span>
              <span>quinta-feira</span>
              <span>sexta-feira</span>
              <span>sábado</span>
            </div>
            <div class="calendar-grid" id="calendar-grid">
              {''.join(days_html)}
            </div>
          </div>
          <div class="calendar-legend">
            <span>Domingos já são marcados automaticamente.</span>
            <span>Dias selecionados também recebem ****.</span>
          </div>
        </div>
      </div>
      <script>
      (function () {{
        const month = document.getElementById("mes");
        const year = document.getElementById("ano");
        const grid = document.getElementById("calendar-grid");
        const title = document.getElementById("calendar-title");
        const toggle = document.getElementById("calendar-toggle-button");
        const panel = document.getElementById("extra-days-calendar");
        const selectedDays = new Set(
          Array.from(document.querySelectorAll('input[name="dias_extras"]:checked')).map((input) => input.value)
        );
        const monthNames = {str([label.lower() for _, label in MONTH_OPTIONS])};
        const weekdayNames = [
          "domingo",
          "segunda-feira",
          "terça-feira",
          "quarta-feira",
          "quinta-feira",
          "sexta-feira",
          "sábado",
        ];
        function dayCell(day, weekday) {{
          const label = document.createElement("label");
          label.className = "day-check";
          if (weekday === 0) label.classList.add("sunday");
          if (weekday === 6) label.classList.add("saturday");
          const input = document.createElement("input");
          input.type = "checkbox";
          input.name = "dias_extras";
          input.value = String(day);
          input.checked = selectedDays.has(String(day));
          input.addEventListener("change", () => {{
            if (input.checked) selectedDays.add(input.value);
            else selectedDays.delete(input.value);
          }});
          const span = document.createElement("span");
          span.textContent = String(day);
          label.appendChild(input);
          label.appendChild(span);
          return label;
        }}
        function blankCell() {{
          const cell = document.createElement("div");
          cell.className = "day-check blank";
          cell.innerHTML = "<span></span>";
          return cell;
        }}
        function updateCalendar() {{
          const mes = Number(month.value);
          const ano = Number(year.value);
          if (!mes || !ano) return;
          const daysInMonth = new Date(ano, mes, 0).getDate();
          title.textContent = `${{monthNames[mes - 1]}} ${{ano}}`;
          grid.replaceChildren();
          const firstDay = new Date(ano, mes - 1, 1).getDay();
          for (let index = 0; index < firstDay; index += 1) grid.appendChild(blankCell());
          for (let day = 1; day <= daysInMonth; day += 1) {{
            const weekday = new Date(ano, mes - 1, day).getDay();
            grid.appendChild(dayCell(day, weekday));
          }}
          while (grid.children.length % 7 !== 0) grid.appendChild(blankCell());
        }}
        toggle.addEventListener("click", () => {{
          const shouldOpen = !panel.classList.contains("open");
          if (shouldOpen) {{
            panel.hidden = false;
            requestAnimationFrame(() => panel.classList.add("open"));
            updateCalendar();
          }} else {{
            panel.classList.remove("open");
            setTimeout(() => {{
              if (!panel.classList.contains("open")) panel.hidden = true;
            }}, 240);
          }}
          toggle.setAttribute("aria-expanded", String(shouldOpen));
        }});
        month.addEventListener("change", updateCalendar);
        year.addEventListener("input", updateCalendar);
      }})();
      </script>
    """


def temperature_map_form_page(
    error: str = "",
    data: dict[str, list[str]] | None = None,
    field_error: str | None = None,
    generated_filename: str = "",
    pdf_status: str = "",
) -> bytes:
    error_html = f'<section class="error">{esc(error)}</section>' if error else ""
    current_year = today_iso()[:4]
    generated_html = temperature_maps_section(generated_filename, pdf_status)
    content = f"""
{error_html}
<form method="post" action="/mapa-temperatura" autocomplete="off">
  <section>
    <h2>Mapa de temperatura</h2>
    <div class="grid">
      <div>
        <label for="mes">Mês</label>
        <select id="mes" name="mes" autocomplete="off"{field_class(field_error, 'mes')}>
          {month_select(data)}
        </select>
      </div>
      <div>
        <label for="ano">Ano</label>
        <input id="ano" name="ano" value="{esc(form_value(data, 'ano', current_year))}" inputmode="numeric" autocomplete="off" autocapitalize="off" autocorrect="off" spellcheck="false"{field_class(field_error, 'ano')} required>
      </div>
      <div>
        <label for="filial">Filial</label>
        <input id="filial" name="filial" value="{esc(form_value(data, 'filial', '386'))}" inputmode="numeric" placeholder="000" maxlength="3" autocomplete="off" autocapitalize="off" autocorrect="off" spellcheck="false"{field_class(field_error, 'filial')} required>
      </div>
      {extra_days_calendar(data)}
    </div>
    <div class="actions" style="margin-top: 14px;">
      <button class="primary" type="submit">Gerar mapa</button>
    </div>
  </section>
</form>
{generated_html}
"""
    return layout("Mapa de temperatura", content)


def temperature_maps_section(generated_filename: str = "", pdf_status: str = "") -> str:
    maps = list_temperature_maps()
    if not maps:
        return ""
    rows = []
    for xlsx_path in maps:
        pdf_path = temperature_output_file(f"{xlsx_path.stem}.pdf")
        pdf_html = ""
        if pdf_path is not None:
            pdf_html = (
                f'<a class="button secondary" href="/mapa-temperatura/download/{quote(pdf_path.name)}" '
                'target="_blank" rel="noopener">Abrir PDF</a>'
            )
        elif xlsx_path.name == generated_filename and pdf_status:
            pdf_html = (
                f'<button class="pdf-status" type="button" aria-label="{esc(pdf_status)}">'
                f'!<span class="popover">{esc(pdf_status)}</span></button>'
            )
        else:
            pdf_html = '<span class="muted">PDF indisponível</span>'
        row_class = ' class="highlight-row"' if xlsx_path.name == generated_filename else ""
        rows.append(
            f"""
        <tr{row_class}>
          <td>{esc(xlsx_path.name)}</td>
          <td class="row-actions">
            <div class="action-group">
              <a class="button primary" href="/mapa-temperatura/download/{quote(xlsx_path.name)}">Baixar XLSX</a>
              {pdf_html}
              <form method="post" action="/mapa-temperatura/excluir/{quote(xlsx_path.name)}" onsubmit="return confirm('Excluir este mapa gerado?');">
                <button class="danger" type="submit">Excluir</button>
              </form>
            </div>
          </td>
        </tr>
            """
        )
    return f"""
<section>
  <h2 class="success-title">Mapas gerados</h2>
  <div class="table-wrap">
    <table class="history temperature-history">
      <thead><tr><th></th><th></th></tr></thead>
      <tbody>{''.join(rows)}</tbody>
    </table>
  </div>
</section>
"""


def temperature_map_result_page(filename: str, pdf_status: str = "") -> bytes:
    content = temperature_maps_section(filename, pdf_status)
    return layout("Mapa gerado", content)


class Handler(BaseHTTPRequestHandler):
    def request_session_token(self) -> str | None:
        raw_cookie = self.headers.get("Cookie")
        if not raw_cookie:
            return None
        cookie = SimpleCookie()
        cookie.load(raw_cookie)
        morsel = cookie.get(SESSION_COOKIE)
        return morsel.value if morsel else None

    def current_user(self):
        return get_user_by_session(self.request_session_token())

    def is_https(self) -> bool:
        return self.headers.get("X-Forwarded-Proto", "").lower() == "https"

    def session_cookie(self, token: str) -> str:
        parts = [
            f"{SESSION_COOKIE}={token}",
            "Path=/",
            f"Max-Age={SESSION_MAX_AGE}",
            "HttpOnly",
            "SameSite=Lax",
        ]
        if self.is_https():
            parts.append("Secure")
        return "; ".join(parts)

    def clear_session_cookie(self) -> str:
        parts = [
            f"{SESSION_COOKIE}=",
            "Path=/",
            "Max-Age=0",
            "HttpOnly",
            "SameSite=Lax",
        ]
        if self.is_https():
            parts.append("Secure")
        return "; ".join(parts)

    def respond(self, status: int, body: bytes, content_type: str = "text/html; charset=utf-8") -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("X-Frame-Options", "DENY")
        self.send_header("Referrer-Policy", "same-origin")
        if content_type.startswith("text/html"):
            self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def redirect(self, path: str, extra_headers: dict[str, str] | None = None) -> None:
        self.send_response(303)
        self.send_header("Location", path)
        self.send_header("Cache-Control", "no-store")
        if extra_headers:
            for key, value in extra_headers.items():
                self.send_header(key, value)
        self.end_headers()

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        path = parsed.path
        if path == "/healthz":
            return self.respond(200, b"ok", "text/plain; charset=utf-8")
        if path == f"/assets/{LOGO_FILENAME}":
            logo_path = ASSETS_DIR / LOGO_FILENAME
            if logo_path.exists():
                return self.send_file(logo_path, "image/svg+xml", disposition="inline")
            return self.respond(404, layout("Erro", '<section class="error">Logo não encontrada.</section>'))
        if path == "/setup":
            if has_users():
                return self.redirect("/login")
            return self.respond(200, setup_page())
        if path == "/login":
            if not has_users():
                return self.redirect("/setup")
            if self.current_user() is not None:
                return self.redirect("/")
            return self.respond(200, login_page())
        if path == "/favicon.ico":
            return self.respond(404, b"", "text/plain")

        if not has_users():
            return self.redirect("/setup")
        if self.current_user() is None:
            return self.redirect("/login")

        if path == "/":
            return self.respond(200, form_page())
        if path == "/mapa-temperatura":
            query = parse_qs(parsed.query, keep_blank_values=True)
            return self.respond(
                200,
                temperature_map_form_page(
                    generated_filename=first(query, "arquivo"),
                    pdf_status=first(query, "pdf_status"),
                ),
            )
        if path == "/mapa-temperatura/resultado":
            query = parse_qs(parsed.query, keep_blank_values=True)
            filename = first(query, "arquivo")
            pdf_status = first(query, "pdf_status")
            return self.respond(200, temperature_map_result_page(filename, pdf_status))
        if path.startswith("/mapa-temperatura/download/"):
            return self.handle_temperature_download(path)
        if path == "/historico":
            return self.respond(200, history_page())
        if path == "/farmaceuticos":
            return self.respond(200, farmaceuticos_page())
        if path == "/validar-modelo":
            return self.respond(200, template_validation_page())
        if path == "/exportar":
            csv_path = export_history_csv()
            return self.send_file(csv_path, "text/csv; charset=utf-8")
        if path.startswith("/orcamentos/"):
            try:
                orcamento_id = int(path.rsplit("/", 1)[-1])
            except ValueError:
                return self.respond(404, layout("Erro", '<section class="error">Endereço inválido.</section>'))
            return self.respond(200, detail_page(orcamento_id))
        if path.startswith("/download/"):
            return self.handle_download(path)
        return self.respond(404, layout("Erro", '<section class="error">Página não encontrada.</section>'))

    def do_POST(self) -> None:
        path = urlparse(self.path).path
        if path == "/setup":
            return self.handle_setup()
        if path == "/login":
            return self.handle_login()
        if path == "/logout":
            token = self.request_session_token()
            delete_session(token)
            return self.redirect("/login", {"Set-Cookie": self.clear_session_cookie()})

        if not has_users():
            return self.redirect("/setup")
        if self.current_user() is None:
            return self.redirect("/login")

        if path == "/farmaceuticos":
            return self.handle_add_farmaceutico()
        if path == "/mapa-temperatura":
            return self.handle_temperature_map()
        if path.startswith("/mapa-temperatura/excluir/"):
            return self.handle_temperature_delete(path)
        if path.startswith("/farmaceuticos/") and path.endswith("/excluir"):
            return self.handle_delete_farmaceutico(path)
        if path.startswith("/orcamentos/") and path.endswith("/gerar-pdf"):
            return self.handle_generate_pdf(path)
        if path.startswith("/orcamentos/") and path.endswith("/excluir"):
            return self.handle_delete(path)
        if path != "/orcamentos":
            return self.respond(404, layout("Erro", '<section class="error">Página não encontrada.</section>'))

        length = int(self.headers.get("Content-Length", "0"))
        data = parse_form(self.rfile.read(length))
        try:
            try:
                cpf = require_valid_cpf(first(data, "cpf"))
            except ValueError as exc:
                return self.respond(400, form_page(f"Campo CPF: {exc}", data, "cpf"))
            try:
                telefone = require_valid_cell_phone(first(data, "telefone"))
            except ValueError as exc:
                return self.respond(400, form_page(f"Campo Telefone: {exc}", data, "telefone"))

            orcamento = Orcamento(
                cliente_nome=first(data, "cliente_nome"),
                farmaceutico_responsavel=first(data, "farmaceutico_responsavel"),
                cpf=cpf,
                telefone=telefone,
                email=first(data, "email") or None,
                data_orcamento=first(data, "data_orcamento") or today_iso(),
                localidade=first(data, "localidade") or "Campo Grande",
                itens=build_items(data),
            )
            if not orcamento.cliente_nome:
                raise ValueError("Informe o nome do cliente.")
            farmaceuticos_ativos = {row["nome"] for row in list_farmaceuticos()}
            if (
                orcamento.farmaceutico_responsavel
                and orcamento.farmaceutico_responsavel not in farmaceuticos_ativos
            ):
                return self.respond(
                    400,
                    form_page(
                        "Campo Farm. resp.: selecione um responsável cadastrado.",
                        data,
                        "farmaceutico_responsavel",
                    ),
                )
            draft_path = render_orcamento(orcamento)
            pdf_path = None
            pdf_status = "PDF pendente: aguardando geração do arquivo final."
            orcamento_id = save_orcamento(orcamento, draft_path, pdf_path, pdf_status)
            final_docx = render_orcamento(orcamento, sequence=orcamento_id)
            if final_docx != draft_path:
                draft_path.unlink(missing_ok=True)
            pdf_path, pdf_status = convert_to_pdf(final_docx)
            with connect() as conn:
                conn.execute(
                    "UPDATE orcamentos SET docx_path = ?, pdf_path = ?, pdf_status = ? WHERE id = ?",
                    (str(final_docx), str(pdf_path) if pdf_path else None, pdf_status, orcamento_id),
            )
            return self.redirect(f"/orcamentos/{orcamento_id}")
        except Exception as exc:
            return self.respond(400, form_page(str(exc), data))

    def handle_setup(self) -> None:
        if has_users():
            return self.redirect("/login")
        length = int(self.headers.get("Content-Length", "0"))
        data = parse_form(self.rfile.read(length))
        username = first(data, "username")
        password = first(data, "password")
        password_confirm = first(data, "password_confirm")
        try:
            if password != password_confirm:
                raise ValueError("As senhas não conferem.")
            user_id = create_user(username, password)
            token = create_session(user_id)
        except Exception as exc:
            return self.respond(400, setup_page(str(exc)))
        return self.redirect("/", {"Set-Cookie": self.session_cookie(token)})

    def handle_login(self) -> None:
        if not has_users():
            return self.redirect("/setup")
        length = int(self.headers.get("Content-Length", "0"))
        data = parse_form(self.rfile.read(length))
        user = verify_login(first(data, "username"), first(data, "password"))
        if user is None:
            return self.respond(401, login_page("Usuário ou senha inválidos."))
        token = create_session(int(user["id"]))
        return self.redirect("/", {"Set-Cookie": self.session_cookie(token)})

    def handle_temperature_map(self) -> None:
        length = int(self.headers.get("Content-Length", "0"))
        data = parse_form(self.rfile.read(length))
        try:
            try:
                mapa = parse_temperature_map_input(
                    first(data, "mes"),
                    first(data, "ano"),
                    first(data, "filial"),
                    data.get("dias_extras", []),
                )
            except ValueError as exc:
                message = str(exc)
                field_error = None
                if message.startswith("Campo Mês:"):
                    field_error = "mes"
                elif message.startswith("Campo Ano:"):
                    field_error = "ano"
                elif message.startswith("Campo Filial:"):
                    field_error = "filial"
                return self.respond(400, temperature_map_form_page(message, data, field_error))

            xlsx_path = render_temperature_map(mapa)
            _pdf_path, pdf_status = convert_temperature_map_to_pdf(xlsx_path)
            return self.redirect(
                "/mapa-temperatura"
                f"?arquivo={quote(xlsx_path.name)}&pdf_status={quote(pdf_status)}"
            )
        except Exception as exc:
            return self.respond(400, temperature_map_form_page(str(exc), data))

    def handle_add_farmaceutico(self) -> None:
        length = int(self.headers.get("Content-Length", "0"))
        data = parse_form(self.rfile.read(length))
        try:
            add_farmaceutico(first(data, "nome"))
        except Exception as exc:
            return self.respond(400, farmaceuticos_page(str(exc)))
        return self.redirect("/farmaceuticos")

    def handle_delete_farmaceutico(self, path: str) -> None:
        parts = path.strip("/").split("/")
        if len(parts) != 3:
            return self.respond(404, layout("Erro", '<section class="error">Exclusão inválida.</section>'))
        _, id_raw, action = parts
        if action != "excluir":
            return self.respond(404, layout("Erro", '<section class="error">Exclusão inválida.</section>'))
        try:
            farmaceutico_id = int(id_raw)
        except ValueError:
            return self.respond(404, layout("Erro", '<section class="error">Responsável inválido.</section>'))
        delete_farmaceutico(farmaceutico_id)
        return self.redirect("/farmaceuticos")

    def handle_generate_pdf(self, path: str) -> None:
        parts = path.strip("/").split("/")
        if len(parts) != 3:
            return self.respond(404, layout("Erro", '<section class="error">Geração de PDF inválida.</section>'))
        _, id_raw, action = parts
        if action != "gerar-pdf":
            return self.respond(404, layout("Erro", '<section class="error">Geração de PDF inválida.</section>'))
        try:
            orcamento_id = int(id_raw)
        except ValueError:
            return self.respond(404, layout("Erro", '<section class="error">Orçamento inválido.</section>'))

        row = get_orcamento(orcamento_id)
        if row is None:
            return self.respond(404, layout("Erro", '<section class="error">Orçamento não encontrado.</section>'))

        try:
            docx_path = ensure_docx_file(orcamento_id, row)
        except Exception as exc:
            with connect() as conn:
                conn.execute(
                    "UPDATE orcamentos SET pdf_path = NULL, pdf_status = ? WHERE id = ?",
                    (f"PDF não gerado: não foi possível recriar o DOCX. {exc}", orcamento_id),
                )
            return self.redirect(f"/orcamentos/{orcamento_id}")

        pdf_path, pdf_status = convert_to_pdf(docx_path)
        update_orcamento_files(orcamento_id, pdf_path=pdf_path, pdf_status=pdf_status)
        return self.redirect(f"/orcamentos/{orcamento_id}")

    def handle_delete(self, path: str) -> None:
        parts = path.strip("/").split("/")
        if len(parts) != 3:
            return self.respond(404, layout("Erro", '<section class="error">Exclusão inválida.</section>'))
        _, id_raw, action = parts
        if action != "excluir":
            return self.respond(404, layout("Erro", '<section class="error">Exclusão inválida.</section>'))
        try:
            orcamento_id = int(id_raw)
        except ValueError:
            return self.respond(404, layout("Erro", '<section class="error">Orçamento inválido.</section>'))

        row = delete_orcamento(orcamento_id)
        if row is not None:
            for key in ("docx_path", "pdf_path"):
                value = row[key]
                if value:
                    Path(value).unlink(missing_ok=True)
        return self.redirect("/historico")

    def handle_temperature_download(self, path: str) -> None:
        filename = unquote(path.rsplit("/", 1)[-1])
        file_path = temperature_output_file(filename)
        if file_path is None or file_path.suffix.lower() not in {".xlsx", ".pdf"}:
            return self.respond(404, layout("Erro", '<section class="error">Arquivo não encontrado.</section>'))
        content_type = mimetypes.guess_type(file_path.name)[0] or "application/octet-stream"
        disposition = "inline" if file_path.suffix.lower() == ".pdf" else "attachment"
        return self.send_file(file_path, content_type, disposition=disposition)

    def handle_temperature_delete(self, path: str) -> None:
        filename = unquote(path.rsplit("/", 1)[-1])
        file_path = temperature_output_file(filename)
        if file_path is None or file_path.suffix.lower() != ".xlsx":
            return self.respond(404, layout("Erro", '<section class="error">Mapa não encontrado.</section>'))
        pdf_path = temperature_output_file(f"{file_path.stem}.pdf")
        file_path.unlink(missing_ok=True)
        if pdf_path is not None:
            pdf_path.unlink(missing_ok=True)
        return self.redirect("/mapa-temperatura")

    def handle_download(self, path: str) -> None:
        parts = path.strip("/").split("/")
        if len(parts) != 3:
            return self.respond(404, layout("Erro", '<section class="error">Download inválido.</section>'))
        _, id_raw, kind = parts
        try:
            row = get_orcamento(int(id_raw))
        except ValueError:
            row = None
        if row is None or kind not in ("docx", "pdf"):
            return self.respond(404, layout("Erro", '<section class="error">Arquivo não encontrado.</section>'))
        orcamento_id = int(id_raw)
        if kind == "docx":
            try:
                file_path = ensure_docx_file(orcamento_id, row)
            except Exception as exc:
                return self.respond(
                    500,
                    layout("Erro", f'<section class="error">Não foi possível gerar o DOCX: {esc(exc)}</section>'),
                )
        else:
            file_path = existing_path(row["pdf_path"])
            if file_path is None:
                try:
                    docx_path = ensure_docx_file(orcamento_id, row)
                except Exception as exc:
                    return self.respond(
                        500,
                        layout("Erro", f'<section class="error">Não foi possível gerar o DOCX para criar o PDF: {esc(exc)}</section>'),
                    )
                pdf_path, pdf_status = convert_to_pdf(docx_path)
                update_orcamento_files(orcamento_id, pdf_path=pdf_path, pdf_status=pdf_status)
                if pdf_path is None:
                    return self.respond(404, layout("Erro", f'<section class="error">{esc(pdf_status)}</section>'))
                file_path = pdf_path
        content_type = mimetypes.guess_type(file_path.name)[0] or "application/octet-stream"
        disposition = "inline" if kind == "pdf" else "attachment"
        return self.send_file(file_path, content_type, disposition=disposition)

    def send_file(self, path: Path, content_type: str, disposition: str = "attachment") -> None:
        data = path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Referrer-Policy", "same-origin")
        if path.suffix.lower() in {".docx", ".xlsx", ".pdf", ".csv"}:
            self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Content-Disposition", f'{disposition}; filename="{path.name}"')
        self.end_headers()
        self.wfile.write(data)

    def log_message(self, format: str, *args) -> None:
        return


def run() -> None:
    init_db()
    preferred_port = int(os.environ.get("PORT", PORT))
    server = None
    selected_port = preferred_port
    candidates = [preferred_port] if os.environ.get("RENDER") else range(preferred_port, preferred_port + 20)
    for candidate in candidates:
        try:
            server = ThreadingHTTPServer((HOST, candidate), Handler)
            selected_port = candidate
            break
        except OSError:
            continue
    if server is None:
        raise SystemExit("Não foi possível encontrar uma porta livre para iniciar o servidor.")
    print(f"Servidor iniciado em http://{HOST}:{selected_port}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nServidor encerrado.")
