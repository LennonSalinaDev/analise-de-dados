from __future__ import annotations

import html
import os
import mimetypes
from decimal import Decimal
from pathlib import Path
from urllib.parse import parse_qs, urlparse
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
    connect,
    delete_orcamento,
    get_itens,
    get_orcamento,
    init_db,
    list_orcamentos,
    save_orcamento,
)


HOST = os.environ.get("HOST", "127.0.0.1")
PORT = 8000


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


def layout(title: str, content: str) -> bytes:
    page = f"""<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{esc(title)}</title>
  <style>
    :root {{
      --bg: #f6f7f9;
      --panel: #ffffff;
      --ink: #1d2430;
      --muted: #5f6b7a;
      --line: #d9dee7;
      --accent: #1967d2;
      --accent-dark: #0f4fa8;
      --ok: #157347;
      --warn: #9a5b00;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: Arial, Helvetica, sans-serif;
      background: var(--bg);
      color: var(--ink);
    }}
    header {{
      background: #ffffff;
      border-bottom: 1px solid var(--line);
      padding: 18px 28px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 16px;
    }}
    header h1 {{
      margin: 0;
      font-size: 20px;
      letter-spacing: 0;
    }}
    nav a {{
      color: var(--accent);
      text-decoration: none;
      font-weight: 700;
      margin-left: 18px;
    }}
    main {{
      width: min(1180px, calc(100vw - 32px));
      margin: 24px auto 48px;
    }}
    section {{
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 20px;
      margin-bottom: 18px;
    }}
    h2 {{
      font-size: 16px;
      margin: 0 0 16px;
    }}
    label {{
      display: block;
      font-size: 13px;
      color: var(--muted);
      margin-bottom: 6px;
      font-weight: 700;
    }}
    input {{
      width: 100%;
      height: 38px;
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 8px 10px;
      font-size: 14px;
      background: #fff;
      color: var(--ink);
    }}
    input:focus {{
      outline: 2px solid rgba(25, 103, 210, .2);
      border-color: var(--accent);
    }}
    input.invalid {{
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
    }}
    button.primary, .button.primary {{
      background: var(--accent);
      color: #fff;
    }}
    button.primary:hover, .button.primary:hover {{ background: var(--accent-dark); }}
    button.secondary, .button.secondary {{
      background: #e9edf5;
      color: var(--ink);
    }}
    button.danger, .button.danger {{
      background: #d24a3f;
      color: #fff;
    }}
    button.danger:hover, .button.danger:hover {{
      background: #bf382e;
    }}
    button.success, .button.success {{
      background: #2f9a60;
      color: #fff;
    }}
    button.success:hover, .button.success:hover {{
      background: #258750;
    }}
    button.icon {{
      width: 34px;
      min-height: 34px;
      padding: 0;
      background: #f0f2f6;
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
    .history th:last-child,
    .history td:last-child {{
      width: 128px;
      max-width: 128px;
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
    .error {{
      background: #fff4f4;
      border-color: #ffd0d0;
      color: #8a1f1f;
      font-weight: 700;
    }}
    .total-box {{
      font-size: 18px;
      font-weight: 800;
      text-align: right;
    }}
    @media (max-width: 860px) {{
      header {{ align-items: flex-start; flex-direction: column; }}
      nav a {{ margin-left: 0; margin-right: 12px; }}
      .grid {{ grid-template-columns: 1fr; }}
      .span-2, .span-4 {{ grid-column: span 1; }}
      table {{ min-width: 850px; }}
      .table-wrap {{ overflow-x: auto; }}
    }}
  </style>
</head>
<body>
  <header>
    <h1>Gerador de Orçamentos CLAMED</h1>
    <nav>
      <a href="/">Novo orçamento</a>
      <a href="/historico">Histórico</a>
      <a href="/validar-modelo">Validar modelo</a>
      <a href="/exportar">Exportar CSV</a>
    </nav>
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
      <div class="span-2">
        <label for="farmaceutico_responsavel">Farmacêutico(a) responsável</label>
        <input id="farmaceutico_responsavel" name="farmaceutico_responsavel" value="{esc(form_value(data, 'farmaceutico_responsavel'))}" autocomplete="off" autocapitalize="off" autocorrect="off" spellcheck="false" required>
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
    <div class="actions">
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
document.querySelectorAll("form, input").forEach((element) => {{
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


class Handler(BaseHTTPRequestHandler):
    def respond(self, status: int, body: bytes, content_type: str = "text/html; charset=utf-8") -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def redirect(self, path: str) -> None:
        self.send_response(303)
        self.send_header("Location", path)
        self.end_headers()

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        path = parsed.path
        if path == "/":
            return self.respond(200, form_page())
        if path == "/historico":
            return self.respond(200, history_page())
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
            if not orcamento.farmaceutico_responsavel:
                raise ValueError("Informe o farmacêutico(a) responsável.")
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
    for candidate in range(preferred_port, preferred_port + 20):
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
