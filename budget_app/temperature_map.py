from __future__ import annotations

import os
import secrets
import shutil
import subprocess
from calendar import monthrange
from copy import copy
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment


TEMPLATE_PATH = Path("modelos/mapa_temperatura/modelo_mapa_temperatura_medicamentos.xlsx")
OUTPUT_DIR = Path("saida/mapas_temperatura")
FIRST_DAY_ROW = 10
LAST_DAY_ROW = 40
FIRST_VARIANT_COLUMN = 2
LAST_VARIANT_COLUMN = 21
INACTIVE_MARK = "****"

MONTH_OPTIONS = [
    (1, "Janeiro"),
    (2, "Fevereiro"),
    (3, "Março"),
    (4, "Abril"),
    (5, "Maio"),
    (6, "Junho"),
    (7, "Julho"),
    (8, "Agosto"),
    (9, "Setembro"),
    (10, "Outubro"),
    (11, "Novembro"),
    (12, "Dezembro"),
]
MONTH_NAMES = {month: name for month, name in MONTH_OPTIONS}


@dataclass(frozen=True)
class TemperatureMapInput:
    mes: int
    ano: int
    filial: str
    dias_extras: frozenset[int] = frozenset()

    @property
    def mes_nome(self) -> str:
        return MONTH_NAMES[self.mes]


def safe_filename_part(value: str) -> str:
    allowed = []
    for char in value.lower():
        if char.isalnum():
            allowed.append(char)
        elif char in {" ", "-", "_"}:
            allowed.append("_")
    return "".join(allowed).strip("_") or "mapa"


def parse_extra_days(values: list[str] | None) -> frozenset[int]:
    days = set()
    for value in values or []:
        value = value.strip()
        if not value:
            continue
        try:
            day = int(value)
        except ValueError as exc:
            raise ValueError("Campo Calendário: selecione apenas dias válidos.") from exc
        if day < 1 or day > 31:
            raise ValueError("Campo Calendário: selecione apenas dias entre 1 e 31.")
        days.add(day)
    return frozenset(days)


def parse_temperature_map_input(
    mes: str,
    ano: str,
    filial: str,
    dias_extras: list[str] | None = None,
) -> TemperatureMapInput:
    try:
        mes_int = int(mes)
    except ValueError as exc:
        raise ValueError("Campo Mês: selecione um mês válido.") from exc
    if mes_int not in MONTH_NAMES:
        raise ValueError("Campo Mês: selecione um mês válido.")

    try:
        ano_int = int(ano)
    except ValueError as exc:
        raise ValueError("Campo Ano: informe um ano válido.") from exc
    if ano_int < 2000 or ano_int > 2100:
        raise ValueError("Campo Ano: informe um ano entre 2000 e 2100.")

    filial_digits = "".join(char for char in filial if char.isdigit())
    if len(filial_digits) != 3:
        raise ValueError("Campo Filial: informe exatamente 3 dígitos.")

    return TemperatureMapInput(
        mes=mes_int,
        ano=ano_int,
        filial=filial_digits,
        dias_extras=parse_extra_days(dias_extras),
    )


def ensure_temperature_template() -> Path:
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"Modelo XLSX não encontrado em {TEMPLATE_PATH}")
    return TEMPLATE_PATH


def inactive_temperature_days(data: TemperatureMapInput) -> set[int]:
    days_in_month = monthrange(data.ano, data.mes)[1]
    inactive_days = set(data.dias_extras)
    inactive_days.update(day for day in range(days_in_month + 1, 32))
    for day in range(1, days_in_month + 1):
        if date(data.ano, data.mes, day).weekday() == 6:
            inactive_days.add(day)
    return inactive_days


def center_cell(cell) -> None:
    current = copy(cell.alignment)
    cell.alignment = Alignment(
        horizontal="center",
        vertical="center",
        text_rotation=current.textRotation,
        wrap_text=current.wrap_text,
        shrink_to_fit=current.shrink_to_fit,
        indent=current.indent,
    )


def fill_inactive_temperature_rows(worksheet, data: TemperatureMapInput) -> None:
    inactive_days = inactive_temperature_days(data)
    for day in inactive_days:
        row = FIRST_DAY_ROW + day - 1
        if row > LAST_DAY_ROW:
            continue
        for column in range(FIRST_VARIANT_COLUMN, LAST_VARIANT_COLUMN + 1):
            cell = worksheet.cell(row=row, column=column)
            cell.value = INACTIVE_MARK
            center_cell(cell)


def render_temperature_map(data: TemperatureMapInput) -> Path:
    template_path = ensure_temperature_template()
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    workbook = load_workbook(template_path)
    worksheet = workbook.active
    worksheet["M7"] = data.mes_nome
    worksheet["R7"] = data.ano
    worksheet["U7"] = int(data.filial)
    worksheet["U7"].number_format = "000"
    center_cell(worksheet["U7"])
    fill_inactive_temperature_rows(worksheet, data)

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nonce = secrets.token_hex(3)
    filename = (
        f"mapa_temperatura_filial_{data.filial}_"
        f"{safe_filename_part(data.mes_nome)}_{data.ano}_{stamp}_{nonce}.xlsx"
    )
    output_path = OUTPUT_DIR / filename
    workbook.save(output_path)
    return output_path


def powershell_quote(value: str) -> str:
    return "'" + value.replace("'", "''") + "'"


def convert_xlsx_to_pdf_with_excel(xlsx_path: Path) -> tuple[Path | None, str]:
    if os.name != "nt":
        return None, "Microsoft Excel disponível apenas no Windows."

    powershell = shutil.which("powershell") or shutil.which("pwsh")
    if not powershell:
        return None, "PowerShell não encontrado para acionar o Microsoft Excel."

    source_path = xlsx_path.resolve()
    pdf_path = xlsx_path.with_suffix(".pdf")
    temp_path = pdf_path.with_name(f"{pdf_path.stem}.tmp.pdf")
    temp_path.unlink(missing_ok=True)

    script = f"""
$ErrorActionPreference = 'Stop'
$xlsxPath = {powershell_quote(str(source_path))}
$pdfPath = {powershell_quote(str(temp_path.resolve()))}
if (-not (Test-Path -LiteralPath $xlsxPath)) {{
  throw "XLSX não encontrado: $xlsxPath"
}}
if (Test-Path -LiteralPath $pdfPath) {{
  Remove-Item -LiteralPath $pdfPath -Force
}}
$excel = $null
$workbook = $null
try {{
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $workbook = $excel.Workbooks.Open($xlsxPath, 3, $true)
  $workbook.ExportAsFixedFormat(0, $pdfPath)
}} finally {{
  if ($workbook -ne $null) {{
    $workbook.Close($false) | Out-Null
  }}
  if ($excel -ne $null) {{
    $excel.Quit() | Out-Null
  }}
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()
}}
"""
    result = subprocess.run(
        [
            powershell,
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-Command",
            script,
        ],
        check=False,
        capture_output=True,
        text=True,
        timeout=120,
    )
    if result.returncode == 0 and temp_path.exists():
        temp_path.replace(pdf_path)
        return pdf_path, "PDF gerado com sucesso pelo Microsoft Excel."

    temp_path.unlink(missing_ok=True)
    detail = (result.stderr or result.stdout or "erro desconhecido").strip()
    return None, f"Microsoft Excel não gerou o PDF: {detail}"


def find_soffice() -> str | None:
    found = shutil.which("soffice") or shutil.which("libreoffice")
    if found:
        return found
    for candidate in (
        Path("C:/Program Files/LibreOffice/program/soffice.exe"),
        Path("C:/Program Files (x86)/LibreOffice/program/soffice.exe"),
        Path.home() / "AppData/Local/Programs/LibreOffice/program/soffice.exe",
    ):
        if candidate.exists():
            return str(candidate)
    return None


def convert_xlsx_to_pdf_with_libreoffice(xlsx_path: Path) -> tuple[Path | None, str]:
    soffice = find_soffice()
    if not soffice:
        return None, "LibreOffice/soffice não encontrado."

    profile_dir = Path("tmp/libreoffice-profile").resolve()
    profile_dir.mkdir(parents=True, exist_ok=True)
    source_path = xlsx_path.resolve()
    output_dir = xlsx_path.parent.resolve()
    pdf_path = xlsx_path.with_suffix(".pdf")
    pdf_path.unlink(missing_ok=True)

    try:
        result = subprocess.run(
            [
                soffice,
                "--headless",
                "--nologo",
                "--nofirststartwizard",
                f"-env:UserInstallation=file:///{profile_dir.as_posix()}",
                "--convert-to",
                "pdf",
                "--outdir",
                str(output_dir),
                str(source_path),
            ],
            check=False,
            capture_output=True,
            text=True,
            timeout=90,
        )
    except Exception as exc:
        return None, f"LibreOffice não gerou o PDF: {exc}"

    if result.returncode == 0 and pdf_path.exists():
        return pdf_path, "PDF gerado com sucesso pelo LibreOffice."
    detail = (result.stderr or result.stdout or "erro desconhecido").strip()
    return None, f"LibreOffice não gerou o PDF: {detail}"


def convert_xlsx_to_pdf_with_docker(xlsx_path: Path) -> tuple[Path | None, str]:
    docker = shutil.which("docker")
    if not docker:
        return None, "Docker não encontrado."

    image = os.environ.get("LIBREOFFICE_DOCKER_IMAGE", "orcamento-clamed-local")
    output_dir = xlsx_path.parent.resolve()
    pdf_path = xlsx_path.with_suffix(".pdf")
    pdf_path.unlink(missing_ok=True)
    container_file = f"/work/{xlsx_path.name}"

    try:
        result = subprocess.run(
            [
                docker,
                "run",
                "--rm",
                "-v",
                f"{output_dir}:/work",
                "-w",
                "/work",
                "--entrypoint",
                "soffice",
                image,
                "--headless",
                "--nologo",
                "--nofirststartwizard",
                "--convert-to",
                "pdf",
                "--outdir",
                "/work",
                container_file,
            ],
            check=False,
            capture_output=True,
            text=True,
            timeout=120,
        )
    except Exception as exc:
        return None, f"Docker/LibreOffice não gerou o PDF: {exc}"

    if result.returncode == 0 and pdf_path.exists():
        return pdf_path, "PDF gerado com sucesso pelo LibreOffice no Docker."
    detail = (result.stderr or result.stdout or "erro desconhecido").strip()
    return None, f"Docker/LibreOffice não gerou o PDF: {detail}"


def convert_temperature_map_to_pdf(xlsx_path: Path) -> tuple[Path | None, str]:
    errors = []
    for converter in (
        convert_xlsx_to_pdf_with_libreoffice,
        convert_xlsx_to_pdf_with_docker,
        convert_xlsx_to_pdf_with_excel,
    ):
        pdf_path, status = converter(xlsx_path)
        if pdf_path is not None:
            return pdf_path, status
        errors.append(status)
    return None, (
        "PDF não gerado neste ambiente. O XLSX foi criado normalmente. "
        "Para gerar o PDF com o layout original, instale o LibreOffice, "
        "ative o Docker Desktop com a imagem local configurada, ou execute no Render "
        "com a imagem Docker atualizada."
    )


def output_file(filename: str) -> Path | None:
    if not filename or Path(filename).name != filename:
        return None
    path = OUTPUT_DIR / filename
    try:
        path.resolve().relative_to(OUTPUT_DIR.resolve())
    except ValueError:
        return None
    return path if path.exists() else None


def list_temperature_maps() -> list[Path]:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    return sorted(OUTPUT_DIR.glob("*.xlsx"), key=lambda path: path.stat().st_mtime, reverse=True)
