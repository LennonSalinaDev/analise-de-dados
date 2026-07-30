from pathlib import Path
from shutil import copy2


SOURCE = Path(r"C:\Users\AspireVX15\Downloads\Orçamento_CLAMED_novo.docx")
TARGET = Path("modelos/Orçamento_CLAMED.docx")


def main() -> None:
    TARGET.parent.mkdir(parents=True, exist_ok=True)
    if not SOURCE.exists():
        raise SystemExit(f"Modelo original não encontrado: {SOURCE}")
    copy2(SOURCE, TARGET)
    print(f"Template criado em: {TARGET.resolve()}")


if __name__ == "__main__":
    main()
