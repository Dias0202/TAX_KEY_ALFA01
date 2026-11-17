"""
TaxKeyAlfa 2.0 — Sistema Interativo de Chaves Taxonômicas de Culicidae Neotropicais
Baseado em Forattini (2002) - Culicidologia Médica Vol. 2 & Consoli & Lourenço-de-Oliveira (1994).
"""

import os
import sys

from core.engine import KeyEngine
from core.repository import SpecimenRepository
from ui.app_window import TaxKeyApp


def get_base_dir() -> str:
    """Retorna o diretório base do projeto ou executável PyInstaller."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))


def main() -> None:
    base_dir = get_base_dir()
    keys_dir = os.path.join(base_dir, "data", "keys")
    glossary_path = os.path.join(base_dir, "data", "glossary.json")
    sqlite_path = os.path.join(os.getcwd(), "dados_mosquitos.db")
    excel_path = os.path.join(os.getcwd(), "dados_mosquitos.xlsx")

    engine = KeyEngine(keys_dir=keys_dir, glossary_path=glossary_path)
    repository = SpecimenRepository(sqlite_path=sqlite_path, excel_path=excel_path)

    app = TaxKeyApp(engine=engine, repository=repository, assets_dir=base_dir)
    app.mainloop()


if __name__ == "__main__":
    main()
