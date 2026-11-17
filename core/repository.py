from contextlib import closing
import csv
import os
import re
import sqlite3
from typing import Dict, List, Optional, Tuple

import openpyxl

from core.schema import SpecimenRecord

COLUMNS = [
    "id_mosquito",
    "parque",
    "metodo_coleta",
    "ponto",
    "tempo",
    "data",
    "coletor",
    "subfamilia",
    "tribo",
    "genero",
    "subgenero",
    "especie",
    "trilha_chave",
]

LEGACY_EXCEL_COLUMNS = [
    "id_mosquito",
    "parque",
    "metodo_coleta",
    "ponto",
    "tempo",
    "data",
    "coletor",
    "subfamilia",
    "tribo",
    "genero",
    "especie",
]


class SpecimenRepository:
    """Gerencia a persistência de dados em SQLite, Excel (.xlsx), CSV e SQL Server."""

    def __init__(
        self,
        sqlite_path: str = "dados_mosquitos.db",
        excel_path: str = "dados_mosquitos.xlsx",
    ):
        self.sqlite_path = sqlite_path
        self.excel_path = excel_path
        self.db_host: str = ""
        self.db_name: str = ""
        self.db_table: str = ""
        self._init_sqlite()
        self._import_existing_excel_if_empty()

    def _init_sqlite(self) -> None:
        with closing(sqlite3.connect(self.sqlite_path)) as conn:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS especimes (
                    pk INTEGER PRIMARY KEY AUTOINCREMENT,
                    id_mosquito TEXT,
                    parque TEXT,
                    metodo_coleta TEXT,
                    ponto TEXT,
                    tempo TEXT,
                    data TEXT,
                    coletor TEXT,
                    subfamilia TEXT,
                    tribo TEXT,
                    genero TEXT,
                    subgenero TEXT,
                    especie TEXT,
                    trilha_chave TEXT,
                    criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
                """
            )
            conn.commit()

    def _import_existing_excel_if_empty(self) -> None:
        """Importa dados existentes da planilha dados_mosquitos.xlsx se o banco SQLite estiver vazio."""
        if not os.path.exists(self.excel_path):
            return
        with closing(sqlite3.connect(self.sqlite_path)) as conn:
            cur = conn.execute("SELECT COUNT(*) FROM especimes")
            count = cur.fetchone()[0]
            if count > 0:
                return

        try:
            wb = openpyxl.load_workbook(self.excel_path)
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
            if len(rows) <= 1:
                return
            headers = [str(h).strip() if h is not None else "" for h in rows[0]]
            for row in rows[1:]:
                if not any(row):
                    continue
                row_dict = {headers[i]: ("" if row[i] is None else str(row[i])) for i in range(min(len(headers), len(row)))}
                rec = SpecimenRecord(
                    id_mosquito=row_dict.get("id_mosquito", ""),
                    parque=row_dict.get("parque", ""),
                    metodo_coleta=row_dict.get("metodo_coleta", ""),
                    ponto=row_dict.get("ponto", ""),
                    tempo=row_dict.get("tempo", ""),
                    data=row_dict.get("data", ""),
                    coletor=row_dict.get("coletor", ""),
                    subfamilia=row_dict.get("subfamilia", ""),
                    tribo=row_dict.get("tribo", ""),
                    genero=row_dict.get("genero", ""),
                    subgenero=row_dict.get("subgenero", ""),
                    especie=row_dict.get("especie", ""),
                    trilha_chave=row_dict.get("trilha_chave", "Importado do Excel"),
                )
                self._insert_sqlite(rec)
        except Exception as exc:
            print(f"Aviso ao importar Excel existente: {exc}")

    def _insert_sqlite(self, record: SpecimenRecord) -> int:
        with closing(sqlite3.connect(self.sqlite_path)) as conn:
            cur = conn.execute(
                """
                INSERT INTO especimes (
                    id_mosquito, parque, metodo_coleta, ponto, tempo, data, coletor,
                    subfamilia, tribo, genero, subgenero, especie, trilha_chave
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    record.id_mosquito,
                    record.parque,
                    record.metodo_coleta,
                    record.ponto,
                    record.tempo,
                    record.data,
                    record.coletor,
                    record.subfamilia,
                    record.tribo,
                    record.genero,
                    record.subgenero,
                    record.especie,
                    record.trilha_chave,
                ),
            )
            conn.commit()
            return int(cur.lastrowid or 0)

    def save_record(self, record: SpecimenRecord, sync_excel: bool = True, sync_sql_server: bool = False) -> Tuple[bool, str]:
        """Salva o registro no SQLite, sincroniza com Excel e opcionalmente com SQL Server."""
        try:
            self._insert_sqlite(record)
            if sync_excel:
                self.export_to_excel(self.excel_path)
            if sync_sql_server and self.db_host and self.db_name and self.db_table:
                ok_sql, msg_sql = self.save_to_sql_server(record)
                if not ok_sql:
                    return True, f"Salvo localmente (SQLite + Excel), mas falhou no SQL Server: {msg_sql}"
            return True, "Espécime registrado com sucesso (SQLite + Excel)."
        except Exception as exc:
            return False, f"Erro ao salvar registro: {exc}"

    def list_records(self, search_query: str = "") -> List[Dict[str, str]]:
        with closing(sqlite3.connect(self.sqlite_path)) as conn:
            conn.row_factory = sqlite3.Row
            if search_query.strip():
                q = f"%{search_query.strip()}%"
                cur = conn.execute(
                    """
                    SELECT * FROM especimes
                    WHERE id_mosquito LIKE ? OR parque LIKE ? OR coletor LIKE ?
                       OR subfamilia LIKE ? OR tribo LIKE ? OR genero LIKE ?
                       OR subgenero LIKE ? OR especie LIKE ?
                    ORDER BY pk DESC
                    """,
                    (q, q, q, q, q, q, q, q),
                )
            else:
                cur = conn.execute("SELECT * FROM especimes ORDER BY pk DESC")
            return [dict(r) for r in cur.fetchall()]

    def delete_record(self, pk: int, sync_excel: bool = True) -> None:
        with closing(sqlite3.connect(self.sqlite_path)) as conn:
            conn.execute("DELETE FROM especimes WHERE pk = ?", (pk,))
            conn.commit()
        if sync_excel:
            self.export_to_excel(self.excel_path)

    def export_to_excel(self, target_path: Optional[str] = None) -> str:
        path = target_path or self.excel_path
        records = self.list_records()
        # Inverter para ordem cronológica na planilha
        records.reverse()

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Mosquitos"
        ws.append(COLUMNS)
        for rec in records:
            ws.append([rec.get(col, "") for col in COLUMNS])
        wb.save(path)
        return path

    def export_to_csv(self, target_path: str) -> str:
        records = self.list_records()
        records.reverse()
        with open(target_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(COLUMNS)
            for rec in records:
                writer.writerow([rec.get(col, "") for col in COLUMNS])
        return target_path

    def configure_sql_server(self, host: str, database: str, table: str) -> Tuple[bool, str]:
        self.db_host = host.strip()
        self.db_name = database.strip()
        self.db_table = table.strip()
        if not re.match(r"^[A-Za-z_][A-Za-z0-9_\.]*$", self.db_table):
            return False, "Nome de tabela inválido."
        try:
            import pyodbc
            driver = "{ODBC Driver 17 for SQL Server}"
            conn_str = (
                f"DRIVER={driver};SERVER={self.db_host};DATABASE={self.db_name};"
                "Trusted_Connection=yes;Encrypt=no;"
            )
            conn = pyodbc.connect(conn_str, timeout=4)
            conn.close()
            return True, "Conexão com SQL Server testada com sucesso."
        except Exception as exc:
            return False, f"Falha ao conectar ao SQL Server: {exc}"

    def save_to_sql_server(self, record: SpecimenRecord) -> Tuple[bool, str]:
        if not re.match(r"^[A-Za-z_][A-Za-z0-9_\.]*$", self.db_table):
            return False, "Nome de tabela SQL inválido."
        try:
            import pyodbc
            driver = "{ODBC Driver 17 for SQL Server}"
            conn_str = (
                f"DRIVER={driver};SERVER={self.db_host};DATABASE={self.db_name};"
                "Trusted_Connection=yes;Encrypt=no;"
            )
            conn = pyodbc.connect(conn_str, timeout=4)
            cursor = conn.cursor()
            cursor.execute(
                f"""
                INSERT INTO {self.db_table}
                (id_mosquito, parque, metodo, ponto, tempo, dia, coletor, subfamilia, tribo, genero, especie)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    record.id_mosquito,
                    record.parque,
                    record.metodo_coleta,
                    record.ponto,
                    record.tempo,
                    record.data,
                    record.coletor,
                    record.subfamilia,
                    record.tribo,
                    record.genero,
                    record.especie,
                ),
            )
            conn.commit()
            cursor.close()
            conn.close()
            return True, "Salvo no SQL Server."
        except Exception as exc:
            return False, str(exc)
