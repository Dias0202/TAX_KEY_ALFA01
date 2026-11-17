import os
import tempfile
import unittest

from core.engine import KeyEngine
from core.repository import SpecimenRepository
from core.schema import SpecimenRecord

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
KEYS_DIR = os.path.join(ROOT_DIR, "data", "keys")
GLOSSARY_PATH = os.path.join(ROOT_DIR, "data", "glossary.json")


class TestTaxonomicKeyGraphAndEngine(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.engine = KeyEngine(KEYS_DIR, GLOSSARY_PATH)

    def test_total_loaded_steps_and_glossary(self):
        self.assertGreaterEqual(len(self.engine.steps), 339)
        self.assertGreaterEqual(len(self.engine.glossary), 45)

    def test_graph_structural_integrity(self):
        report = self.engine.validate_graph()
        self.assertEqual(report["missing_targets"], [], f"Alvos inexistentes: {report['missing_targets']}")
        self.assertEqual(report["unreachable"], [], f"Passos inacessíveis/órfãos: {report['unreachable']}")
        self.assertEqual(report["empty_terminals"], [], f"Terminais vazios: {report['empty_terminals']}")
        self.assertEqual(
            report["non_binomial_species"],
            [],
            f"Espécies fora do padrão binomial: {report['non_binomial_species']}",
        )
        self.assertEqual(report["cycles"], [], f"Ciclos detectados: {report['cycles']}")

    def test_legacy_bugs_fixed(self):
        # 1. D49B -> Aedes fluviatilis (era "")
        self.assertEqual(self.engine.steps["D49"].option_b.set_taxonomy.get("especie"), "Aedes fluviatilis")
        self.assertTrue(self.engine.steps["D49"].option_b.is_terminal)

        # 2. D50B -> Aedes kompi (saltava para D51)
        self.assertEqual(self.engine.steps["D50"].option_b.set_taxonomy.get("especie"), "Aedes kompi")
        self.assertTrue(self.engine.steps["D50"].option_b.is_terminal)

        # 3. D63B -> D66 e D64B -> D65 (restaurando D66..D76)
        self.assertEqual(self.engine.steps["D63"].option_b.next_step, "D66")
        self.assertEqual(self.engine.steps["D64"].option_b.next_step, "D65")

        # 4. F1B -> F7 (restaurando F7..F29 de Psorophora)
        self.assertEqual(self.engine.steps["F1"].option_b.next_step, "F7")
        self.assertIn("Psorophora", self.engine.steps["F1"].option_a.text)

        # 5. H10A -> H11 e H10B -> H12 (restaurando H11 e H12 de Sabethes)
        self.assertEqual(self.engine.steps["H10"].option_a.next_step, "H11")
        self.assertEqual(self.engine.steps["H10"].option_b.next_step, "H12")

        # 6. Psorophora mapeado na tribo Aedini (era Sabethini)
        self.assertEqual(self.engine.genus_hierarchy["Psorophora"]["tribo"], "Aedini")
        self.assertEqual(self.engine.genus_hierarchy["Psorophora"]["subfamilia"], "Culicinae")

    def test_all_major_genera_have_species_catalog(self):
        expected_genera = [
            "Anopheles",
            "Chagasia",
            "Toxorhynchites",
            "Uranotaenia",
            "Mansonia",
            "Coquillettidia",
            "Culiseta",
            "Aedeomyia",
            "Orthopodomyia",
            "Aedes",
            "Haemagogus",
            "Psorophora",
            "Sabethes",
            "Limatus",
            "Trichoprosopon",
            "Shannoniana",
            "Phoniomyia",
            "Wyeomyia",
            "Johnbelkinia",
            "Runchomyia",
            "Isostomyia",
            "Lutzia",
            "Culex",
        ]
        for gen in expected_genera:
            self.assertIn(gen, self.engine.genus_hierarchy, f"Gênero ausente na hierarquia: {gen}")
            self.assertTrue(len(self.engine.species_by_genus.get(gen, [])) > 0, f"Gênero sem espécies: {gen}")

    def test_navigation_and_back_button(self):
        self.engine.reset()
        self.assertEqual(self.engine.current_step_id, "A1")

        # Ir para Anophelinae -> Anopheles -> Kerteszia -> Anopheles bambusicolus
        self.engine.choose_option("A")  # A1A -> AN_GEN1 (Anophelinae)
        self.assertEqual(self.engine.current_taxonomy["subfamilia"], "Anophelinae")
        self.assertEqual(self.engine.current_step_id, "AN_GEN1")

        self.engine.choose_option("B")  # AN_GEN1B -> AN_ESP1 (Anopheles)
        self.assertEqual(self.engine.current_taxonomy["genero"], "Anopheles")

        # Testar botão Voltar
        self.assertTrue(self.engine.go_back())
        self.assertEqual(self.engine.current_step_id, "AN_GEN1")
        self.assertEqual(self.engine.current_taxonomy["genero"], "")

        # Avançar novamente até espécie terminal
        self.engine.choose_option("B")  # AN_ESP1
        self.engine.choose_option("A")  # AN_ESP2 (Kerteszia)
        finished = self.engine.choose_option("A")  # Anopheles bambusicolus
        self.assertTrue(finished)
        self.assertEqual(self.engine.current_taxonomy["especie"], "Anopheles bambusicolus")
        self.assertEqual(self.engine.current_taxonomy["subgenero"], "Kerteszia")

    def test_repository_sqlite_excel_csv(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            db_path = os.path.join(tmpdir, "test.db")
            xlsx_path = os.path.join(tmpdir, "test.xlsx")
            csv_path = os.path.join(tmpdir, "test.csv")

            repo = SpecimenRepository(sqlite_path=db_path, excel_path=xlsx_path)
            rec = SpecimenRecord(
                id_mosquito="101",
                parque="Mangabeiras",
                metodo_coleta="BG UP",
                ponto="1",
                tempo="Manhã",
                data="17/11/2026",
                coletor="Gabriel",
                subfamilia="Culicinae",
                tribo="Aedini",
                genero="Aedes",
                especie="Aedes aegypti",
                trilha_chave="A1(B) > A2(B) > B4(A) > C1(B) > C2(B) > C3(B) > D1(A) > D2(A)",
            )
            ok, msg = repo.save_record(rec)
            self.assertTrue(ok, msg)
            self.assertTrue(os.path.exists(xlsx_path))

            rows = repo.list_records()
            self.assertEqual(len(rows), 1)
            self.assertEqual(rows[0]["coletor"], "Gabriel")
            self.assertEqual(rows[0]["data"], "17/11/2026")
            self.assertEqual(rows[0]["especie"], "Aedes aegypti")

            repo.export_to_csv(csv_path)
            self.assertTrue(os.path.exists(csv_path))

    def test_ui_smoke_initialization_and_navigation(self):
        from ui.app_window import TaxKeyApp

        with tempfile.TemporaryDirectory() as tmpdir:
            db_path = os.path.join(tmpdir, "ui_test.db")
            xlsx_path = os.path.join(tmpdir, "ui_test.xlsx")
            repo = SpecimenRepository(sqlite_path=db_path, excel_path=xlsx_path)
            engine = KeyEngine(KEYS_DIR, GLOSSARY_PATH)

            app = TaxKeyApp(engine=engine, repository=repo, assets_dir=ROOT_DIR)
            app.update_idletasks()

            # Verificar passo inicial A1
            self.assertIn("[A1]", app.lbl_step_title.cget("text"))
            app.on_choose_option("B")  # A1B -> A2
            self.assertIn("[A2]", app.lbl_step_title.cget("text"))
            app.on_back_step()
            self.assertIn("[A1]", app.lbl_step_title.cget("text"))

            app.destroy()


if __name__ == "__main__":
    unittest.main()
