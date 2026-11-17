import glob
import json
import os
import re
from typing import Dict, List, Optional, Set, Tuple

from core.schema import HistoryEntry, KeyOption, KeyStep


class KeyEngine:
    """Motor de navegação orientado a dados para chaves taxonômicas dicotômicas."""

    def __init__(self, keys_dir: str, glossary_path: Optional[str] = None):
        self.keys_dir = keys_dir
        self.glossary_path = glossary_path
        self.steps: Dict[str, KeyStep] = {}
        self.glossary: Dict[str, str] = {}

        # Catálogo taxonômico extraído automaticamente do grafo
        self.genus_hierarchy: Dict[str, Dict[str, str]] = {}
        self.genus_entry_step: Dict[str, str] = {}
        self.tribe_entry_step: Dict[str, str] = {}
        self.subfamily_entry_step: Dict[str, str] = {}
        self.species_by_genus: Dict[str, List[str]] = {}
        self.species_hierarchy: Dict[str, Dict[str, str]] = {}

        # Estado de navegação atual
        self.current_step_id: Optional[str] = None
        self.is_finished: bool = False
        self.history: List[HistoryEntry] = []
        self.current_taxonomy: Dict[str, str] = {
            "subfamilia": "",
            "tribo": "",
            "genero": "",
            "subgenero": "",
            "especie": "",
        }

        self.load_all()
        self.reset()

    def load_all(self) -> None:
        self.steps.clear()
        pattern = os.path.join(self.keys_dir, "*.json")
        for path in sorted(glob.glob(pattern)):
            with open(path, "r", encoding="utf-8") as f:
                payload = json.load(f)
            for item in payload.get("steps", []):
                step = KeyStep.from_dict(item)
                self.steps[step.id] = step

        if self.glossary_path and os.path.exists(self.glossary_path):
            with open(self.glossary_path, "r", encoding="utf-8") as f:
                self.glossary = json.load(f)

        self._build_taxonomy_catalog()

    def _build_taxonomy_catalog(self) -> None:
        """Percorre o grafo a partir de A1 propagando a hierarquia taxonômica."""
        self.genus_hierarchy.clear()
        self.genus_entry_step.clear()
        self.tribe_entry_step.clear()
        self.subfamily_entry_step.clear()
        self.species_by_genus.clear()
        self.species_hierarchy.clear()

        if "A1" not in self.steps:
            return

        # DFS com propagação de contexto taxonômico
        stack: List[Tuple[str, Dict[str, str]]] = [
            ("A1", {"subfamilia": "", "tribo": "", "genero": "", "subgenero": "", "especie": ""})
        ]
        visited_states: Set[Tuple[str, Tuple[Tuple[str, str], ...]]] = set()

        while stack:
            step_id, tax = stack.pop()
            if step_id not in self.steps:
                continue
            state_key = (step_id, tuple(sorted(tax.items())))
            if state_key in visited_states:
                continue
            visited_states.add(state_key)

            step = self.steps[step_id]
            for opt in (step.option_a, step.option_b):
                next_tax = dict(tax)
                for k, v in opt.set_taxonomy.items():
                    if v:
                        next_tax[k] = v

                # Inferir gênero a partir do binômio se necessário
                sp = next_tax.get("especie", "")
                if sp and not next_tax.get("genero") and " " in sp:
                    next_tax["genero"] = sp.split(" ")[0]

                subf = next_tax.get("subfamilia", "")
                tribo = next_tax.get("tribo", "")
                gen = next_tax.get("genero", "")
                subgen = next_tax.get("subgenero", "")

                if subf and opt.next_step and subf not in self.subfamily_entry_step:
                    self.subfamily_entry_step[subf] = opt.next_step
                if tribo and opt.next_step and tribo not in self.tribe_entry_step:
                    self.tribe_entry_step[tribo] = opt.next_step
                if gen:
                    if gen not in self.genus_hierarchy:
                        self.genus_hierarchy[gen] = {
                            "subfamilia": subf,
                            "tribo": tribo,
                            "genero": gen,
                        }
                    if opt.next_step and gen not in self.genus_entry_step:
                        self.genus_entry_step[gen] = opt.next_step

                if sp:
                    self.species_hierarchy[sp] = {
                        "subfamilia": subf,
                        "tribo": tribo,
                        "genero": gen,
                        "subgenero": subgen,
                        "especie": sp,
                    }
                    if gen:
                        sp_list = self.species_by_genus.setdefault(gen, [])
                        if sp not in sp_list:
                            sp_list.append(sp)

                if not opt.is_terminal and opt.next_step:
                    stack.append((opt.next_step, next_tax))

        for gen in self.species_by_genus:
            self.species_by_genus[gen].sort()

    def validate_graph(self) -> Dict[str, List[str]]:
        """Executa verificação completa de integridade estrutural e taxonômica do grafo."""
        missing_targets: List[str] = []
        empty_terminals: List[str] = []
        non_binomial_species: List[str] = []
        in_degree: Dict[str, int] = {sid: 0 for sid in self.steps}

        for sid, step in self.steps.items():
            for opt in (step.option_a, step.option_b):
                if opt.is_terminal:
                    if not any(opt.set_taxonomy.values()):
                        empty_terminals.append(f"{sid}{opt.label}")
                    sp = opt.set_taxonomy.get("especie", "")
                    if sp and len(sp.split()) < 2:
                        non_binomial_species.append(f"{sid}{opt.label}: '{sp}'")
                else:
                    if not opt.next_step or opt.next_step not in self.steps:
                        missing_targets.append(f"{sid}{opt.label} -> {opt.next_step}")
                    else:
                        in_degree[opt.next_step] = in_degree.get(opt.next_step, 0) + 1

        # Verificar alcançabilidade a partir de A1
        reachable: Set[str] = set()
        queue = ["A1"] if "A1" in self.steps else []
        while queue:
            curr = queue.pop()
            if curr in reachable or curr not in self.steps:
                continue
            reachable.add(curr)
            st = self.steps[curr]
            for opt in (st.option_a, st.option_b):
                if not opt.is_terminal and opt.next_step:
                    queue.append(opt.next_step)

        unreachable = sorted(set(self.steps.keys()) - reachable)

        # Verificar ciclos (DFS com pilha de recursão)
        cycles: List[str] = []
        visiting: Set[str] = set()
        visited: Set[str] = set()

        def dfs_cycle(node: str, path: List[str]) -> None:
            visiting.add(node)
            st = self.steps[node]
            for opt in (st.option_a, st.option_b):
                nxt = opt.next_step
                if not opt.is_terminal and nxt in self.steps:
                    if nxt in visiting:
                        cycles.append(" -> ".join(path + [nxt]))
                    elif nxt not in visited:
                        dfs_cycle(nxt, path + [nxt])
            visiting.remove(node)
            visited.add(node)

        if "A1" in self.steps:
            dfs_cycle("A1", ["A1"])

        return {
            "missing_targets": missing_targets,
            "unreachable": unreachable,
            "empty_terminals": empty_terminals,
            "non_binomial_species": non_binomial_species,
            "cycles": cycles,
        }

    def reset(self, start_step: str = "A1", initial_taxonomy: Optional[Dict[str, str]] = None) -> None:
        self.current_step_id = start_step if start_step in self.steps else None
        self.is_finished = False
        self.history.clear()
        self.current_taxonomy = {
            "subfamilia": "",
            "tribo": "",
            "genero": "",
            "subgenero": "",
            "especie": "",
        }
        if initial_taxonomy:
            for k, v in initial_taxonomy.items():
                if k in self.current_taxonomy and v:
                    self.current_taxonomy[k] = v

    @property
    def current_step(self) -> Optional[KeyStep]:
        if self.current_step_id and not self.is_finished:
            return self.steps.get(self.current_step_id)
        return None

    def choose_option(self, option_label: str) -> bool:
        """Executa a escolha da opção 'A' ou 'B' no passo atual. Retorna True se a chave terminou."""
        step = self.current_step
        if not step:
            return self.is_finished

        opt: KeyOption = step.option_a if option_label.upper() == "A" else step.option_b
        tax_before = dict(self.current_taxonomy)

        for k, v in opt.set_taxonomy.items():
            if k in self.current_taxonomy and v:
                self.current_taxonomy[k] = v

        # Preencher hierarquia superior caso tenha sido definida espécie ou gênero diretamente
        sp = self.current_taxonomy.get("especie", "")
        if sp and sp in self.species_hierarchy:
            for k, v in self.species_hierarchy[sp].items():
                if v and not self.current_taxonomy.get(k):
                    self.current_taxonomy[k] = v

        gen = self.current_taxonomy.get("genero", "")
        if gen and gen in self.genus_hierarchy:
            for k, v in self.genus_hierarchy[gen].items():
                if v and not self.current_taxonomy.get(k):
                    self.current_taxonomy[k] = v

        tax_after = dict(self.current_taxonomy)
        summary = opt.text[:90] + ("..." if len(opt.text) > 90 else "")
        self.history.append(
            HistoryEntry(
                step_id=step.id,
                step_title=step.title,
                chosen_option=opt.label,
                option_summary=summary,
                taxonomy_before=tax_before,
                taxonomy_after=tax_after,
            )
        )

        if opt.is_terminal or not opt.next_step:
            self.is_finished = True
        else:
            self.current_step_id = opt.next_step
            self.is_finished = False

        return self.is_finished

    def go_back(self) -> bool:
        """Retorna um passo na chave dicotômica. Retorna True se conseguiu voltar."""
        if not self.history:
            return False
        last = self.history.pop()
        self.current_step_id = last.step_id
        self.current_taxonomy = dict(last.taxonomy_before)
        self.is_finished = False
        return True

    def jump_to_history(self, index: int) -> bool:
        """Retorna até um passo específico do histórico (0-indexed)."""
        if index < 0 or index >= len(self.history):
            return False
        target_entry = self.history[index]
        self.history = self.history[:index]
        self.current_step_id = target_entry.step_id
        self.current_taxonomy = dict(target_entry.taxonomy_before)
        self.is_finished = False
        return True

    def finish_at_current_level(self) -> None:
        """Encerra a chave no nível taxonômico atual (ex.: espécime danificado)."""
        self.is_finished = True

    def start_from_genus(self, genero: str) -> bool:
        """Inicia a chave diretamente na chave de espécies do gênero selecionado."""
        if genero not in self.genus_hierarchy:
            return False
        hier = dict(self.genus_hierarchy[genero])
        entry = self.genus_entry_step.get(genero)
        if entry and entry in self.steps:
            self.reset(start_step=entry, initial_taxonomy=hier)
            return True
        else:
            # Gênero monotípico
            sp_list = self.species_by_genus.get(genero, [])
            if len(sp_list) == 1:
                hier["especie"] = sp_list[0]
            self.reset(start_step="A1", initial_taxonomy=hier)
            self.is_finished = True
            return True

    def set_direct_taxon(
        self,
        subfamilia: str = "",
        tribo: str = "",
        genero: str = "",
        subgenero: str = "",
        especie: str = "",
    ) -> None:
        """Define diretamente a identificação taxonômica preenchendo automaticamente os níveis superiores."""
        self.reset()
        if especie and especie in self.species_hierarchy:
            for k, v in self.species_hierarchy[especie].items():
                if v:
                    self.current_taxonomy[k] = v
        elif genero and genero in self.genus_hierarchy:
            for k, v in self.genus_hierarchy[genero].items():
                if v:
                    self.current_taxonomy[k] = v

        if subfamilia:
            self.current_taxonomy["subfamilia"] = subfamilia
        if tribo:
            self.current_taxonomy["tribo"] = tribo
        if genero:
            self.current_taxonomy["genero"] = genero
        if subgenero:
            self.current_taxonomy["subgenero"] = subgenero
        if especie:
            self.current_taxonomy["especie"] = especie

        self.history.append(
            HistoryEntry(
                step_id="DIRETO",
                step_title="Seleção Direta no Catálogo",
                chosen_option="★",
                option_summary=f"Definido diretamente: {especie or genero or tribo or subfamilia}",
                taxonomy_before={"subfamilia": "", "tribo": "", "genero": "", "subgenero": "", "especie": ""},
                taxonomy_after=dict(self.current_taxonomy),
            )
        )
        self.is_finished = True

    def get_trail_string(self) -> str:
        if not self.history:
            return ""
        return " > ".join(f"{h.step_id}({h.chosen_option})" for h in self.history)

    def extract_glossary_terms(self, text: str) -> List[Tuple[str, str]]:
        """Extrai termos do glossário presentes no texto atual para exibição contextual."""
        found: List[Tuple[str, str]] = []
        seen: Set[str] = set()
        # Procurar siglas entre parênteses ex: (Mpn), (Ta-4), (Esc)
        parens = re.findall(r"\(([A-Za-z0-9\-]+)\)", text)
        for token in parens:
            if token in self.glossary and token not in seen:
                found.append((token, self.glossary[token]))
                seen.add(token)
            elif token.startswith("Ta-") and "Ta-1..5" not in seen and "Ta-1..5" in self.glossary:
                found.append(("Ta-1..5", self.glossary["Ta-1..5"]))
                seen.add("Ta-1..5")

        # Procurar siglas conhecidas no corpo do texto
        for abbr, desc in self.glossary.items():
            if abbr in seen:
                continue
            if abbr == "Ta-1..5" and ("Ta-" in text or "Tarsômero" in text or "tarsômero" in text):
                found.append((abbr, desc))
                seen.add(abbr)
            elif len(abbr) >= 2 and re.search(rf"\b{re.escape(abbr)}\b", text):
                found.append((abbr, desc))
                seen.add(abbr)
        return found
