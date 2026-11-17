from dataclasses import dataclass, field
from typing import Dict, List, Optional


@dataclass
class KeyOption:
    label: str  # "A" ou "B"
    text: str
    image: Optional[str] = "ArboTeamPng.png"
    next_step: Optional[str] = None
    set_taxonomy: Dict[str, str] = field(default_factory=dict)
    is_terminal: bool = False

    @classmethod
    def from_dict(cls, data: dict) -> "KeyOption":
        return cls(
            label=data.get("label", ""),
            text=data.get("text", ""),
            image=data.get("image", "ArboTeamPng.png"),
            next_step=data.get("next_step"),
            set_taxonomy=dict(data.get("set_taxonomy") or {}),
            is_terminal=bool(data.get("is_terminal", False)),
        )


@dataclass
class KeyStep:
    id: str
    key_id: str
    step_number: int
    title: str
    taxon_context: str
    option_a: KeyOption
    option_b: KeyOption

    @classmethod
    def from_dict(cls, data: dict) -> "KeyStep":
        return cls(
            id=data["id"],
            key_id=data["key_id"],
            step_number=int(data["step_number"]),
            title=data.get("title", data["id"]),
            taxon_context=data.get("taxon_context", ""),
            option_a=KeyOption.from_dict(data["option_a"]),
            option_b=KeyOption.from_dict(data["option_b"]),
        )


@dataclass
class HistoryEntry:
    step_id: str
    step_title: str
    chosen_option: str  # "A" ou "B"
    option_summary: str
    taxonomy_before: Dict[str, str]
    taxonomy_after: Dict[str, str]


@dataclass
class SpecimenRecord:
    id_mosquito: str
    parque: str
    metodo_coleta: str
    ponto: str
    tempo: str
    data: str
    coletor: str
    subfamilia: str = ""
    tribo: str = ""
    genero: str = ""
    subgenero: str = ""
    especie: str = ""
    trilha_chave: str = ""

    def to_dict(self) -> Dict[str, str]:
        return {
            "id_mosquito": self.id_mosquito,
            "parque": self.parque,
            "metodo_coleta": self.metodo_coleta,
            "ponto": self.ponto,
            "tempo": self.tempo,
            "data": self.data,
            "coletor": self.coletor,
            "subfamilia": self.subfamilia,
            "tribo": self.tribo,
            "genero": self.genero,
            "subgenero": self.subgenero,
            "especie": self.especie,
            "trilha_chave": self.trilha_chave,
        }
