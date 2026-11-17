from datetime import datetime
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Dict, Optional

import customtkinter as ctk
from PIL import Image

from core.engine import KeyEngine
from core.repository import SpecimenRepository
from core.schema import KeyOption, KeyStep, SpecimenRecord

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# Paleta de cores científica (Dark Slate & Emerald)
BG_DARK = "#14171c"
BG_PANEL = "#1e222a"
BG_CARD = "#262b35"
ACCENT_PRIMARY = "#10b981"
ACCENT_PRIMARY_HOVER = "#059669"
ACCENT_OPT_A = "#2563eb"
ACCENT_OPT_A_HOVER = "#1d4ed8"
ACCENT_OPT_B = "#7c3aed"
ACCENT_OPT_B_HOVER = "#6d28d9"
TEXT_MUTED = "#94a3b8"
BORDER_COLOR = "#334155"


class TaxKeyApp(ctk.CTk):
    """Aplicação Desktop Nativa Windows para Identificação Taxonômica de Culicidae."""

    def __init__(self, engine: KeyEngine, repository: SpecimenRepository, assets_dir: str):
        super().__init__()
        self.engine = engine
        self.repository = repository
        self.assets_dir = assets_dir

        self.title("TaxKeyAlfa 2.0 — Chaves Taxonômicas de Culicidae Neotropicais (Forattini)")
        self.geometry("1440x860")
        self.minsize(1180, 720)
        self.configure(fg_color=BG_DARK)

        logo_path = os.path.join(self.assets_dir, "LogoFinal.png")
        if os.path.exists(logo_path):
            try:
                icon_img = tk.PhotoImage(file=logo_path)
                self.iconphoto(False, icon_img)
                self._icon_ref = icon_img
            except Exception:
                pass

        # Cache de imagens CTkImage
        self._image_cache: Dict[tuple, ctk.CTkImage] = {}

        # Variáveis do lote de coleta (Sidebar)
        self.var_mosquito_id = tk.StringVar(value=self._suggest_next_id())
        self.var_auto_increment = tk.BooleanVar(value=True)
        self.var_park = tk.StringVar(value="Mangabeiras")
        self.var_method = tk.StringVar(value="BG UP")
        self.var_point = tk.StringVar(value="1")
        self.var_time = tk.StringVar(value="Manhã")
        self.var_date = tk.StringVar(value=datetime.now().strftime("%d/%m/%Y"))
        self.var_collector = tk.StringVar(value="")
        self.var_sync_sql = tk.BooleanVar(value=False)

        # Construir layout principal
        self.grid_columnconfigure(0, weight=0, minsize=310)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._build_sidebar()
        self._build_main_tabs()
        self._bind_keyboard_shortcuts()

        # Renderizar estado inicial
        self.refresh_key_view()
        self.refresh_records_table()

    def _suggest_next_id(self) -> str:
        records = self.repository.list_records()
        for r in records:
            val = str(r.get("id_mosquito", "")).strip()
            if val.isdigit():
                return str(int(val) + 1)
        return "1"

    def _load_ctk_image(self, filename: str, size: tuple = (220, 220)) -> Optional[ctk.CTkImage]:
        key = (filename, size)
        if key in self._image_cache:
            return self._image_cache[key]
        path = os.path.join(self.assets_dir, filename)
        if not os.path.exists(path):
            path = os.path.join(self.assets_dir, "ArboTeamPng.png")
        if not os.path.exists(path):
            return None
        try:
            pil_img = Image.open(path)
            ctk_img = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=size)
            self._image_cache[key] = ctk_img
            return ctk_img
        except Exception:
            return None

    # ==========================================================================
    # SIDEBAR: DADOS DE COLETA & STATUS TAXONÔMICO
    # ==========================================================================
    def _build_sidebar(self) -> None:
        self.sidebar = ctk.CTkScrollableFrame(
            self,
            width=300,
            fg_color=BG_PANEL,
            corner_radius=0,
            scrollbar_button_color="#334155",
        )
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_columnconfigure(0, weight=1)

        # Cabeçalho com Logo
        header_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        header_frame.grid(row=0, column=0, padx=14, pady=(12, 6), sticky="ew")

        logo_img = self._load_ctk_image("LogoFinal.png", size=(56, 56))
        if logo_img:
            lbl_logo = ctk.CTkLabel(header_frame, image=logo_img, text="")
            lbl_logo.pack(side="left", padx=(0, 10))

        title_box = ctk.CTkFrame(header_frame, fg_color="transparent")
        title_box.pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(
            title_box,
            text="TaxKeyAlfa 2.0",
            font=ctk.CTkFont(size=18, weight="bold"),
            anchor="w",
        ).pack(fill="x")
        total_sp = sum(len(v) for v in self.engine.species_by_genus.values())
        ctk.CTkLabel(
            title_box,
            text=f"{len(self.engine.steps)} passos • {len(self.engine.genus_hierarchy)} gêneros • {total_sp} spp.",
            font=ctk.CTkFont(size=11),
            text_color=ACCENT_PRIMARY,
            anchor="w",
        ).pack(fill="x")

        # Card: Dados da Coleta
        coll_card = ctk.CTkFrame(self.sidebar, fg_color=BG_CARD, corner_radius=10, border_width=1, border_color=BORDER_COLOR)
        coll_card.grid(row=1, column=0, padx=12, pady=8, sticky="ew")
        coll_card.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            coll_card,
            text="📍 DADOS DO ESPÉCIME / LOTE",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#e2e8f0",
        ).grid(row=0, column=0, padx=12, pady=(10, 6), sticky="w")

        # ID do Mosquito + Auto-incremento
        id_row = ctk.CTkFrame(coll_card, fg_color="transparent")
        id_row.grid(row=1, column=0, padx=12, pady=4, sticky="ew")
        id_row.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(id_row, text="Nº / ID do Mosquito:", font=ctk.CTkFont(size=12)).grid(row=0, column=0, sticky="w")
        self.entry_id = ctk.CTkEntry(
            id_row,
            textvariable=self.var_mosquito_id,
            height=32,
            font=ctk.CTkFont(size=14, weight="bold"),
        )
        self.entry_id.grid(row=1, column=0, pady=(2, 0), sticky="ew")

        chk_auto = ctk.CTkCheckBox(
            id_row,
            text="Auto +1",
            variable=self.var_auto_increment,
            width=70,
            checkbox_width=18,
            checkbox_height=18,
            font=ctk.CTkFont(size=11),
        )
        chk_auto.grid(row=1, column=1, padx=(8, 0), pady=(2, 0))

        # Parque
        self._add_combo_field(
            coll_card,
            row=2,
            label="Parque / Localidade:",
            variable=self.var_park,
            values=["Mangabeiras", "Lagoa do nado", "Estação ecologica", "Ursulina", "Jacques cousteau", "Horto"],
        )

        # Método de coleta
        self._add_combo_field(
            coll_card,
            row=3,
            label="Método de Coleta:",
            variable=self.var_method,
            values=["BG UP", "BG DOWN", "PUÇA UP", "PUÇA DOWN", "ASPIRADOR", "CDC", "SHANNON", "OVITRAMPA"],
        )

        # Ponto e Tempo na mesma linha
        pt_row = ctk.CTkFrame(coll_card, fg_color="transparent")
        pt_row.grid(row=4, column=0, padx=12, pady=4, sticky="ew")
        pt_row.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkLabel(pt_row, text="Ponto:", font=ctk.CTkFont(size=12)).grid(row=0, column=0, sticky="w")
        ctk.CTkComboBox(
            pt_row,
            variable=self.var_point,
            values=["1", "2", "3", "4", "5"],
            height=30,
        ).grid(row=1, column=0, padx=(0, 4), pady=(2, 0), sticky="ew")

        ctk.CTkLabel(pt_row, text="Turno / Tempo:", font=ctk.CTkFont(size=12)).grid(row=0, column=1, padx=(4, 0), sticky="w")
        ctk.CTkComboBox(
            pt_row,
            variable=self.var_time,
            values=["Manhã", "Tarde", "Noite", "Crepuscular"],
            height=30,
        ).grid(row=1, column=1, padx=(4, 0), pady=(2, 0), sticky="ew")

        # Data e Coletor
        dt_row = ctk.CTkFrame(coll_card, fg_color="transparent")
        dt_row.grid(row=5, column=0, padx=12, pady=(4, 12), sticky="ew")
        dt_row.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkLabel(dt_row, text="Data:", font=ctk.CTkFont(size=12)).grid(row=0, column=0, sticky="w")
        ctk.CTkEntry(dt_row, textvariable=self.var_date, height=30, placeholder_text="DD/MM/AAAA").grid(
            row=1, column=0, padx=(0, 4), pady=(2, 0), sticky="ew"
        )

        ctk.CTkLabel(dt_row, text="Coletor:", font=ctk.CTkFont(size=12)).grid(row=0, column=1, padx=(4, 0), sticky="w")
        ctk.CTkEntry(dt_row, textvariable=self.var_collector, height=30, placeholder_text="Nome").grid(
            row=1, column=1, padx=(4, 0), pady=(2, 0), sticky="ew"
        )

        # Card: Classificação Taxonômica Atual
        tax_card = ctk.CTkFrame(self.sidebar, fg_color=BG_CARD, corner_radius=10, border_width=1, border_color=BORDER_COLOR)
        tax_card.grid(row=2, column=0, padx=12, pady=8, sticky="ew")
        tax_card.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            tax_card,
            text="🧬 IDENTIFICAÇÃO EM CURSO",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#e2e8f0",
        ).grid(row=0, column=0, columnspan=2, padx=12, pady=(10, 6), sticky="w")

        self.tax_labels: Dict[str, ctk.CTkLabel] = {}
        ranks = [
            ("subfamilia", "Subfamília:"),
            ("tribo", "Tribo:"),
            ("genero", "Gênero:"),
            ("subgenero", "Subgênero:"),
            ("especie", "Espécie:"),
        ]
        for idx, (key, title) in enumerate(ranks, start=1):
            ctk.CTkLabel(tax_card, text=title, font=ctk.CTkFont(size=12), text_color=TEXT_MUTED).grid(
                row=idx, column=0, padx=(12, 6), pady=3, sticky="w"
            )
            val_lbl = ctk.CTkLabel(
                tax_card,
                text="—",
                font=ctk.CTkFont(size=12, weight="bold", slant="italic" if key in ("genero", "subgenero", "especie") else "roman"),
                text_color="#f8fafc",
                anchor="w",
            )
            val_lbl.grid(row=idx, column=1, padx=(0, 12), pady=3, sticky="ew")
            self.tax_labels[key] = val_lbl

        self.btn_save_sidebar = ctk.CTkButton(
            tax_card,
            text="💾 Salvar Espécime Atual",
            fg_color=ACCENT_PRIMARY,
            hover_color=ACCENT_PRIMARY_HOVER,
            height=36,
            font=ctk.CTkFont(size=13, weight="bold"),
            command=self.save_current_specimen,
        )
        self.btn_save_sidebar.grid(row=6, column=0, columnspan=2, padx=12, pady=(10, 12), sticky="ew")

        # Atalhos de Teclado Info
        kb_card = ctk.CTkFrame(self.sidebar, fg_color=BG_CARD, corner_radius=10)
        kb_card.grid(row=3, column=0, padx=12, pady=8, sticky="ew")
        ctk.CTkLabel(
            kb_card,
            text="⌨️ ATALHOS DE LABORATÓRIO\n• [A] ou [1]: Escolher Opção A\n• [B] ou [2]: Escolher Opção B\n• [Backspace]: Voltar 1 Passo\n• [Ctrl+S]: Salvar Espécime",
            font=ctk.CTkFont(size=11),
            text_color=TEXT_MUTED,
            justify="left",
        ).pack(padx=12, pady=10, anchor="w")

    def _add_combo_field(self, parent, row: int, label: str, variable: tk.StringVar, values: list) -> None:
        box = ctk.CTkFrame(parent, fg_color="transparent")
        box.grid(row=row, column=0, padx=12, pady=4, sticky="ew")
        box.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(box, text=label, font=ctk.CTkFont(size=12)).grid(row=0, column=0, sticky="w")
        ctk.CTkComboBox(box, variable=variable, values=values, height=30).grid(row=1, column=0, pady=(2, 0), sticky="ew")

    # ==========================================================================
    # ABAS PRINCIPAIS
    # ==========================================================================
    def _build_main_tabs(self) -> None:
        self.tabview = ctk.CTkTabview(
            self,
            fg_color=BG_PANEL,
            segmented_button_selected_color=ACCENT_PRIMARY,
            segmented_button_selected_hover_color=ACCENT_PRIMARY_HOVER,
            corner_radius=12,
        )
        self.tabview.grid(row=0, column=1, padx=14, pady=10, sticky="nsew")

        self.tab_key = self.tabview.add("🔬 Chave Dicotômica")
        self.tab_catalog = self.tabview.add("⚡ Catálogo & Busca Direta")
        self.tab_records = self.tabview.add("📊 Registros Salvos")
        self.tab_glossary = self.tabview.add("📖 Glossário Anatômico")

        self._build_key_tab()
        self._build_catalog_tab()
        self._build_records_tab()
        self._build_glossary_tab()

    # ==========================================================================
    # ABA 1: CHAVE DICOTÔMICA INTERATIVA
    # ==========================================================================
    def _build_key_tab(self) -> None:
        self.tab_key.grid_columnconfigure((0, 1), weight=1)
        self.tab_key.grid_rowconfigure(2, weight=1)

        # Topo: Título do Passo + Atalho "Ir para Gênero"
        top_bar = ctk.CTkFrame(self.tab_key, fg_color=BG_CARD, corner_radius=10)
        top_bar.grid(row=0, column=0, columnspan=2, padx=8, pady=(6, 4), sticky="ew")
        top_bar.grid_columnconfigure(0, weight=1)

        title_col = ctk.CTkFrame(top_bar, fg_color="transparent")
        title_col.grid(row=0, column=0, padx=14, pady=8, sticky="ew")

        self.lbl_step_context = ctk.CTkLabel(
            title_col,
            text="Família Culicidae",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=ACCENT_PRIMARY,
            anchor="w",
        )
        self.lbl_step_context.pack(fill="x")

        self.lbl_step_title = ctk.CTkLabel(
            title_col,
            text="Subfamílias de Culicidae — Passo 1",
            font=ctk.CTkFont(size=20, weight="bold"),
            anchor="w",
        )
        self.lbl_step_title.pack(fill="x")

        # Seletor rápido de salto para chave de gênero
        jump_box = ctk.CTkFrame(top_bar, fg_color="transparent")
        jump_box.grid(row=0, column=1, padx=14, pady=8, sticky="e")
        ctk.CTkLabel(jump_box, text="Ir direto para chave do gênero:", font=ctk.CTkFont(size=12), text_color=TEXT_MUTED).pack(
            side="left", padx=(0, 8)
        )
        genera_sorted = sorted(self.engine.genus_hierarchy.keys())
        self.combo_quick_genus = ctk.CTkComboBox(
            jump_box,
            values=["Início (A1 - Subfamílias)"] + genera_sorted,
            width=210,
            command=self._on_quick_genus_jump,
        )
        self.combo_quick_genus.set("Início (A1 - Subfamílias)")
        self.combo_quick_genus.pack(side="left")

        # Barra de Breadcrumb (Histórico clicável)
        self.breadcrumb_frame = ctk.CTkScrollableFrame(
            self.tab_key,
            height=36,
            orientation="horizontal",
            fg_color=BG_DARK,
            corner_radius=8,
        )
        self.breadcrumb_frame.grid(row=1, column=0, columnspan=2, padx=8, pady=4, sticky="ew")

        # Container Central: Cards A e B (ou Tela de Conclusão)
        self.center_container = ctk.CTkFrame(self.tab_key, fg_color="transparent")
        self.center_container.grid(row=2, column=0, columnspan=2, padx=4, pady=4, sticky="nsew")
        self.center_container.grid_columnconfigure((0, 1), weight=1)
        self.center_container.grid_rowconfigure(0, weight=1)

        self.card_a = self._create_option_card(self.center_container, "A", ACCENT_OPT_A, ACCENT_OPT_A_HOVER, col=0)
        self.card_b = self._create_option_card(self.center_container, "B", ACCENT_OPT_B, ACCENT_OPT_B_HOVER, col=1)

        # Painel de Confirmação Final (exibido quando is_finished=True)
        self.finish_panel = ctk.CTkFrame(
            self.center_container,
            fg_color=BG_CARD,
            corner_radius=12,
            border_width=2,
            border_color=ACCENT_PRIMARY,
        )
        self.finish_panel.grid_columnconfigure(0, weight=1)

        self.lbl_finish_header = ctk.CTkLabel(
            self.finish_panel,
            text="✅ IDENTIFICAÇÃO CONCLUÍDA",
            font=ctk.CTkFont(size=22, weight="bold"),
            text_color=ACCENT_PRIMARY,
        )
        self.lbl_finish_header.pack(pady=(28, 10))

        self.lbl_finish_details = ctk.CTkLabel(
            self.finish_panel,
            text="",
            font=ctk.CTkFont(size=16),
            justify="center",
        )
        self.lbl_finish_details.pack( padx=24, pady=10)

        self.lbl_finish_trail = ctk.CTkLabel(
            self.finish_panel,
            text="",
            font=ctk.CTkFont(size=12),
            text_color=TEXT_MUTED,
            wraplength=760,
        )
        self.lbl_finish_trail.pack(padx=24, pady=(4, 18))

        finish_btns = ctk.CTkFrame(self.finish_panel, fg_color="transparent")
        finish_btns.pack(pady=12)

        ctk.CTkButton(
            finish_btns,
            text="💾 Salvar e Identificar Próximo Mosquito",
            fg_color=ACCENT_PRIMARY,
            hover_color=ACCENT_PRIMARY_HOVER,
            height=44,
            width=300,
            font=ctk.CTkFont(size=15, weight="bold"),
            command=self.save_and_next_mosquito,
        ).pack(side="left", padx=10)

        ctk.CTkButton(
            finish_btns,
            text="⬅ Voltar 1 Passo",
            fg_color="#475569",
            hover_color="#334155",
            height=44,
            width=160,
            font=ctk.CTkFont(size=14),
            command=self.on_back_step,
        ).pack(side="left", padx=10)

        ctk.CTkButton(
            finish_btns,
            text="🔄 Reiniciar do Início (A1)",
            fg_color="#334155",
            hover_color="#1e293b",
            height=44,
            width=200,
            font=ctk.CTkFont(size=14),
            command=self.on_reset_key,
        ).pack(side="left", padx=10)

        # Barra Contextual de Glossário (mostra termos detectados na pergunta atual)
        self.glossary_hint_bar = ctk.CTkFrame(self.tab_key, fg_color=BG_DARK, corner_radius=8)
        self.glossary_hint_bar.grid(row=3, column=0, columnspan=2, padx=8, pady=4, sticky="ew")
        self.lbl_glossary_hints = ctk.CTkLabel(
            self.glossary_hint_bar,
            text="",
            font=ctk.CTkFont(size=12),
            text_color="#cbd5e1",
            justify="left",
            anchor="w",
            wraplength=980,
        )
        self.lbl_glossary_hints.pack(fill="x", padx=12, pady=6)

        # Barra Inferior de Controle da Chave
        bottom_nav = ctk.CTkFrame(self.tab_key, fg_color="transparent")
        bottom_nav.grid(row=4, column=0, columnspan=2, padx=8, pady=(4, 6), sticky="ew")

        self.btn_back = ctk.CTkButton(
            bottom_nav,
            text="⬅ Voltar Passo [Backspace]",
            fg_color="#475569",
            hover_color="#334155",
            width=190,
            height=34,
            command=self.on_back_step,
        )
        self.btn_back.pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            bottom_nav,
            text="🔄 Reiniciar Chave (A1)",
            fg_color="#334155",
            hover_color="#1e293b",
            width=170,
            height=34,
            command=self.on_reset_key,
        ).pack(side="left", padx=8)

        self.btn_stop_here = ctk.CTkButton(
            bottom_nav,
            text="⏹ Confirmar no Nível Taxonômico Atual",
            fg_color="#0f766e",
            hover_color="#115e59",
            width=260,
            height=34,
            command=self.on_finish_current_level,
        )
        self.btn_stop_here.pack(side="right")

    def _create_option_card(self, parent, opt_letter: str, color: str, hover_color: str, col: int) -> dict:
        frame = ctk.CTkFrame(
            parent,
            fg_color=BG_CARD,
            corner_radius=12,
            border_width=1,
            border_color=BORDER_COLOR,
        )
        frame.grid(row=0, column=col, padx=6, pady=4, sticky="nsew")
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(1, weight=1)

        # Cabeçalho do Card
        hdr = ctk.CTkFrame(frame, fg_color="transparent")
        hdr.grid(row=0, column=0, padx=14, pady=(12, 4), sticky="ew")

        badge = ctk.CTkLabel(
            hdr,
            text=f"OPÇÃO {opt_letter}",
            fg_color=color,
            corner_radius=6,
            font=ctk.CTkFont(size=14, weight="bold"),
            padx=12,
            pady=4,
        )
        badge.pack(side="left")

        target_badge = ctk.CTkLabel(
            hdr,
            text="",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=ACCENT_PRIMARY,
        )
        target_badge.pack(side="right")

        # Caixa de Texto da Descrição Morfológica
        textbox = ctk.CTkTextbox(
            frame,
            font=ctk.CTkFont(size=15),
            fg_color=BG_DARK,
            text_color="#f1f5f9",
            wrap="word",
            corner_radius=8,
            height=130,
        )
        textbox.grid(row=1, column=0, padx=14, pady=6, sticky="nsew")

        # Área de Imagem + Botão Zoom
        img_frame = ctk.CTkFrame(frame, fg_color="transparent")
        img_frame.grid(row=2, column=0, padx=14, pady=4, sticky="ew")

        img_label = ctk.CTkLabel(img_frame, text="")
        img_label.pack(side="left", expand=True)

        btn_zoom = ctk.CTkButton(
            img_frame,
            text="🔍 Ampliar Prancha",
            width=130,
            height=28,
            fg_color="#334155",
            hover_color="#475569",
            font=ctk.CTkFont(size=12),
            command=lambda l=opt_letter: self._open_image_zoom_modal(l),
        )
        btn_zoom.pack(side="right", padx=6)

        # Botão de Escolha Principal
        btn_select = ctk.CTkButton(
            frame,
            text=f"Escolher Opção {opt_letter}  [Tecla {opt_letter}]",
            fg_color=color,
            hover_color=hover_color,
            height=44,
            font=ctk.CTkFont(size=15, weight="bold"),
            command=lambda l=opt_letter: self.on_choose_option(l),
        )
        btn_select.grid(row=3, column=0, padx=14, pady=(6, 14), sticky="ew")

        return {
            "frame": frame,
            "target_badge": target_badge,
            "textbox": textbox,
            "img_label": img_label,
            "btn_select": btn_select,
            "current_image_file": "ArboTeamPng.png",
        }

    # ==========================================================================
    # ABA 2: CATÁLOGO & BUSCA DIRETA DE TÁXONS
    # ==========================================================================
    def _build_catalog_tab(self) -> None:
        self.tab_catalog.grid_columnconfigure(0, weight=1)
        self.tab_catalog.grid_rowconfigure(1, weight=1)

        filter_bar = ctk.CTkFrame(self.tab_catalog, fg_color=BG_CARD, corner_radius=10)
        filter_bar.grid(row=0, column=0, padx=8, pady=8, sticky="ew")

        ctk.CTkLabel(filter_bar, text="🔎 Filtrar por Gênero:", font=ctk.CTkFont(size=13, weight="bold")).pack(
            side="left", padx=(14, 6), pady=10
        )

        genera_list = ["Todos os Gêneros"] + sorted(self.engine.genus_hierarchy.keys())
        self.var_catalog_genus = tk.StringVar(value="Todos os Gêneros")
        self.combo_catalog_genus = ctk.CTkComboBox(
            filter_bar,
            variable=self.var_catalog_genus,
            values=genera_list,
            width=200,
            command=lambda _: self.refresh_catalog_list(),
        )
        self.combo_catalog_genus.pack(side="left", padx=6, pady=10)

        ctk.CTkLabel(filter_bar, text="Busca por nome:", font=ctk.CTkFont(size=13, weight="bold")).pack(
            side="left", padx=(16, 6), pady=10
        )
        self.var_catalog_search = tk.StringVar()
        self.var_catalog_search.trace_add("write", lambda *_: self.refresh_catalog_list())
        ctk.CTkEntry(
            filter_bar,
            textvariable=self.var_catalog_search,
            placeholder_text="Ex: aegypti, darlingi, Sabethes, Melanoconion...",
            width=300,
        ).pack(side="left", padx=6, pady=10)

        ctk.CTkButton(
            filter_bar,
            text="🔑 Abrir Chave do Gênero Selecionado",
            fg_color=ACCENT_OPT_A,
            hover_color=ACCENT_OPT_A_HOVER,
            command=self._open_selected_genus_key_from_catalog,
        ).pack(side="right", padx=14, pady=10)

        # Tabela do Catálogo Taxonômico
        table_frame = ctk.CTkFrame(self.tab_catalog, fg_color=BG_CARD, corner_radius=10)
        table_frame.grid(row=1, column=0, padx=8, pady=4, sticky="nsew")
        table_frame.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(0, weight=1)

        self._configure_treeview_style()

        cols = ("subfamilia", "tribo", "genero", "subgenero", "especie")
        self.catalog_tree = ttk.Treeview(table_frame, columns=cols, show="headings", selectmode="browse")
        headers = {
            "subfamilia": ("Subfamília", 140),
            "tribo": ("Tribo", 140),
            "genero": ("Gênero", 160),
            "subgenero": ("Subgênero", 160),
            "especie": ("Espécie (Binômio)", 280),
        }
        for c, (text, w) in headers.items():
            self.catalog_tree.heading(c, text=text)
            self.catalog_tree.column(c, width=w, anchor="w")

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.catalog_tree.yview)
        self.catalog_tree.configure(yscrollcommand=vsb.set)
        self.catalog_tree.grid(row=0, column=0, sticky="nsew", padx=(8, 0), pady=8)
        vsb.grid(row=0, column=1, sticky="ns", padx=(0, 8), pady=8)

        # Barra de ação do catálogo
        cat_actions = ctk.CTkFrame(self.tab_catalog, fg_color="transparent")
        cat_actions.grid(row=2, column=0, padx=8, pady=8, sticky="ew")

        ctk.CTkButton(
            cat_actions,
            text="✅ Definir Táxon Selecionado Diretamente",
            fg_color=ACCENT_PRIMARY,
            hover_color=ACCENT_PRIMARY_HOVER,
            height=38,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self._select_taxon_from_catalog,
        ).pack(side="right", padx=6)

        ctk.CTkButton(
            cat_actions,
            text="💾 Definir e Já Salvar Espécime",
            fg_color="#059669",
            hover_color="#047857",
            height=38,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=lambda: self._select_taxon_from_catalog(auto_save=True),
        ).pack(side="right", padx=6)

        self.refresh_catalog_list()

    # ==========================================================================
    # ABA 3: REGISTROS SALVOS & BANCO DE DADOS
    # ==========================================================================
    def _build_records_tab(self) -> None:
        self.tab_records.grid_columnconfigure(0, weight=1)
        self.tab_records.grid_rowconfigure(1, weight=1)

        top_bar = ctk.CTkFrame(self.tab_records, fg_color=BG_CARD, corner_radius=10)
        top_bar.grid(row=0, column=0, padx=8, pady=8, sticky="ew")

        ctk.CTkLabel(top_bar, text="🔎 Buscar Registro:", font=ctk.CTkFont(size=13, weight="bold")).pack(
            side="left", padx=(14, 6), pady=10
        )
        self.var_record_search = tk.StringVar()
        self.var_record_search.trace_add("write", lambda *_: self.refresh_records_table())
        ctk.CTkEntry(
            top_bar,
            textvariable=self.var_record_search,
            placeholder_text="Buscar por ID, Parque, Coletor, Gênero, Espécie...",
            width=320,
        ).pack(side="left", padx=6, pady=10)

        self.lbl_records_count = ctk.CTkLabel(top_bar, text="0 registros", font=ctk.CTkFont(size=13), text_color=ACCENT_PRIMARY)
        self.lbl_records_count.pack(side="left", padx=14)

        ctk.CTkButton(
            top_bar,
            text="🗄 Configurar SQL Server",
            fg_color="#334155",
            hover_color="#475569",
            command=self._open_sql_server_modal,
        ).pack(side="right", padx=12, pady=10)

        # Tabela de Registros
        rec_frame = ctk.CTkFrame(self.tab_records, fg_color=BG_CARD, corner_radius=10)
        rec_frame.grid(row=1, column=0, padx=8, pady=4, sticky="nsew")
        rec_frame.grid_columnconfigure(0, weight=1)
        rec_frame.grid_rowconfigure(0, weight=1)

        cols = (
            "pk",
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
        )
        self.records_tree = ttk.Treeview(rec_frame, columns=cols, show="headings", selectmode="browse")
        col_defs = [
            ("pk", "#", 45),
            ("id_mosquito", "ID Mosq.", 75),
            ("parque", "Parque", 120),
            ("metodo_coleta", "Método", 85),
            ("ponto", "Pt", 45),
            ("tempo", "Turno", 70),
            ("data", "Data", 90),
            ("coletor", "Coletor", 95),
            ("subfamilia", "Subfamília", 100),
            ("tribo", "Tribo", 95),
            ("genero", "Gênero", 105),
            ("subgenero", "Subgênero", 100),
            ("especie", "Espécie", 170),
            ("trilha_chave", "Trilha da Chave", 200),
        ]
        for col_id, text, width in col_defs:
            self.records_tree.heading(col_id, text=text)
            self.records_tree.column(col_id, width=width, anchor="w")

        vsb = ttk.Scrollbar(rec_frame, orient="vertical", command=self.records_tree.yview)
        hsb = ttk.Scrollbar(rec_frame, orient="horizontal", command=self.records_tree.xview)
        self.records_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.records_tree.grid(row=0, column=0, sticky="nsew", padx=(8, 0), pady=(8, 0))
        vsb.grid(row=0, column=1, sticky="ns", padx=(0, 8), pady=(8, 0))
        hsb.grid(row=1, column=0, sticky="ew", padx=(8, 0), pady=(0, 8))

        # Botões de Exportação e Exclusão
        bot_bar = ctk.CTkFrame(self.tab_records, fg_color="transparent")
        bot_bar.grid(row=2, column=0, padx=8, pady=8, sticky="ew")

        ctk.CTkButton(
            bot_bar,
            text="🗑 Excluir Registro Selecionado",
            fg_color="#dc2626",
            hover_color="#b91c1c",
            command=self._delete_selected_record,
        ).pack(side="left", padx=4)

        ctk.CTkButton(
            bot_bar,
            text="📄 Exportar para CSV (.csv)",
            fg_color="#2563eb",
            hover_color="#1d4ed8",
            command=self._export_csv_dialog,
        ).pack(side="right", padx=6)

        ctk.CTkButton(
            bot_bar,
            text="📥 Exportar Planilha Excel (.xlsx)",
            fg_color=ACCENT_PRIMARY,
            hover_color=ACCENT_PRIMARY_HOVER,
            command=self._export_excel_dialog,
        ).pack(side="right", padx=6)

    # ==========================================================================
    # ABA 4: GLOSSÁRIO ANATÔMICO DE FORATTINI
    # ==========================================================================
    def _build_glossary_tab(self) -> None:
        self.tab_glossary.grid_columnconfigure(0, weight=1)
        self.tab_glossary.grid_rowconfigure(1, weight=1)

        top = ctk.CTkFrame(self.tab_glossary, fg_color=BG_CARD, corner_radius=10)
        top.grid(row=0, column=0, padx=8, pady=8, sticky="ew")

        ctk.CTkLabel(
            top,
            text="📖 Glossário de Siglas Morfológicas (Forattini, 2002 — Culicidologia Médica Vol. 2)",
            font=ctk.CTkFont(size=15, weight="bold"),
        ).pack(side="left", padx=14, pady=10)

        self.var_glossary_search = tk.StringVar()
        self.var_glossary_search.trace_add("write", lambda *_: self._refresh_glossary_view())
        ctk.CTkEntry(
            top,
            textvariable=self.var_glossary_search,
            placeholder_text="Filtrar sigla ou estrutura (ex: Mpn, Esc, Ta, cerda)...",
            width=320,
        ).pack(side="right", padx=14, pady=10)

        self.glossary_scroll = ctk.CTkScrollableFrame(self.tab_glossary, fg_color=BG_CARD, corner_radius=10)
        self.glossary_scroll.grid(row=1, column=0, padx=8, pady=4, sticky="nsew")
        self.glossary_scroll.grid_columnconfigure(1, weight=1)

        self._refresh_glossary_view()

    def _refresh_glossary_view(self) -> None:
        for w in self.glossary_scroll.winfo_children():
            w.destroy()
        q = self.var_glossary_search.get().strip().lower()
        row_idx = 0
        for abbr, desc in sorted(self.engine.glossary.items()):
            if q and q not in abbr.lower() and q not in desc.lower():
                continue
            ctk.CTkLabel(
                self.glossary_scroll,
                text=abbr,
                fg_color="#1e3a8a",
                corner_radius=6,
                width=85,
                font=ctk.CTkFont(size=13, weight="bold"),
            ).grid(row=row_idx, column=0, padx=10, pady=5, sticky="w")
            ctk.CTkLabel(
                self.glossary_scroll,
                text=desc,
                font=ctk.CTkFont(size=13),
                anchor="w",
                justify="left",
                wraplength=880,
            ).grid(row=row_idx, column=1, padx=10, pady=5, sticky="ew")
            row_idx += 1

    def _configure_treeview_style(self) -> None:
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Treeview",
            background="#1e222a",
            foreground="#f1f5f9",
            fieldbackground="#1e222a",
            rowheight=28,
            borderwidth=0,
            font=("Segoe UI", 10),
        )
        style.configure(
            "Treeview.Heading",
            background="#334155",
            foreground="#f8fafc",
            font=("Segoe UI", 10, "bold"),
            relief="flat",
        )
        style.map("Treeview", background=[("selected", "#2563eb")], foreground=[("selected", "#ffffff")])

    # ==========================================================================
    # LÓGICA DE NAVEGAÇÃO E ATUALIZAÇÃO DA UI
    # ==========================================================================
    def refresh_key_view(self) -> None:
        # 1. Atualizar badges taxonômicos na Sidebar
        for k, lbl in self.tax_labels.items():
            val = self.engine.current_taxonomy.get(k, "")
            lbl.configure(text=val if val else "—", text_color=ACCENT_PRIMARY if val else "#94a3b8")

        # 2. Atualizar Breadcrumb
        for w in self.breadcrumb_frame.winfo_children():
            w.destroy()

        if not self.engine.history:
            ctk.CTkLabel(
                self.breadcrumb_frame,
                text="Trilha de decisão: Início da chave (nenhum passo respondido ainda)",
                font=ctk.CTkFont(size=12),
                text_color=TEXT_MUTED,
            ).pack(side="left", padx=8, pady=4)
        else:
            for idx, entry in enumerate(self.engine.history):
                btn = ctk.CTkButton(
                    self.breadcrumb_frame,
                    text=f"{entry.step_id} ({entry.chosen_option})",
                    width=64,
                    height=26,
                    fg_color="#334155",
                    hover_color="#475569",
                    font=ctk.CTkFont(size=11, weight="bold"),
                    command=lambda i=idx: self.on_jump_history(i),
                )
                btn.pack(side="left", padx=3, pady=4)

        self.btn_back.configure(state="normal" if self.engine.history else "disabled")

        # 3. Verificar se a chave chegou ao fim
        if self.engine.is_finished or not self.engine.current_step:
            self.card_a["frame"].grid_forget()
            self.card_b["frame"].grid_forget()
            self.finish_panel.grid(row=0, column=0, columnspan=2, padx=12, pady=12, sticky="nsew")

            tax = self.engine.current_taxonomy
            details = (
                f"Espécime Nº: {self.var_mosquito_id.get()}\n\n"
                f"Subfamília:  {tax.get('subfamilia') or '—'}\n"
                f"Tribo:       {tax.get('tribo') or '—'}\n"
                f"Gênero:      {tax.get('genero') or '—'}\n"
                f"Subgênero:   {tax.get('subgenero') or '—'}\n"
                f"Espécie:     {tax.get('especie') or '—'}"
            )
            self.lbl_finish_details.configure(text=details)
            trail = self.engine.get_trail_string()
            self.lbl_finish_trail.configure(text=f"Trilha percorrida ({len(self.engine.history)} passos): {trail}")
            self.lbl_glossary_hints.configure(text="Identificação finalizada. Clique em 'Salvar e Identificar Próximo Mosquito' para gravar no banco.")
            return

        # 4. Exibir passo dicotômico ativo
        self.finish_panel.grid_forget()
        self.card_a["frame"].grid(row=0, column=0, padx=6, pady=4, sticky="nsew")
        self.card_b["frame"].grid(row=0, column=1, padx=6, pady=4, sticky="nsew")

        step: KeyStep = self.engine.current_step
        self.lbl_step_context.configure(text=f"Hierarquia: {step.taxon_context or 'Culicidae'}")
        self.lbl_step_title.configure(text=f"[{step.id}]  {step.title}")

        self._populate_option_card(self.card_a, step.option_a)
        self._populate_option_card(self.card_b, step.option_b)

        # 5. Atualizar dicas do glossário para os termos da pergunta atual
        combined_text = f"{step.option_a.text} {step.option_b.text}"
        terms = self.engine.extract_glossary_terms(combined_text)
        if terms:
            hints_str = "💡 Glossário contextual:   " + "   |   ".join(f"{k}: {v}" for k, v in terms[:4])
        else:
            hints_str = "💡 Dica: Use as teclas [A] e [B] para navegar rapidamente pelas dicotomias, ou [Backspace] para voltar."
        self.lbl_glossary_hints.configure(text=hints_str)

    def _populate_option_card(self, card_widgets: dict, opt: KeyOption) -> None:
        # Badge de destino
        if opt.is_terminal:
            taxon_reached = (
                opt.set_taxonomy.get("especie")
                or opt.set_taxonomy.get("genero")
                or opt.set_taxonomy.get("tribo")
                or opt.set_taxonomy.get("subfamilia")
                or "Fim"
            )
            card_widgets["target_badge"].configure(text=f"🎯 Identifica: {taxon_reached}", text_color=ACCENT_PRIMARY)
        else:
            extra_tax = ", ".join(v for v in opt.set_taxonomy.values() if v)
            prefix = f"({extra_tax}) " if extra_tax else ""
            card_widgets["target_badge"].configure(text=f"{prefix}➔ Vai para {opt.next_step}", text_color="#93c5fd")

        # Texto descritivo
        tb: ctk.CTkTextbox = card_widgets["textbox"]
        tb.configure(state="normal")
        tb.delete("1.0", "end")
        tb.insert("1.0", opt.text)
        tb.configure(state="disabled")

        # Imagem
        img_file = opt.image or "ArboTeamPng.png"
        card_widgets["current_image_file"] = img_file
        ctk_img = self._load_ctk_image(img_file, size=(210, 210))
        if ctk_img:
            card_widgets["img_label"].configure(image=ctk_img, text="")

    def on_choose_option(self, opt_letter: str) -> None:
        self.engine.choose_option(opt_letter)
        self.refresh_key_view()

    def on_back_step(self) -> None:
        if self.engine.go_back():
            self.refresh_key_view()

    def on_jump_history(self, index: int) -> None:
        if self.engine.jump_to_history(index):
            self.refresh_key_view()

    def on_reset_key(self) -> None:
        self.combo_quick_genus.set("Início (A1 - Subfamílias)")
        self.engine.reset()
        self.refresh_key_view()

    def on_finish_current_level(self) -> None:
        if not any(self.engine.current_taxonomy.values()):
            messagebox.showwarning("Aviso", "Ainda não há nenhum nível taxonômico (Subfamília/Tribo/Gênero) definido.")
            return
        self.engine.finish_at_current_level()
        self.refresh_key_view()

    def _on_quick_genus_jump(self, choice: str) -> None:
        if choice.startswith("Início"):
            self.engine.reset()
        else:
            self.engine.start_from_genus(choice)
        self.refresh_key_view()

    # ==========================================================================
    # CATÁLOGO & PERSISTÊNCIA DE REGISTROS
    # ==========================================================================
    def refresh_catalog_list(self) -> None:
        for item in self.catalog_tree.get_children():
            self.catalog_tree.delete(item)

        gen_filter = self.var_catalog_genus.get()
        q = self.var_catalog_search.get().strip().lower()

        for sp, hier in sorted(self.engine.species_hierarchy.items()):
            if gen_filter != "Todos os Gêneros" and hier.get("genero") != gen_filter:
                continue
            haystack = " ".join(hier.values()).lower()
            if q and q not in haystack:
                continue
            self.catalog_tree.insert(
                "",
                "end",
                values=(
                    hier.get("subfamilia", ""),
                    hier.get("tribo", ""),
                    hier.get("genero", ""),
                    hier.get("subgenero", ""),
                    hier.get("especie", ""),
                ),
            )

    def _open_selected_genus_key_from_catalog(self) -> None:
        sel = self.catalog_tree.selection()
        genero = ""
        if sel:
            vals = self.catalog_tree.item(sel[0], "values")
            if len(vals) >= 3:
                genero = vals[2]
        if not genero and self.var_catalog_genus.get() != "Todos os Gêneros":
            genero = self.var_catalog_genus.get()
        if not genero:
            messagebox.showinfo("Selecionar Gênero", "Selecione uma linha na tabela ou escolha um gênero no filtro acima.")
            return
        self.engine.start_from_genus(genero)
        self.combo_quick_genus.set(genero)
        self.tabview.set("🔬 Chave Dicotômica")
        self.refresh_key_view()

    def _select_taxon_from_catalog(self, auto_save: bool = False) -> None:
        sel = self.catalog_tree.selection()
        if not sel:
            messagebox.showwarning("Aviso", "Selecione uma espécie na tabela do catálogo primeiro.")
            return
        vals = self.catalog_tree.item(sel[0], "values")
        self.engine.set_direct_taxon(
            subfamilia=vals[0],
            tribo=vals[1],
            genero=vals[2],
            subgenero=vals[3],
            especie=vals[4],
        )
        self.refresh_key_view()
        if auto_save:
            self.save_and_next_mosquito()
        else:
            self.tabview.set("🔬 Chave Dicotômica")

    def _build_current_record(self) -> SpecimenRecord:
        tax = self.engine.current_taxonomy
        return SpecimenRecord(
            id_mosquito=self.var_mosquito_id.get().strip() or "S/N",
            parque=self.var_park.get().strip(),
            metodo_coleta=self.var_method.get().strip(),
            ponto=self.var_point.get().strip(),
            tempo=self.var_time.get().strip(),
            data=self.var_date.get().strip(),
            coletor=self.var_collector.get().strip(),
            subfamilia=tax.get("subfamilia", ""),
            tribo=tax.get("tribo", ""),
            genero=tax.get("genero", ""),
            subgenero=tax.get("subgenero", ""),
            especie=tax.get("especie", ""),
            trilha_chave=self.engine.get_trail_string(),
        )

    def save_current_specimen(self) -> bool:
        if not any(self.engine.current_taxonomy.values()):
            messagebox.showwarning("Identificação Vazia", "Realize pelo menos um passo da identificação antes de salvar.")
            return False
        rec = self._build_current_record()
        ok, msg = self.repository.save_record(rec, sync_excel=True, sync_sql_server=self.var_sync_sql.get())
        if ok:
            self.refresh_records_table()
            if self.var_auto_increment.get() and rec.id_mosquito.isdigit():
                self.var_mosquito_id.set(str(int(rec.id_mosquito) + 1))
            messagebox.showinfo("Salvo com Sucesso", f"Mosquito #{rec.id_mosquito} ({rec.especie or rec.genero or rec.tribo or rec.subfamilia}) registrado!\n{msg}")
            return True
        else:
            messagebox.showerror("Erro ao Salvar", msg)
            return False

    def save_and_next_mosquito(self) -> None:
        if self.save_current_specimen():
            self.on_reset_key()

    def refresh_records_table(self) -> None:
        for item in self.records_tree.get_children():
            self.records_tree.delete(item)
        rows = self.repository.list_records(self.var_record_search.get())
        self.lbl_records_count.configure(text=f"{len(rows)} registro(s)")
        for r in rows:
            self.records_tree.insert(
                "",
                "end",
                values=(
                    r.get("pk", ""),
                    r.get("id_mosquito", ""),
                    r.get("parque", ""),
                    r.get("metodo_coleta", ""),
                    r.get("ponto", ""),
                    r.get("tempo", ""),
                    r.get("data", ""),
                    r.get("coletor", ""),
                    r.get("subfamilia", ""),
                    r.get("tribo", ""),
                    r.get("genero", ""),
                    r.get("subgenero", ""),
                    r.get("especie", ""),
                    r.get("trilha_chave", ""),
                ),
            )

    def _delete_selected_record(self) -> None:
        sel = self.records_tree.selection()
        if not sel:
            messagebox.showwarning("Aviso", "Selecione um registro na tabela para excluir.")
            return
        vals = self.records_tree.item(sel[0], "values")
        pk = int(vals[0])
        if messagebox.askyesno("Confirmar Exclusão", f"Deseja realmente excluir o registro #{pk} (Mosquito {vals[1]})?"):
            self.repository.delete_record(pk)
            self.refresh_records_table()

    def _export_excel_dialog(self) -> None:
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Planilha Excel", "*.xlsx")],
            initialfile="dados_mosquitos_export.xlsx",
        )
        if path:
            self.repository.export_to_excel(path)
            messagebox.showinfo("Exportação Concluída", f"Planilha Excel exportada para:\n{path}")

    def _export_csv_dialog(self) -> None:
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("Arquivo CSV", "*.csv")],
            initialfile="dados_mosquitos_export.csv",
        )
        if path:
            self.repository.export_to_csv(path)
            messagebox.showinfo("Exportação Concluída", f"Arquivo CSV exportado para:\n{path}")

    # ==========================================================================
    # MODAIS: ZOOM DE PRANCHA MORFOLÓGICA E CONFIGURAÇÃO SQL SERVER
    # ==========================================================================
    def _open_image_zoom_modal(self, opt_letter: str) -> None:
        card = self.card_a if opt_letter == "A" else self.card_b
        img_file = card.get("current_image_file", "ArboTeamPng.png")
        step = self.engine.current_step
        opt_text = (step.option_a.text if opt_letter == "A" else step.option_b.text) if step else ""

        modal = ctk.CTkToplevel(self)
        modal.title(f"Prancha Morfológica — Passo {self.engine.current_step_id} (Opção {opt_letter})")
        modal.geometry("760x720")
        modal.transient(self)
        modal.grab_set()

        ctk_img = self._load_ctk_image(img_file, size=(520, 520))
        if ctk_img:
            ctk.CTkLabel(modal, image=ctk_img, text="").pack(pady=14)

        ctk.CTkLabel(
            modal,
            text=f"Opção {opt_letter}: {opt_text}",
            font=ctk.CTkFont(size=14),
            wraplength=700,
            justify="left",
        ).pack(padx=20, pady=10)

    def _open_sql_server_modal(self) -> None:
        modal = ctk.CTkToplevel(self)
        modal.title("Configurar Sincronização com SQL Server (pyodbc)")
        modal.geometry("440x340")
        modal.transient(self)
        modal.grab_set()

        v_host = tk.StringVar(value=self.repository.db_host)
        v_db = tk.StringVar(value=self.repository.db_name)
        v_tbl = tk.StringVar(value=self.repository.db_table)

        for lbl, var in [("Servidor (Host):", v_host), ("Nome do Banco:", v_db), ("Nome da Tabela:", v_tbl)]:
            ctk.CTkLabel(modal, text=lbl, font=ctk.CTkFont(size=13)).pack(anchor="w", padx=24, pady=(10, 2))
            ctk.CTkEntry(modal, textvariable=var, width=380).pack(padx=24)

        ctk.CTkCheckBox(
            modal,
            text="Sincronizar automaticamente com SQL Server ao salvar",
            variable=self.var_sync_sql,
        ).pack(anchor="w", padx=24, pady=14)

        def _test_and_save():
            ok, msg = self.repository.configure_sql_server(v_host.get(), v_db.get(), v_tbl.get())
            if ok:
                self.var_sync_sql.set(True)
                messagebox.showinfo("SQL Server", msg, parent=modal)
                modal.destroy()
            else:
                messagebox.showerror("Erro SQL Server", msg, parent=modal)

        ctk.CTkButton(
            modal,
            text="Testar Conexão e Salvar",
            fg_color=ACCENT_PRIMARY,
            hover_color=ACCENT_PRIMARY_HOVER,
            command=_test_and_save,
        ).pack(pady=10)

    def _bind_keyboard_shortcuts(self) -> None:
        def _is_typing_in_entry(event) -> bool:
            w = event.widget
            return isinstance(w, (tk.Entry, tk.Text, ctk.CTkEntry, ctk.CTkTextbox))

        def _handle_key_a(event):
            if not _is_typing_in_entry(event) and self.tabview.get() == "🔬 Chave Dicotômica" and not self.engine.is_finished:
                self.on_choose_option("A")

        def _handle_key_b(event):
            if not _is_typing_in_entry(event) and self.tabview.get() == "🔬 Chave Dicotômica" and not self.engine.is_finished:
                self.on_choose_option("B")

        def _handle_back(event):
            if not _is_typing_in_entry(event) and self.tabview.get() == "🔬 Chave Dicotômica":
                self.on_back_step()

        self.bind("<KeyPress-a>", _handle_key_a)
        self.bind("<KeyPress-A>", _handle_key_a)
        self.bind("<KeyPress-1>", _handle_key_a)
        self.bind("<KeyPress-b>", _handle_key_b)
        self.bind("<KeyPress-B>", _handle_key_b)
        self.bind("<KeyPress-2>", _handle_key_b)
        self.bind("<BackSpace>", _handle_back)
        self.bind("<Control-s>", lambda e: self.save_current_specimen())
