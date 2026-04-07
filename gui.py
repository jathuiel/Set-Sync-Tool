"""
gui.py
========================
Interface gráfica do RuleSync Visual — implementada com CustomTkinter.

Autor: Jathuiel Corrêa / JSC Tecnologia
"""

import json
import threading
import customtkinter as ctk
from tkinter import filedialog
from pathlib import Path

from RuleSync_Visual import read_excel, build_xml, XmlVersion, UserSettings


# ══════════════════════════════════════════════════════════════════════════════
#  TOKENS DE DESIGN — edite aqui para trocar o tema completo
# ══════════════════════════════════════════════════════════════════════════════

APP_NAME     = "Rules Sync Visual"
APP_VERSION  = "v1.1"
APP_SUBTITLE = "Excel → Navisworks Appearance Profiler XML"

WINDOW_W, WINDOW_H = 600, 700

# Modo de aparência: "dark" | "light" | "system"
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")


def load_theme_definition(theme_name: str) -> dict:
    path = Path(__file__).with_name(f"theme-{theme_name}.json")
    if path.exists():
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            return {
                "base": data.get("background", "#f6f6f6"),
                "surface": data.get("background", "#f6f6f6"),
                "surface2": data.get("inputBackground", "#F0F0F0EA"),
                "border": data.get("accent", "#0078D4"),
                "accent": data.get("accent", "#FF453A"),
                "accent_hv": data.get("accentHover", "#FF9F0A"),
                "accent2": data.get("accent", "#0078D4"),
                "warn": "#FF9F0A",
                "err": "#FF453A",
                "fg": data.get("text", "#292929"),
                "fg2": data.get("text", "#292929"),
            }
        except Exception:
            pass
    # fallback dark hardcoded
    return {
        "base": "#292929",
        "surface": "#2C2C2E",
        "surface2": "#3A3A3C",
        "border": "#3A3A3C",
        "accent": "#E74504",
        "accent_hv": "#C93A03",
        "accent2": "#30D158",
        "warn": "#FF9F0A",
        "err": "#FF453A",
        "fg": "#F5ECE4",
        "fg2": "#8E8E93",
    }

# Paleta JSC — fosco escuro + acento laranja
C = {
    "base":      "#292929",   # Fundo raiz da janela
    "surface":   "#2C2C2E",   # Card principal
    "surface2":  "#3A3A3C",   # Inputs e log
    "border":    "#3A3A3C",   # Bordas sutis
    "accent":    "#E74504",   # Laranja JSC
    "accent_hv": "#C93A03",   # Laranja escurecido (hover)
    "accent2":   "#30D158",   # Verde sucesso
    "warn":      "#FF9F0A",   # Amarelo aviso
    "err":       "#FF453A",   # Vermelho erro
    "fg":        "#F5ECE4",   # Texto primário
    "fg2":       "#8E8E93",   # Texto secundário
}

# Raios de canto (pixels)
R_CARD   = 14   # Card principal
R_INPUT  = 10   # Campos de entrada e botões
R_LOG    = 12   # Painel de log

# Tamanhos e famílias de fonte (para usar dentro de CTkFont)
FONT_SIZE_TITLE = 18
FONT_SIZE_BODY  = 14
FONT_SIZE_LABEL = 12
FONT_SIZE_MONO  = 12
FONT_FAMILY_DEFAULT = "Segoe UI"
FONT_FAMILY_MONO    = "Consolas"

PAD             = 20
DEFAULT_OUTPUT  = "Appearance_Profiler.xml"
DEFAULT_PROFILE = "Status de '...'"


# ══════════════════════════════════════════════════════════════════════════════
#  CLASSE PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════

class App(ctk.CTk):
    """Janela principal do RuleSync Visual."""

    def __init__(self):
        super().__init__()

        self.settings = UserSettings.load()
        self._set_theme_data(self.settings.theme)

        self.title(f"{APP_NAME} {APP_VERSION}")
        self.resizable(False, False)

        # Criar fontes após janela raiz existir
        self.F_TITLE = ctk.CTkFont(family=FONT_FAMILY_DEFAULT, size=FONT_SIZE_TITLE, weight="bold")
        self.F_BODY  = ctk.CTkFont(family=FONT_FAMILY_DEFAULT, size=FONT_SIZE_BODY)
        self.F_LABEL = ctk.CTkFont(family=FONT_FAMILY_DEFAULT, size=FONT_SIZE_LABEL, weight="bold")
        self.F_MONO  = ctk.CTkFont(family=FONT_FAMILY_MONO, size=FONT_SIZE_MONO)

        self._build_header()
        self._build_card()
        self._build_log()
        self._center()

    def _set_theme_data(self, theme_name: str):
        theme_name = theme_name if theme_name in ("dark", "light") else "dark"
        self.settings.theme = theme_name
        self.settings.save()

        ctk.set_appearance_mode(theme_name)
        theme_colors = load_theme_definition(theme_name)
        C.update(theme_colors)
        self.configure(fg_color=C["base"])

    def _on_theme_switch(self):
        new_theme = "light" if self.var_theme_mode.get() else "dark"
        self._set_theme_data(new_theme)
        self._rebuild_ui()

    def _rebuild_ui(self):
        stored_values = {
            "template": self.var_template.get(),
            "output": self.var_output.get(),
            "profile": self.var_profile.get(),
            "xml_version": self.var_xmlversion.get(),
            "theme": self.var_theme_mode.get(),
        }

        # Reconstrói a UI para aplicar novas cores no modo dinâmico.
        for child in self.winfo_children():
            child.destroy()

        self._build_header()
        self._build_card()
        self._build_log()
        self._center()

        self.var_template.set(stored_values["template"])
        self.var_output.set(stored_values["output"])
        self.var_profile.set(stored_values["profile"])
        self.var_xmlversion.set(stored_values["xml_version"])
        self.var_theme_mode.set(stored_values["theme"])

    # ── Centralização ─────────────────────────────────────────────────────────

    def _center(self):
        """Posiciona a janela no centro da tela."""
        self.update_idletasks()
        x = (self.winfo_screenwidth()  - WINDOW_W) // 2
        y = (self.winfo_screenheight() - WINDOW_H) // 2
        self.geometry(f"{WINDOW_W}x{WINDOW_H}+{x}+{y}")

    # ── Cabeçalho ─────────────────────────────────────────────────────────────

    def _build_header(self):
        """Linha de acento + nome + subtítulo."""
        # Linha fina de acento laranja JSC (6px, sem raio)
        ctk.CTkFrame(self, height=6, fg_color=C["accent"],
                     corner_radius=0).pack(fill="x")

        header = ctk.CTkFrame(self, fg_color=C["base"], corner_radius=0)
        header.pack(fill="x", padx=PAD, pady=(18, 0))

        ctk.CTkLabel(header, text=APP_NAME, font=self.F_TITLE,
                     text_color=C["fg"]).pack(anchor="w")
        ctk.CTkLabel(header, text=APP_SUBTITLE, font=self.F_BODY,
                     text_color=C["fg2"]).pack(anchor="w", pady=(2, 14))

        # Divisor
        ctk.CTkFrame(self, height=1, fg_color=C["border"],
                     corner_radius=0).pack(fill="x", padx=PAD)

    # ── Card de formulário ────────────────────────────────────────────────────

    def _build_card(self):
        """Card com cantos arredondados reais contendo os campos de entrada."""
        card = ctk.CTkFrame(self, fg_color=C["surface"],
                            corner_radius=R_CARD,
                            border_width=1, border_color=C["border"])
        card.pack(fill="x", padx=PAD, pady=PAD)

        inner = ctk.CTkFrame(card, fg_color=C["surface"], corner_radius=0)
        inner.pack(fill="x", padx=PAD, pady=PAD)

        # Template Excel
        self.var_template = ctk.StringVar()
        self._field_label(inner, "Template Excel (.xlsx)")
        self._entry_row(inner, self.var_template, "Procurar", self._browse_template)

        # Arquivo de saída
        self.var_output = ctk.StringVar(value=DEFAULT_OUTPUT)
        self._field_label(inner, "Arquivo de saída (.xml)")
        self._entry_row(inner, self.var_output, "Salvar em", self._browse_output)

        # Nome do perfil
        self.var_profile = ctk.StringVar(value=DEFAULT_PROFILE)
        self._field_label(inner, "Nome do Perfil")
        ctk.CTkEntry(inner, textvariable=self.var_profile, width=260,
                     fg_color=C["surface2"], border_color=C["border"],
                     text_color=C["fg"], font=self.F_BODY,
                     corner_radius=R_INPUT).pack(anchor="w", pady=(0, 4))

        # Versão XML
        self.var_xmlversion = ctk.StringVar(value=self.settings.xml_version)
        self._field_label(inner, "Versão XML")
        ctk.CTkOptionMenu(inner,
                          values=[XmlVersion.V1.value, XmlVersion.V2.value],
                          variable=self.var_xmlversion,
                          fg_color=C["surface2"],
                          button_color=C["surface2"],
                          button_hover_color=C["border"],
                          text_color=C["fg"],
                          dropdown_fg_color=C["surface"],
                          dropdown_text_color=C["fg"],
                          font=self.F_BODY,
                          corner_radius=R_INPUT).pack(anchor="w", pady=(0, 8))

        # Alternador de tema
        self.var_theme_mode = ctk.BooleanVar(value=(self.settings.theme == "light"))
        ctk.CTkSwitch(inner,
                      text="Modo claro",
                      variable=self.var_theme_mode,
                      onvalue=True,
                      offvalue=False,
                      command=self._on_theme_switch,
                      progress_color=C["accent"],
                      button_color=C["surface"],
                      fg_color=C["surface2"],
                      text_color=C["fg"],
                      font=self.F_BODY).pack(anchor="w", pady=(0, 12))

        # Botão exportar
        ctk.CTkButton(inner,
                      text="Exportar XML",
                      font=self.F_LABEL,
                      fg_color=C["accent"],
                      hover_color=C["accent_hv"],
                      text_color=C["fg"],
                      corner_radius=R_INPUT,
                      command=self._run_export).pack(anchor="e", pady=(18, 0))

    # ── Painel de log ─────────────────────────────────────────────────────────

    def _build_log(self):
        """Área de feedback com fonte mono e cantos arredondados nativos."""
        frame = ctk.CTkFrame(self, fg_color=C["base"], corner_radius=0)
        frame.pack(fill="both", expand=True, padx=PAD, pady=(0, PAD))

        self.log = ctk.CTkTextbox(
            frame,
            fg_color=C["surface2"],
            border_color=C["border"],
            border_width=1,
            text_color=C["fg2"],
            font=self.F_MONO,
            corner_radius=R_LOG,
            wrap="word",
            state="disabled")
        self.log.pack(fill="both", expand=True)

        # Tags de cor para os níveis de mensagem
        self.log.tag_config("ok",   foreground=C["accent2"])
        self.log.tag_config("warn", foreground=C["warn"])
        self.log.tag_config("err",  foreground=C["err"])

    # ── Helpers de widget ─────────────────────────────────────────────────────

    def _field_label(self, parent, text: str):
        """Rótulo de campo em uppercase, estilo iOS."""
        ctk.CTkLabel(parent, text=text.upper(), font=self.F_LABEL,
                     text_color=C["fg2"]).pack(anchor="w", pady=(12, 3))

    def _entry_row(self, parent, var, btn_text: str, btn_cmd):
        """Linha composta: entrada de texto + botão lateral."""
        row = ctk.CTkFrame(parent, fg_color=C["surface"], corner_radius=0)
        row.pack(fill="x", pady=(0, 2))

        ctk.CTkEntry(row, textvariable=var,
                     fg_color=C["surface2"], border_color=C["border"],
                     text_color=C["fg"], font=self.F_BODY,
                     corner_radius=R_INPUT).pack(side="left", fill="x", expand=True)

        ctk.CTkButton(row, text=btn_text, width=90,
                      fg_color=C["surface2"], hover_color=C["border"],
                      text_color=C["fg"], font=self.F_BODY,
                      border_width=1, border_color=C["border"],
                      corner_radius=R_INPUT,
                      command=btn_cmd).pack(side="left", padx=(8, 0))

    # ── Diálogos de arquivo ───────────────────────────────────────────────────

    def _browse_template(self):
        """Seleciona o Excel e sugere pasta de saída automaticamente."""
        path = filedialog.askopenfilename(
            title="Selecionar template Excel",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("Todos", "*.*")])
        if path:
            self.var_template.set(path)
            if self.var_output.get() in ("", DEFAULT_OUTPUT):
                self.var_output.set(str(Path(path).parent / DEFAULT_OUTPUT))

    def _browse_output(self):
        """Define o caminho e nome do XML de saída."""
        path = filedialog.asksaveasfilename(
            title="Salvar XML como",
            defaultextension=".xml",
            filetypes=[("XML", "*.xml"), ("Todos", "*.*")],
            initialfile=DEFAULT_OUTPUT)
        if path:
            self.var_output.set(path)

    # ── Log helpers ───────────────────────────────────────────────────────────

    def _log(self, msg: str, tag: str = ""):
        """Insere linha no log com tag de cor opcional."""
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n", tag)
        self.log.see("end")
        self.log.configure(state="disabled")

    def _clear_log(self):
        """Limpa o log antes de cada nova exportação."""
        self.log.configure(state="normal")
        self.log.delete("0.0", "end")
        self.log.configure(state="disabled")

    # ── Exportação ────────────────────────────────────────────────────────────

    def _run_export(self):
        """Valida campos e dispara exportação em thread separada."""
        tmpl = self.var_template.get().strip()
        out  = self.var_output.get().strip()
        prof = self.var_profile.get().strip() or DEFAULT_PROFILE
        xml_version = self.var_xmlversion.get().strip() or XmlVersion.V2.value
        self.settings.xml_version = xml_version
        self.settings.save()

        if not tmpl:
            return self._log("Erro: selecione o template Excel.", "err")
        if not Path(tmpl).exists():
            return self._log(f"Erro: arquivo não encontrado — {tmpl}", "err")
        if not out:
            return self._log("Erro: defina o caminho do arquivo XML.", "err")

        self._clear_log()
        threading.Thread(target=self._export_worker,
                         args=(tmpl, out, prof), daemon=True).start()

    def _export_worker(self, template: str, output: str, profile: str):
        """Worker em thread — atualiza a UI via self.after() com segurança."""
        def log(msg, tag=""): self.after(0, self._log, msg, tag)
        try:
            log(f"Lendo: {Path(template).name}")
            data = read_excel(template)
            log(f"  {len(data)} regras encontradas.", "ok")

            log(f"Perfil: {profile}")
            xml_content = build_xml(data, profile, xml_version=self.settings.xml_version)

            with open(output, "w", encoding="utf-8") as f:
                f.write(xml_content)

            log(f"Exportado: {Path(output).name}", "ok")
            log(f"Caminho: {output}")
            log(f"Concluído — {len(data)} regras exportadas.", "ok")

        except KeyError as e:
            log(f"Aba não encontrada no Excel: {e}", "err")
        except Exception as e:
            log(f"Erro inesperado: {e}", "err")