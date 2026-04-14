"""
gui.py
========================
Interface grafica principal para exportacao de XML do Navisworks.

- Aba 1: Appearance Profiler
- Aba 2: Search Sets
"""

from __future__ import annotations

import threading
from pathlib import Path
from tkinter import filedialog

import customtkinter as ctk

from RuleSync_Visual import build_xml as build_profiler_xml
from RuleSync_Visual import read_excel as read_profiler_excel
from excel_to_nw_search import convert_excel_to_xml
from excel_to_nw_search import read_config as read_search_config
from excel_to_nw_search import read_selection_sets


APP_NAME = "Sets Sync Tool"
APP_VERSION = "v1.1"
APP_SUBTITLE = "Excel para Navisworks XML"

WINDOW_W, WINDOW_H = 720, 780

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

# Paleta centralizada para manter o visual consistente entre as duas abas.
C = {
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

R_CARD = 14
R_INPUT = 10
R_LOG = 12

FONT_SIZE_TITLE = 18
FONT_SIZE_BODY = 14
FONT_SIZE_LABEL = 12
FONT_SIZE_MONO = 12
FONT_FAMILY_DEFAULT = "Segoe UI"
FONT_FAMILY_MONO = "Consolas"

PAD = 20
DEFAULT_PROFILER_OUTPUT = "Profiler_Appearance.xml"
DEFAULT_PROFILE = "Status de Aparencia"
DEFAULT_SEARCH_OUTPUT = "Set_Search.xml"


class App(ctk.CTk):
    """Janela principal com dois fluxos de exportacao."""

    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} {APP_VERSION}")
        self.resizable(False, False)
        self.configure(fg_color=C["base"])

        self.F_TITLE = ctk.CTkFont(
            family=FONT_FAMILY_DEFAULT, size=FONT_SIZE_TITLE, weight="bold"
        )
        self.F_BODY = ctk.CTkFont(family=FONT_FAMILY_DEFAULT, size=FONT_SIZE_BODY)
        self.F_LABEL = ctk.CTkFont(
            family=FONT_FAMILY_DEFAULT, size=FONT_SIZE_LABEL, weight="bold"
        )
        self.F_MONO = ctk.CTkFont(family=FONT_FAMILY_MONO, size=FONT_SIZE_MONO)

        self._build_header()
        self._build_tabs()
        self._center()

    def _center(self):
        """Centraliza a janela usando as dimensoes fixas definidas no modulo."""
        self.update_idletasks()
        x = (self.winfo_screenwidth() - WINDOW_W) // 2
        y = (self.winfo_screenheight() - WINDOW_H) // 2
        self.geometry(f"{WINDOW_W}x{WINDOW_H}+{x}+{y}")

    def _build_header(self):
        """Monta o topo comum da aplicacao, compartilhado pelas duas abas."""
        ctk.CTkFrame(self, height=6, fg_color=C["accent"], corner_radius=0).pack(
            fill="x"
        )

        header = ctk.CTkFrame(self, fg_color=C["base"], corner_radius=0)
        header.pack(fill="x", padx=PAD, pady=(18, 0))

        ctk.CTkLabel(
            header, text=APP_NAME, font=self.F_TITLE, text_color=C["fg"]
        ).pack(anchor="w")
        ctk.CTkLabel(
            header, text=APP_SUBTITLE, font=self.F_BODY, text_color=C["fg2"]
        ).pack(anchor="w", pady=(2, 14))

        ctk.CTkFrame(self, height=1, fg_color=C["border"], corner_radius=0).pack(
            fill="x", padx=PAD
        )

    def _build_tabs(self):
        """Cria as abas principais e delega o conteudo para metodos especificos."""
        self.tabview = ctk.CTkTabview(
            self,
            fg_color=C["surface"],
            segmented_button_fg_color=C["surface2"],
            segmented_button_selected_color=C["accent"],
            segmented_button_selected_hover_color=C["accent_hv"],
            segmented_button_unselected_color=C["surface2"],
            segmented_button_unselected_hover_color=C["border"],
            text_color=C["fg"],
            corner_radius=R_CARD,
            border_width=1,
            border_color=C["border"],
        )
        self.tabview.pack(fill="both", expand=True, padx=PAD, pady=PAD)

        profiler_tab = self.tabview.add("Appearance Profiler")
        search_tab = self.tabview.add("Search Sets")

        profiler_tab.configure(fg_color=C["surface"])
        search_tab.configure(fg_color=C["surface"])

        self._build_profiler_tab(profiler_tab)
        self._build_search_tab(search_tab)

    def _build_profiler_tab(self, parent):
        """Monta o formulario e o log do fluxo Appearance Profiler."""
        container = ctk.CTkFrame(parent, fg_color=C["surface"], corner_radius=0)
        container.pack(fill="both", expand=True, padx=PAD, pady=PAD)

        self._build_intro(
            container,
            "Converte a planilha de regras em XML de Appearance Profiler.",
        )

        form = self._build_form_card(container)

        self.var_profiler_template = ctk.StringVar()
        self.var_profiler_output = ctk.StringVar(value=DEFAULT_PROFILER_OUTPUT)
        self.var_profile = ctk.StringVar(value=DEFAULT_PROFILE)

        self._field_label(form, "Template Excel (.xlsx)")
        self._entry_row(
            form,
            self.var_profiler_template,
            "Procurar",
            self._browse_profiler_template,
        )

        self._field_label(form, "Arquivo de saida (.xml)")
        self._entry_row(
            form,
            self.var_profiler_output,
            "Salvar em",
            self._browse_profiler_output,
        )

        self._field_label(form, "Nome do Perfil")
        ctk.CTkEntry(
            form,
            textvariable=self.var_profile,
            fg_color=C["surface2"],
            border_color=C["border"],
            text_color=C["fg"],
            font=self.F_BODY,
            corner_radius=R_INPUT,
        ).pack(fill="x", pady=(0, 4))

        ctk.CTkButton(
            form,
            text="Exportar XML",
            font=self.F_LABEL,
            fg_color=C["accent"],
            hover_color=C["accent_hv"],
            text_color=C["fg"],
            corner_radius=R_INPUT,
            command=self._run_profiler_export,
        ).pack(anchor="e", pady=(18, 0))

        self.log_profiler = self._build_log(container)

    def _build_search_tab(self, parent):
        """Monta o formulario e o log do fluxo Search Sets."""
        container = ctk.CTkFrame(parent, fg_color=C["surface"], corner_radius=0)
        container.pack(fill="both", expand=True, padx=PAD, pady=PAD)

        self._build_intro(
            container,
            "Converte o template de Search Sets em XML no padrao do Navisworks.",
        )

        form = self._build_form_card(container)

        self.var_search_template = ctk.StringVar()
        self.var_search_output = ctk.StringVar(value=DEFAULT_SEARCH_OUTPUT)

        self._field_label(form, "Template Excel (.xlsx)")
        self._entry_row(
            form,
            self.var_search_template,
            "Procurar",
            self._browse_search_template,
        )

        self._field_label(form, "Arquivo de saida (.xml)")
        self._entry_row(
            form,
            self.var_search_output,
            "Salvar em",
            self._browse_search_output,
        )

        helper = (
            "Lê as abas 'CONFIG' e 'SELECTION_SETS' e gera selection sets com 'GUID' baseados na coluna 'ITEM_ID'. O XML gerado pode ser importado diretamente no Navisworks."
        )
        ctk.CTkLabel(
            form,
            text=helper,
            font=self.F_BODY,
            text_color=C["fg2"],
            justify="left",
            wraplength=560,
        ).pack(anchor="w", pady=(12, 0))

        ctk.CTkButton(
            form,
            text="Exportar XML",
            font=self.F_LABEL,
            fg_color=C["accent"],
            hover_color=C["accent_hv"],
            text_color=C["fg"],
            corner_radius=R_INPUT,
            command=self._run_search_export,
        ).pack(anchor="e", pady=(18, 0))

        self.log_search = self._build_log(container)

    def _build_intro(self, parent, text: str):
        """Renderiza o texto introdutorio usado no topo de cada aba."""
        ctk.CTkLabel(
            parent,
            text=text,
            font=self.F_BODY,
            text_color=C["fg2"],
            justify="left",
            wraplength=620,
        ).pack(anchor="w", pady=(0, 12))

    def _build_form_card(self, parent):
        """Cria um container visual padrao para os campos do formulario."""
        card = ctk.CTkFrame(
            parent,
            fg_color=C["surface"],
            corner_radius=R_CARD,
            border_width=1,
            border_color=C["border"],
        )
        card.pack(fill="x")

        inner = ctk.CTkFrame(card, fg_color=C["surface"], corner_radius=0)
        inner.pack(fill="x", padx=PAD, pady=PAD)
        return inner

    def _build_log(self, parent):
        """Cria a area de log com tags visuais para sucesso, alerta e erro."""
        frame = ctk.CTkFrame(parent, fg_color=C["base"], corner_radius=0)
        frame.pack(fill="both", expand=True, pady=(16, 0))

        log = ctk.CTkTextbox(
            frame,
            fg_color=C["surface2"],
            border_color=C["border"],
            border_width=1,
            text_color=C["fg2"],
            font=self.F_MONO,
            corner_radius=R_LOG,
            wrap="word",
            state="disabled",
        )
        log.pack(fill="both", expand=True)
        log.tag_config("ok", foreground=C["accent2"])
        log.tag_config("warn", foreground=C["warn"])
        log.tag_config("err", foreground=C["err"])
        return log

    def _field_label(self, parent, text: str):
        """Padroniza os titulos dos campos para evitar repeticao de estilo."""
        ctk.CTkLabel(
            parent,
            text=text.upper(),
            font=self.F_LABEL,
            text_color=C["fg2"],
        ).pack(anchor="w", pady=(12, 3))

    def _entry_row(self, parent, variable, button_text: str, button_cmd):
        """Cria a linha composta por entrada + botao auxiliar (browse/save)."""
        row = ctk.CTkFrame(parent, fg_color=C["surface"], corner_radius=0)
        row.pack(fill="x", pady=(0, 2))

        ctk.CTkEntry(
            row,
            textvariable=variable,
            fg_color=C["surface2"],
            border_color=C["border"],
            text_color=C["fg"],
            font=self.F_BODY,
            corner_radius=R_INPUT,
        ).pack(side="left", fill="x", expand=True)

        ctk.CTkButton(
            row,
            text=button_text,
            width=90,
            fg_color=C["surface2"],
            hover_color=C["border"],
            text_color=C["fg"],
            font=self.F_BODY,
            border_width=1,
            border_color=C["border"],
            corner_radius=R_INPUT,
            command=button_cmd,
        ).pack(side="left", padx=(8, 0))

    def _log(self, widget, msg: str, tag: str = ""):
        """Escreve uma nova linha no log e mantem o scroll no final."""
        widget.configure(state="normal")
        widget.insert("end", msg + "\n", tag)
        widget.see("end")
        widget.configure(state="disabled")

    def _clear_log(self, widget):
        """Limpa o log antes de iniciar uma nova exportacao."""
        widget.configure(state="normal")
        widget.delete("0.0", "end")
        widget.configure(state="disabled")

    def _browse_profiler_template(self):
        """Seleciona o Excel do profiler e sugere o XML no mesmo diretorio."""
        path = filedialog.askopenfilename(
            title="Selecionar template Excel",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("Todos", "*.*")],
        )
        if path:
            self.var_profiler_template.set(path)
            if self.var_profiler_output.get() in ("", DEFAULT_PROFILER_OUTPUT):
                self.var_profiler_output.set(
                    str(Path(path).with_name(DEFAULT_PROFILER_OUTPUT))
                )

    def _browse_profiler_output(self):
        """Permite escolher manualmente onde salvar o XML do profiler."""
        path = filedialog.asksaveasfilename(
            title="Salvar XML como",
            defaultextension=".xml",
            filetypes=[("XML", "*.xml"), ("Todos", "*.*")],
            initialfile=DEFAULT_PROFILER_OUTPUT,
        )
        if path:
            self.var_profiler_output.set(path)

    def _browse_search_template(self):
        """Seleciona o Excel de search set e reaproveita o nome para o XML."""
        path = filedialog.askopenfilename(
            title="Selecionar template Excel de Search Set",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("Todos", "*.*")],
        )
        if path:
            self.var_search_template.set(path)
            current_output = self.var_search_output.get().strip()
            if current_output in ("", DEFAULT_SEARCH_OUTPUT, "nw_search_template_v2.xml"):
                self.var_search_output.set(str(Path(path).with_suffix(".xml")))

    def _browse_search_output(self):
        """Permite escolher manualmente onde salvar o XML de search set."""
        path = filedialog.asksaveasfilename(
            title="Salvar Search Set XML como",
            defaultextension=".xml",
            filetypes=[("XML", "*.xml"), ("Todos", "*.*")],
            initialfile=DEFAULT_SEARCH_OUTPUT,
        )
        if path:
            self.var_search_output.set(path)

    def _run_profiler_export(self):
        """Valida os campos do profiler e dispara a exportacao em background."""
        template = self.var_profiler_template.get().strip()
        output = self.var_profiler_output.get().strip()
        profile = self.var_profile.get().strip() or DEFAULT_PROFILE

        if not template:
            return self._log(self.log_profiler, "Erro: selecione o template Excel.", "err")
        if not Path(template).exists():
            return self._log(
                self.log_profiler,
                f"Erro: arquivo nao encontrado - {template}",
                "err",
            )
        if not output:
            return self._log(
                self.log_profiler,
                "Erro: defina o caminho do arquivo XML.",
                "err",
            )

        self._clear_log(self.log_profiler)
        # A exportacao roda em outra thread para a janela nao travar durante IO/parse.
        threading.Thread(
            target=self._profiler_worker,
            args=(template, output, profile),
            daemon=True,
        ).start()

    def _profiler_worker(self, template: str, output: str, profile: str):
        def log(msg: str, tag: str = ""):
            # Widgets do Tkinter devem ser atualizados na thread principal.
            self.after(0, self._log, self.log_profiler, msg, tag)

        try:
            log(f"Lendo: {Path(template).name}")
            data = read_profiler_excel(template)
            log(f"Regras encontradas: {len(data)}", "ok")
            log(f"Perfil: {profile}")

            xml_content = build_profiler_xml(data, profile)
            Path(output).write_text(xml_content, encoding="utf-8")

            log(f"Exportado: {Path(output).name}", "ok")
            log(f"Caminho: {output}")
            log("Concluido.", "ok")
        except KeyError as exc:
            log(f"Aba nao encontrada no Excel: {exc}", "err")
        except Exception as exc:
            log(f"Erro inesperado: {exc}", "err")

    def _run_search_export(self):
        """Valida os campos do search set e inicia a conversao em background."""
        template = self.var_search_template.get().strip()
        output = self.var_search_output.get().strip()

        if not template:
            return self._log(self.log_search, "Erro: selecione o template Excel.", "err")
        if not Path(template).exists():
            return self._log(
                self.log_search,
                f"Erro: arquivo nao encontrado - {template}",
                "err",
            )
        if not output:
            return self._log(
                self.log_search,
                "Erro: defina o caminho do arquivo XML.",
                "err",
            )

        self._clear_log(self.log_search)
        # A exportacao roda em outra thread para a janela nao travar durante IO/parse.
        threading.Thread(
            target=self._search_worker,
            args=(template, output),
            daemon=True,
        ).start()

    def _search_worker(self, template: str, output: str):
        def log(msg: str, tag: str = ""):
            # Widgets do Tkinter devem ser atualizados na thread principal.
            self.after(0, self._log, self.log_search, msg, tag)

        try:
            log(f"Lendo: {Path(template).name}")
            config = read_search_config(template)
            selectionsets = read_selection_sets(template)
            condition_count = sum(len(items) for items in selectionsets.values())

            log(f"Selection sets: {len(selectionsets)}", "ok")
            log(f"Conditions: {condition_count}", "ok")
            log(
                "Config: "
                f"units={config.get('units', '')}, "
                f"filename={config.get('filename', '')}, "
                f"disjoint={config.get('disjoint', '')}"
            )

            generated_path = convert_excel_to_xml(template, output)

            log(f"Exportado: {generated_path.name}", "ok")
            log(f"Caminho: {generated_path}")
            log("Concluido.", "ok")
        except KeyError as exc:
            log(f"Aba nao encontrada no Excel: {exc}", "err")
        except Exception as exc:
            log(f"Erro inesperado: {exc}", "err")


if __name__ == "__main__":
    app = App()
    app.mainloop()
