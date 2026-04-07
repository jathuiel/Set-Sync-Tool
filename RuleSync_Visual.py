"""
RuleSync_Visual.py
========================
Dependências:
    pip install openpyxl

Autor: Jathuiel Corrêa / JSC Tecnologia
"""

import json
import math
from dataclasses import asdict, dataclass
from enum import Enum
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.dom import minidom

import openpyxl


class XmlVersion(str, Enum):
    V1 = "V1"
    V2 = "V2"


@dataclass
class UserSettings:
    xml_version: str = XmlVersion.V2.value
    theme: str = "dark"

    SETTINGS_FILE = Path(__file__).with_name("user_settings.json")

    @classmethod
    def load(cls):
        try:
            if cls.SETTINGS_FILE.exists():
                raw = cls.SETTINGS_FILE.read_text(encoding="utf-8")
                data = json.loads(raw)
                xml_version = data.get("xml_version", XmlVersion.V2.value)
                theme = data.get("theme", "dark")
                if xml_version not in XmlVersion.__members__ and xml_version not in [v.value for v in XmlVersion]:
                    xml_version = XmlVersion.V2.value
                return cls(xml_version=xml_version, theme=theme)
        except Exception:
            pass
        return cls()

    def save(self):
        try:
            with open(self.SETTINGS_FILE, "w", encoding="utf-8") as f:
                json.dump(asdict(self), f, indent=2, ensure_ascii=False)
        except Exception:
            pass


# ══════════════════════════════════════════════════════════════════════════════
#  SEÇÃO 1 — CONFIGURAÇÃO DO TEMPLATE EXCEL
#  Altere aqui caso mude a estrutura de colunas da aba "Regras"
# ══════════════════════════════════════════════════════════════════════════════

# Nome da aba no Excel que contém as regras de aparência
SHEET_NAME = "Regras"

# Linha onde os dados começam (1-indexed; linha 1=título, 2=cabeçalho)
START_ROW = 3

# Índices de coluna (0-indexed: A=0, B=1, C=2 ...)
COL_SET_NAME = 0    # A – Nome do Set (chave de matching no Navisworks)
COL_CWP      = 1    # B – CWP
COL_GRUPO    = 2    # C – Grupo de Sets
COL_IWP      = 3    # D – IWP / Set Item
COL_TRANSP   = 4    # E – Transparência (0 = opaco, 100 = invisível)
COL_HIDDEN   = 5    # F – Oculto (TRUE / FALSE)
COL_R        = 8    # I – Canal vermelho RGB (0-255)
COL_G        = 9    # J – Canal verde RGB (0-255)
COL_B        = 10   # K – Canal azul RGB (0-255)


# ══════════════════════════════════════════════════════════════════════════════
#  SEÇÃO 2 — LÓGICA DE EXPORTAÇÃO
#  Funções de leitura do Excel e geração do XML.
#  Altere aqui apenas se o formato do XML ou da planilha mudar.
# ══════════════════════════════════════════════════════════════════════════════

def srgb_to_linear(c255: int) -> float:
    """
    Converte um canal de cor de 0-255 para espaço linear (gamma 2.2).
    O Navisworks armazena os canais ScR/ScG/ScB neste espaço.

    Fórmula: (canal / 255) ^ 2.2

    Parâmetros:
        c255 (int): valor inteiro do canal de cor, entre 0 e 255.

    Retorna:
        float: valor linearizado do canal, arredondado a 9 casas decimais.
    """
    v = c255 / 255.0
    if v <= 0:
        return 0.0
    return round(math.pow(v, 2.2), 9)


def read_excel(path: str) -> list[dict]:
    """
    Lê a aba definida em SHEET_NAME a partir de START_ROW.
    Para na primeira linha em que SET NAME estiver vazio.

    Parâmetros:
        path (str): caminho completo para o arquivo Excel (.xlsx / .xlsm).

    Retorna:
        list[dict]: lista de dicionários, um por regra de aparência.
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[SHEET_NAME]
    rows = []

    for row in ws.iter_rows(min_row=START_ROW, values_only=True):
        # Interrompe na primeira linha sem nome de Set
        if not row[COL_SET_NAME]:
            break

        rows.append({
            "set_name":     str(row[COL_SET_NAME]).strip(),
            "cwp":          str(row[COL_CWP]).strip()    if row[COL_CWP]    else "",
            "grupo":        str(row[COL_GRUPO]).strip()  if row[COL_GRUPO]  else "",
            "iwp":          str(row[COL_IWP]).strip()    if row[COL_IWP]    else "",
            "transparency": int(row[COL_TRANSP])         if row[COL_TRANSP] is not None else 0,
            "hidden":       str(row[COL_HIDDEN]).strip() if row[COL_HIDDEN] else "FALSE",
            "R":            int(row[COL_R])              if row[COL_R]      is not None else 0,
            "G":            int(row[COL_G])              if row[COL_G]      is not None else 0,
            "B":            int(row[COL_B])              if row[COL_B]      is not None else 0,
        })

    return rows


def build_xml(rows: list[dict], profile_name: str, xml_version: str = XmlVersion.V2.value) -> str:
    """
    Monta o XML do Appearance Profiler a partir das regras lidas do Excel.

    Estrutura gerada:
        Root
        ├── SingleFilePersistedItemVersion
        │     └── Version: 1
        └── SingleFilePersistedItemAppearanceProfileV2
              ├── Profile
              │     └── PersistedItemAppearanceProfileRuleV1 (uma por IWP)
              │           ├── Transparency
              │           ├── Color (A, R, G, B, ScA, ScR, ScG, ScB)
              │           ├── Hidden
              │           ├── <n>  ← nome do Set
              │           └── SavedItemId  ← CWP + grupo + IWP separados por \\n
              └── <n>  ← nome do perfil

    Parâmetros:
        rows (list[dict]): regras retornadas por read_excel().
        profile_name (str): nome do perfil exibido no Navisworks.

    Retorna:
        str: conteúdo XML formatado, com BOM UTF-8 (\\ufeff) no início.
    """

    if isinstance(xml_version, XmlVersion):
        xml_version = xml_version.value

    if xml_version not in (XmlVersion.V1.value, XmlVersion.V2.value):
        xml_version = XmlVersion.V2.value

    root = ET.Element("Root")

    # Bloco de versão do formato
    ver = ET.SubElement(root, "SingleFilePersistedItemVersion", attrib={
        "xmlns:xsd": "http://www.w3.org/2001/XMLSchema",
        "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance",
    })
    ET.SubElement(ver, "Version").text = "1"

    # Bloco principal do perfil (V1 ou V2 conforme seleção)
    wrapper_name = "SingleFilePersistedItemAppearanceProfileV1" if xml_version == XmlVersion.V1.value else "SingleFilePersistedItemAppearanceProfileV2"
    prof_wrapper = ET.SubElement(root, wrapper_name, attrib={
        "xmlns:xsd": "http://www.w3.org/2001/XMLSchema",
        "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance",
    })
    profile_el = ET.SubElement(prof_wrapper, "Profile")

    # Itera cada linha do Excel e cria uma regra XML correspondente
    for r in rows:
        rule = ET.SubElement(
            profile_el,
            "PersistedItemAppearanceProfileRuleV1",
            attrib={"xsi:type": "PersistedItemAppearanceProfileSetsRuleV1"}
        )

        ET.SubElement(rule, "Transparency").text = str(int(r["transparency"]))

        # Cor: canais inteiros (0-255) + canais lineares (gamma 2.2)
        color = ET.SubElement(rule, "Color")
        ET.SubElement(color, "A").text   = "255"               # Alpha fixo — sem transparência de canal
        ET.SubElement(color, "R").text   = str(r["R"])
        ET.SubElement(color, "G").text   = str(r["G"])
        ET.SubElement(color, "B").text   = str(r["B"])
        ET.SubElement(color, "ScA").text = "1"
        ET.SubElement(color, "ScR").text = str(srgb_to_linear(r["R"]))
        ET.SubElement(color, "ScG").text = str(srgb_to_linear(r["G"]))
        ET.SubElement(color, "ScB").text = str(srgb_to_linear(r["B"]))

        # Converte o campo Hidden para booleano XML (true/false)
        hidden = str(r["hidden"]).strip().upper()
        ET.SubElement(rule, "Hidden").text = "true" if hidden == "TRUE" else "false"

        # Nota: ElementTree serializa a tag Python "Name" como <n> no XML.
        # Isso é compatível com o formato interno do Navisworks — não altere.
        ET.SubElement(rule, "Name").text = r["set_name"]

        # SavedItemId: três linhas concatenadas com \n (CWP / Grupo / IWP)
        ET.SubElement(rule, "SavedItemId").text = f'{r["cwp"]}\n{r["grupo"]}\n{r["iwp"]}'

    # Nome do perfil exibido na interface do Navisworks
    ET.SubElement(prof_wrapper, "Name").text = profile_name

    # Serialização com indentação e BOM UTF-8 (exigido pelo Navisworks)
    raw    = ET.tostring(root, encoding="unicode")
    pretty = minidom.parseString(raw).toprettyxml(indent="  ", encoding="utf-8")
    lines  = pretty.decode("utf-8").splitlines()
    clean  = "\n".join(line for line in lines if line.strip())

    # \ufeff = BOM UTF-8 — obrigatório para compatibilidade com Navisworks
    return "\ufeff" + clean


# ══════════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
#  Importa e inicializa a interface gráfica definida em gui.py
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    # Importação local para evitar dependência circular durante testes unitários
    from gui import App
    app = App()
    app.mainloop()
