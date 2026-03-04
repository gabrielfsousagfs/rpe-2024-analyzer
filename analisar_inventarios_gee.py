"""
Analisador de Inventários de GEE - Identificação de Softwares/Plataformas
=========================================================================
Lê todos os PDFs de uma pasta, identifica menções a softwares/plataformas
usados na elaboração de inventários de GEE e gera uma planilha Excel.

USO LOCAL:
    python analisar_inventarios_gee.py --pasta "./pdfs"

USO NO GITHUB ACTIONS:
    Disparado automaticamente pelo workflow. Os PDFs são baixados
    do Google Drive antes da execução.

CORREÇÕES v2:
    - Sanitização de caracteres de controle (IllegalCharacterError do openpyxl)
    - Padrão "Excel" restringido para evitar falsos positivos massivos
"""

import os
import re
import sys
import argparse
from pathlib import Path
from datetime import datetime

try:
    import fitz  # PyMuPDF
except ImportError:
    print("❌ PyMuPDF não encontrado. Instale com: pip install pymupdf")
    sys.exit(1)

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("❌ openpyxl não encontrado. Instale com: pip install openpyxl")
    sys.exit(1)

try:
    from tqdm import tqdm
    TQDM_DISPONIVEL = True
except ImportError:
    TQDM_DISPONIVEL = False

# ── Configuração ──────────────────────────────────────────────────────────────
PASTA_PDFS       = os.environ.get("PASTA_PDFS", "./pdfs")
MAX_PAGINAS      = None
TAMANHO_CONTEXTO = 200

# Caracteres ilegais para o openpyxl (caracteres de controle ASCII, exceto tab/LF/CR)
_ILLEGAL_CHARS_RE = re.compile(
    r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F\uFFFE\uFFFF]"
)

def sanitizar(texto: str) -> str:
    """Remove caracteres ilegais para o Excel e normaliza espaços."""
    if not isinstance(texto, str):
        return str(texto) if texto is not None else ""
    limpo = _ILLEGAL_CHARS_RE.sub(" ", texto)
    limpo = re.sub(r" {2,}", " ", limpo)
    return limpo.strip()

# ── Softwares e plataformas conhecidos ───────────────────────────────────────
SOFTWARES_CONHECIDOS = {

    "Plataforma Brasileira GEE": [
        r"way\s*carbon",
        r"waycarbon",
        r"deep\s*esg",
        r"climas\b",
        r"ecosystem\b",
        r"pbghg",
        r"programa\s*brasileiro\s*ghg",
        r"registro\s*p[uú]blico\s*de\s*emiss[oõ]es",
        r"gvces",
        r"fgv\s*ces",
        r"seeg\b",
        r"imaflora",
        r"idesam",
        r"ecam\b",
        r"qualidata",
        r"inventário\s*corporativo",
        r"plataforma\s*ghg",
    ],

    "Plataforma/Sistema GEE Internacional": [
        r"carbon\s*analytics",
        r"greenbiz",
        r"sphera",
        r"watershed",
        r"persefoni",
        r"normative",
        r"plan\s*a\b",
        r"net\s*zero\s*cloud",
        r"salesforce\s*net\s*zero",
        r"microsoft\s*cloud\s*for\s*sustainability",
        r"sap\s*sustainability",
        r"enablon",
        r"intelex",
        r"cority",
        r"ecoact",
        r"ecodesk",
        r"carbonfact",
        r"climatiq",
        r"emitwise",
        r"carbonsmart",
        r"carbonchain",
        r"sweep\b",
        r"sinai\s*technologies",
        r"novisto",
        r"sustainability\s*cloud",
        r"workiva",
        r"diligent\s*esg",
        r"one\s*report",
        r"briink",
        r"measurabl",
        r"ecometrica",
        r"carbon\s*foot\s*print\s*(?:software|tool|platform)",
        r"co2\s*logic",
    ],

    "Ferramentas GHG Protocol / IPCC": [
        r"ghg\s*protocol\s*tool",
        r"cross.sector\s*tool",
        r"stationary\s*combustion\s*tool",
        r"mobile\s*combustion\s*tool",
        r"ipcc\s*tool",
        r"global\s*warming\s*potential\s*tool",
    ],

    "CDP": [
        r"carbon\s*disclosure\s*project",
        r"\bcdp\b",
    ],

    "ACV / LCA": [
        r"simapro",
        r"open\s*lca",
        r"gabi\b",
        r"umberto\b",
        r"ecoinvent",
        r"one\s*click\s*lca",
        r"ecochain",
    ],

    "ERP / Sistema Corporativo": [
        r"\bsap\b",
        r"\boracle\b",
        r"\btotvs\b",
        r"senior\s*sistemas",
        r"datasul",
        r"protheus",
        r"\blinx\b",
    ],

    "Energia / Utilidades": [
        r"energy\s*star\s*portfolio\s*manager",
        r"retscreen",
        r"ret\s*screen",
        r"energyplus",
        r"homer\s*energy",
    ],

    "BI / Dados": [
        r"power\s*bi",
        r"\btableau\b",
        r"\bqlik\b",
        r"\blooker\b",
        r"\bmetabase\b",
    ],

    # Padrões restritos para evitar falsos positivos:
    # "Excel" aparece em quase todo inventário GHG Protocol como formato de entrega.
    # Só marcamos quando mencionado explicitamente como ferramenta de cálculo.
    "Planilha": [
        r"(?:elaborad[oa]|calculad[oa]|desenvolvid[oa]|construíd[oa])\s+(?:em|no|na|com)\s+(?:microsoft\s+)?excel",
        r"planilha\s+(?:microsoft\s+)?excel\s+(?:desenvolvida|elaborada|criada|própria|interna|personalizada)",
        r"(?:microsoft\s+)?excel\s+(?:como\s+)?(?:ferramenta|software|sistema)\s+(?:de\s+)?(?:cálculo|inventário|controle|gestão)",
        r"google\s+sheets",
        r"planilha\s+eletr[oô]nica\s+(?:própria|interna|personalizada|desenvolvida)",
    ],
}

PADROES_GENERICOS = [
    r"(?:software|plataforma|sistema|ferramenta|aplicativo)\s+(?:de\s+\w+\s+){0,3}[\"']?([\w\s\-\.]{3,40}?)[\"']?\s*(?:para|foi|,|\.|;)",
    r"(?:utilizad[oa]|usad[oa]|empregad[oa]|adotad[oa])\s+(?:o|a)\s+(?:software|plataforma|sistema|ferramenta)\s+[\"']?([\w\s\-\.]{3,40}?)[\"']",
    r"calculad[oa]\s+(?:no|na|pelo|pela|através\s+do|através\s+da)\s+(?:software|plataforma|sistema)\s+[\"']?([\w\s\-\.]{3,40}?)[\"']?\s*(?:,|\.|;|\n)",
    r"[\"']([\w\s\-\.]{3,40})[\"']\s*(?:software|plataforma|sistema|ferramenta)",
]

PALAVRAS_CONTEXTO_GEE = [
    "emissão", "emissões", "inventário", "carbono", "ghg", "gee",
    "gases", "efeito estufa", "co2", "ch4", "n2o", "escopo", "scope",
    "fator de emissão", "cálculo", "quantificação", "monitoramento",
    "gestão ambiental", "pegada de carbono",
]

PALAVRAS_IGNORAR = {
    "o", "a", "os", "as", "um", "uma", "que", "de", "do", "da",
    "para", "com", "em", "por", "este", "esta", "esse", "essa",
    "dados", "informações", "relatório", "cálculos", "sistema",
    "software", "plataforma", "ferramenta",
}


# ── Funções de extração e análise ────────────────────────────────────────────

def extrair_texto_pdf(caminho_pdf: Path, max_paginas=None) -> str:
    try:
        doc = fitz.open(str(caminho_pdf))
        paginas = list(doc.pages()) if max_paginas is None else list(doc.pages())[:max_paginas]
        texto = "\n".join(p.get_text() for p in paginas)
        doc.close()
        return texto
    except Exception as e:
        return f"__ERRO__: {e}"


def contar_paginas(caminho_pdf: Path, max_paginas=None) -> int:
    try:
        doc = fitz.open(str(caminho_pdf))
        n = len(doc)
        doc.close()
        return n if max_paginas is None else min(n, max_paginas)
    except Exception:
        return 0


def encontrar_softwares_conhecidos(texto: str) -> list:
    texto_lower = texto.lower()
    achados = []
    for categoria, padroes in SOFTWARES_CONHECIDOS.items():
        for padrao in padroes:
            for match in re.finditer(padrao, texto_lower, re.IGNORECASE):
                inicio = max(0, match.start() - TAMANHO_CONTEXTO)
                fim = min(len(texto), match.end() + TAMANHO_CONTEXTO)
                contexto = texto[inicio:fim].replace("\n", " ").strip()
                nome = texto[match.start():match.end()].strip()
                achados.append({"software": nome, "categoria": categoria, "contexto": contexto})
    return achados


def encontrar_softwares_genericos(texto: str) -> list:
    texto_lower = texto.lower()
    if not any(p in texto_lower for p in PALAVRAS_CONTEXTO_GEE):
        return []
    achados = []
    for padrao in PADROES_GENERICOS:
        for match in re.finditer(padrao, texto, re.IGNORECASE):
            try:
                nome = match.group(1).strip()
            except IndexError:
                continue
            if len(nome) < 3 or len(nome) > 60 or nome.lower() in PALAVRAS_IGNORAR:
                continue
            inicio = max(0, match.start() - TAMANHO_CONTEXTO)
            fim = min(len(texto), match.end() + TAMANHO_CONTEXTO)
            contexto = texto[inicio:fim].replace("\n", " ").strip()
            achados.append({"software": nome, "categoria": "Identificado automaticamente", "contexto": contexto})
    return achados


def consolidar_achados(achados: list) -> dict:
    vistos = set()
    r = {"softwares": [], "categorias": [], "contextos": []}
    for a in achados:
        chave = a["software"].lower().strip()
        if chave not in vistos:
            vistos.add(chave)
            r["softwares"].append(a["software"])
            r["categorias"].append(a["categoria"])
            r["contextos"].append(a["contexto"][:500])
    return r


def analisar_pasta(pasta: str, max_paginas=None) -> list:
    pdfs = sorted(Path(pasta).glob("**/*.pdf"))
    if not pdfs:
        print(f"⚠️  Nenhum PDF encontrado em: {pasta}")
        return []

    print(f"\n📂 Pasta: {pasta}")
    print(f"📄 PDFs encontrados: {len(pdfs)}\n")

    resultados = []
    iterador = tqdm(pdfs, desc="Analisando PDFs", unit="pdf") if TQDM_DISPONIVEL else pdfs

    for pdf_path in iterador:
        texto = extrair_texto_pdf(pdf_path, max_paginas)

        if texto.startswith("__ERRO__"):
            resultados.append({
                "arquivo": sanitizar(pdf_path.name),
                "empresa": sanitizar(pdf_path.stem),
                "usa_software": "ERRO",
                "softwares": "",
                "categorias": "",
                "contexto_1": sanitizar(texto),
                "contexto_2": "",
                "contexto_3": "",
                "total_softwares": 0,
                "paginas_analisadas": 0,
                "tamanho_texto": 0,
            })
            continue

        achados = encontrar_softwares_conhecidos(texto) + encontrar_softwares_genericos(texto)
        c = consolidar_achados(achados)

        resultados.append({
            "arquivo": sanitizar(pdf_path.name),
            "empresa": sanitizar(pdf_path.stem),
            "usa_software": "SIM" if c["softwares"] else "NÃO",
            "softwares": sanitizar(" | ".join(c["softwares"])),
            "categorias": sanitizar(" | ".join(c["categorias"])),
            "contexto_1": sanitizar(c["contextos"][0]) if len(c["contextos"]) > 0 else "",
            "contexto_2": sanitizar(c["contextos"][1]) if len(c["contextos"]) > 1 else "",
            "contexto_3": sanitizar(c["contextos"][2]) if len(c["contextos"]) > 2 else "",
            "total_softwares": len(c["softwares"]),
            "paginas_analisadas": contar_paginas(pdf_path, max_paginas),
            "tamanho_texto": len(texto),
        })

    return resultados


# ── Geração da planilha Excel ─────────────────────────────────────────────────

def gerar_excel(resultados: list, caminho_saida: str):
    wb = openpyxl.Workbook()

    COR_CAB  = "1F4E79"
    COR_SIM  = "C6EFCE"
    COR_NAO  = "F4CCCC"
    COR_ERRO = "FCE4D6"
    COR_ALT  = "EBF3FB"

    f_cab  = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    f_sim  = Font(name="Arial", bold=True, color="375623")
    f_nao  = Font(name="Arial", color="7F0000")
    f_norm = Font(name="Arial", size=10)
    al_cen = Alignment(horizontal="center", vertical="center", wrap_text=True)
    al_esq = Alignment(horizontal="left", vertical="top", wrap_text=True)
    borda  = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"),  bottom=Side(style="thin"),
    )

    # ── Aba Resultados ───────────────────────────────────────────────────────
    ws = wb.active
    ws.title = "Resultados"

    cabecalhos = [
        "Arquivo PDF", "Empresa", "Usa Software?", "Softwares / Plataformas",
        "Categorias", "Total", "Páginas", "Contexto 1", "Contexto 2", "Contexto 3",
    ]
    chaves = [
        "arquivo", "empresa", "usa_software", "softwares", "categorias",
        "total_softwares", "paginas_analisadas", "contexto_1", "contexto_2", "contexto_3",
    ]

    for ci, cab in enumerate(cabecalhos, 1):
        c = ws.cell(row=1, column=ci, value=cab)
        c.font = f_cab
        c.fill = PatternFill("solid", start_color=COR_CAB)
        c.alignment = al_cen
        c.border = borda

    for ri, res in enumerate(resultados, 2):
        cor_linha = COR_ALT if ri % 2 == 0 else "FFFFFF"
        for ci, chave in enumerate(chaves, 1):
            valor = res.get(chave, "")
            if isinstance(valor, str):
                valor = sanitizar(valor)
            c = ws.cell(row=ri, column=ci, value=valor)
            c.font = f_norm
            c.border = borda
            c.fill = PatternFill("solid", start_color=cor_linha)
            if chave == "usa_software":
                if valor == "SIM":
                    c.fill = PatternFill("solid", start_color=COR_SIM)
                    c.font = f_sim
                elif valor == "NÃO":
                    c.fill = PatternFill("solid", start_color=COR_NAO)
                    c.font = f_nao
                elif valor == "ERRO":
                    c.fill = PatternFill("solid", start_color=COR_ERRO)
                c.alignment = al_cen
            elif chave in ("total_softwares", "paginas_analisadas"):
                c.alignment = al_cen
            else:
                c.alignment = al_esq

    larguras = [35, 40, 14, 50, 35, 10, 10, 80, 80, 80]
    for i, larg in enumerate(larguras, 1):
        ws.column_dimensions[get_column_letter(i)].width = larg

    for row in ws.iter_rows(min_row=2, max_row=len(resultados) + 1):
        ws.row_dimensions[row[0].row].height = 55

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    # ── Aba Resumo ───────────────────────────────────────────────────────────
    ws2 = wb.create_sheet("Resumo")
    total    = len(resultados)
    com_sw   = sum(1 for r in resultados if r["usa_software"] == "SIM")
    sem_sw   = sum(1 for r in resultados if r["usa_software"] == "NÃO")
    com_erro = sum(1 for r in resultados if r["usa_software"] == "ERRO")

    contagem_sw = {}
    for r in resultados:
        for sw in r["softwares"].split(" | "):
            sw = sw.strip()
            if sw:
                contagem_sw[sw] = contagem_sw.get(sw, 0) + 1
    sw_ord = sorted(contagem_sw.items(), key=lambda x: x[1], reverse=True)

    ws2["A1"] = "RESUMO DA ANÁLISE — INVENTÁRIOS GEE 2024"
    ws2["A1"].font = Font(name="Arial", bold=True, size=14, color="1F4E79")

    dados_resumo = [
        ("Total de PDFs analisados:", total),
        ("Empresas com software/plataforma identificado:", com_sw),
        ("% que usam software:", "=B4/B3" if total > 0 else "N/A"),
        ("Empresas sem software identificado:", sem_sw),
        ("PDFs com erro de leitura:", com_erro),
        ("Data da análise:", datetime.now().strftime("%d/%m/%Y %H:%M")),
    ]

    for i, (label, valor) in enumerate(dados_resumo, 3):
        ws2.cell(row=i, column=1, value=label).font = Font(name="Arial", size=10)
        c = ws2.cell(row=i, column=2, value=valor)
        c.font = Font(name="Arial", size=10)
        if i == 5 and total > 0:
            c.number_format = "0.0%"

    ws2["A10"] = "SOFTWARES / PLATAFORMAS MAIS FREQUENTES"
    ws2["A10"].font = Font(name="Arial", bold=True, size=12, color="1F4E79")

    for col, title in [("A", "Software / Plataforma"), ("B", "Nº de Empresas")]:
        c = ws2[f"{col}11"]
        c.value = title
        c.font = Font(name="Arial", bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", start_color=COR_CAB)
        c.alignment = al_cen

    for i, (sw, cnt) in enumerate(sw_ord, 12):
        cor = COR_ALT if i % 2 == 0 else "FFFFFF"
        ws2[f"A{i}"] = sanitizar(sw)
        ws2[f"B{i}"] = cnt
        for col in ("A", "B"):
            ws2[f"{col}{i}"].fill = PatternFill("solid", start_color=cor)
            ws2[f"{col}{i}"].font = Font(name="Arial", size=10)
        ws2[f"B{i}"].alignment = Alignment(horizontal="center")

    ws2.column_dimensions["A"].width = 50
    ws2.column_dimensions["B"].width = 20

    Path(caminho_saida).parent.mkdir(parents=True, exist_ok=True)
    wb.save(caminho_saida)
    print(f"\n✅ Planilha salva em: {caminho_saida}")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Analisa PDFs de inventários GEE.")
    parser.add_argument("--pasta",   "-p", default=PASTA_PDFS)
    parser.add_argument("--saida",   "-s", default=None)
    parser.add_argument("--paginas", "-n", type=int, default=MAX_PAGINAS)
    args = parser.parse_args()

    if args.saida is None:
        data_hoje = datetime.now().strftime("%Y-%m-%d")
        args.saida = str(Path(args.pasta) / f"resultados_gee_{data_hoje}.xlsx")

    print("=" * 60)
    print("  ANALISADOR DE INVENTÁRIOS GEE — SOFTWARES E PLATAFORMAS")
    print("=" * 60)

    resultados = analisar_pasta(args.pasta, args.paginas)
    if not resultados:
        print("Nenhum resultado para salvar.")
        return

    total  = len(resultados)
    com_sw = sum(1 for r in resultados if r["usa_software"] == "SIM")
    erros  = sum(1 for r in resultados if r["usa_software"] == "ERRO")

    print(f"\n📊 Resumo:")
    print(f"   Total de PDFs:              {total}")
    print(f"   Com software/plataforma:    {com_sw} ({100*com_sw/total:.1f}%)")
    print(f"   Sem software identificado:  {total - com_sw - erros}")
    print(f"   Erros de leitura:           {erros}")

    gerar_excel(resultados, args.saida)


if __name__ == "__main__":
    main()
