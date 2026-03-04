"""
Analisador de Inventários de GEE - Identificação de Softwares/Plataformas
=========================================================================
v3 - Mudanças:
    - Remove falsos positivos do GHG Protocol / Registro Público
    - Extrai "Nome Fantasia" da página 2 de cada PDF
    - Extrai e-mail do responsável da seção "Dados do inventário"
    - Sanitização de caracteres ilegais para o Excel
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
TAMANHO_CONTEXTO = 250

_ILLEGAL_CHARS_RE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F\uFFFE\uFFFF]")

def sanitizar(texto: str) -> str:
    if not isinstance(texto, str):
        return str(texto) if texto is not None else ""
    limpo = _ILLEGAL_CHARS_RE.sub(" ", texto)
    return re.sub(r" {2,}", " ", limpo).strip()


# ── Softwares externos (sem GHG Protocol / Registro Público) ─────────────────
#
# REMOVIDOS intencionalmente para evitar falsos positivos:
#   - "Programa Brasileiro GHG Protocol" → é o programa de registro, não um software
#   - "Registro Público de Emissões"     → idem
#   - "pbghg", "plataforma ghg"          → idem
#   - Categoria "Ferramentas GHG Protocol / IPCC" inteira → são metodologias, não softwares
#   - "CDP / Carbon Disclosure Project"  → é um programa de reporte, não software de cálculo
#   - "SEEG", "inventário corporativo"   → referências ao registro, não ferramentas
#
SOFTWARES_CONHECIDOS = {

    "Plataforma Brasileira GEE": [
        r"way\s*carbon",
        r"waycarbon",
        r"deep\s*esg",
        r"\bclimas\b",
        r"\becosystem\b",
        r"\bgvces\b",
        r"fgv\s*ces",
        r"\bimaflora\b",
        r"\bidesam\b",
        r"\becam\b",
        r"\bqualidata\b",
    ],

    "Plataforma/Sistema GEE Internacional": [
        r"carbon\s*analytics",
        r"\bgreenbiz\b",
        r"\bsphera\b",
        r"\bwatershed\b",
        r"\bpersefoni\b",
        r"\bnormative\b",
        r"plan\s*a\b",
        r"net\s*zero\s*cloud",
        r"salesforce\s*net\s*zero",
        r"microsoft\s*cloud\s*for\s*sustainability",
        r"sap\s*sustainability",
        r"\benablon\b",
        r"\bintelex\b",
        r"\bcority\b",
        r"\becoact\b",
        r"\becodesk\b",
        r"\bcarbonfact\b",
        r"\bclimatiq\b",
        r"\bemitwise\b",
        r"\bcarbonsmart\b",
        r"\bcarbonchain\b",
        r"\bsweep\b",
        r"sinai\s*technologies",
        r"\bnovisto\b",
        r"sustainability\s*cloud",
        r"\bworkiva\b",
        r"diligent\s*esg",
        r"\bbriink\b",
        r"\bmeasurabl\b",
        r"\becometrica\b",
        r"co2\s*logic",
    ],

    "ACV / LCA": [
        r"\bsimapro\b",
        r"open\s*lca",
        r"\bgabi\b",
        r"\bumberto\b",
        r"\becoinvent\b",
        r"one\s*click\s*lca",
        r"\becochain\b",
    ],

    "ERP / Sistema Corporativo": [
        r"\bsap\b",
        r"\boracle\b",
        r"\btotvs\b",
        r"senior\s*sistemas",
        r"\bdatasul\b",
        r"\bprotheus\b",
        r"\blinx\b",
    ],

    "Energia / Utilidades": [
        r"energy\s*star\s*portfolio\s*manager",
        r"retscreen",
        r"ret\s*screen",
        r"\benergyplus\b",
        r"homer\s*energy",
    ],

    "BI / Dados": [
        r"power\s*bi",
        r"\btableau\b",
        r"\bqlik\b",
        r"\blooker\b",
        r"\bmetabase\b",
    ],

    # Planilha apenas quando citada como ferramenta de cálculo, não como formato de entrega
    "Planilha": [
        r"(?:elaborad[oa]|calculad[oa]|desenvolvid[oa]|construíd[oa])\s+(?:em|no|na|com)\s+(?:microsoft\s+)?excel",
        r"planilha\s+(?:microsoft\s+)?excel\s+(?:desenvolvida|elaborada|criada|própria|interna|personalizada)",
        r"(?:microsoft\s+)?excel\s+(?:como\s+)?(?:ferramenta|software|sistema)\s+(?:de\s+)?(?:cálculo|inventário|controle|gestão)",
        r"google\s+sheets",
        r"planilha\s+eletr[oô]nica\s+(?:própria|interna|personalizada|desenvolvida)",
    ],
}

# Padrões genéricos — capturam softwares não listados acima
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
    "software", "plataforma", "ferramenta", "ghg", "gee", "co2",
    "protocolo", "protocol", "programa", "brasileiro",
}


# ── Extração de texto por página ──────────────────────────────────────────────

def extrair_paginas(caminho_pdf: Path, max_paginas=None):
    """Retorna lista de strings, uma por página."""
    try:
        doc = fitz.open(str(caminho_pdf))
        paginas = list(doc.pages()) if max_paginas is None else list(doc.pages())[:max_paginas]
        textos = [p.get_text() for p in paginas]
        doc.close()
        return textos
    except Exception as e:
        return [f"__ERRO__: {e}"]


def contar_paginas(caminho_pdf: Path, max_paginas=None) -> int:
    try:
        doc = fitz.open(str(caminho_pdf))
        n = len(doc)
        doc.close()
        return n if max_paginas is None else min(n, max_paginas)
    except Exception:
        return 0


# ── Extração de Nome Fantasia ─────────────────────────────────────────────────

def extrair_nome_fantasia(paginas: list) -> str:
    """
    Busca 'Nome Fantasia' nas primeiras páginas do PDF.
    O formulário GHG Protocol apresenta o campo assim:
        Nome Fantasia
        <valor na linha seguinte>
    ou
        Nome Fantasia: <valor>
    """
    # Busca nas primeiras 3 páginas (geralmente está na pág 1 ou 2)
    texto_busca = "\n".join(paginas[:3]) if len(paginas) >= 3 else "\n".join(paginas)

    padroes = [
        # "Nome Fantasia\nACME Ltda"
        r"[Nn]ome\s+[Ff]antasia\s*[:\n]\s*([^\n]{2,100})",
        # "Nome Fantasia   ACME Ltda" (espaços em vez de newline, em PDFs compactados)
        r"[Nn]ome\s+[Ff]antasia\s{2,}([^\n]{2,100})",
    ]

    for padrao in padroes:
        m = re.search(padrao, texto_busca)
        if m:
            nome = m.group(1).strip()
            # Remove lixo típico: números de página, datas, rodapés
            nome = re.sub(r"\s+\d{1,2}/\d{1,2}/\d{4}.*$", "", nome).strip()
            nome = re.sub(r"\s+\d+\s*$", "", nome).strip()
            if len(nome) > 1:
                return sanitizar(nome)

    return ""


# ── Extração de E-mail do Responsável ────────────────────────────────────────

_EMAIL_RE = re.compile(r"[\w\.\-\+]+@[\w\.\-]+\.[a-zA-Z]{2,}")

def extrair_email_responsavel(paginas: list) -> str:
    """
    Busca o e-mail na seção 'Dados do inventário' / 'E-mail do responsável'.
    Estratégia:
      1. Procura a label 'e-mail do responsável' e captura o e-mail próximo
      2. Se não achar, pega o primeiro e-mail das primeiras páginas
    """
    texto_busca = "\n".join(paginas[:4]) if len(paginas) >= 4 else "\n".join(paginas)

    # Tenta encontrar o e-mail logo após a label do campo
    padroes_label = [
        r"[Ee]-?mail\s+do\s+[Rr]esponsável\s*[:\n]\s*([\w\.\-\+]+@[\w\.\-]+\.[a-zA-Z]{2,})",
        r"[Ee]-?mail\s+do\s+[Rr]esponsável\s*[:\n][^\n@]{0,60}?([\w\.\-\+]+@[\w\.\-]+\.[a-zA-Z]{2,})",
        r"[Ee]-?mail\s*[:\n]\s*([\w\.\-\+]+@[\w\.\-]+\.[a-zA-Z]{2,})",
    ]

    for padrao in padroes_label:
        m = re.search(padrao, texto_busca)
        if m:
            return sanitizar(m.group(1).strip())

    # Fallback: primeiro e-mail encontrado nas primeiras páginas
    emails = _EMAIL_RE.findall(texto_busca)
    if emails:
        return sanitizar(emails[0])

    return ""


# ── Detecção de softwares ─────────────────────────────────────────────────────

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


# ── Análise da pasta ──────────────────────────────────────────────────────────

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
        paginas = extrair_paginas(pdf_path, max_paginas)

        # Verifica erro de leitura
        if paginas and paginas[0].startswith("__ERRO__"):
            resultados.append({
                "arquivo":            sanitizar(pdf_path.name),
                "empresa":            "",
                "email_responsavel":  "",
                "usa_software":       "ERRO",
                "softwares":          "",
                "categorias":         "",
                "contexto_1":         sanitizar(paginas[0]),
                "contexto_2":         "",
                "contexto_3":         "",
                "total_softwares":    0,
                "paginas_analisadas": 0,
                "tamanho_texto":      0,
            })
            continue

        texto_completo = "\n".join(paginas)

        # Extrai metadados
        nome_empresa     = extrair_nome_fantasia(paginas)
        email_responsavel = extrair_email_responsavel(paginas)

        # Se não encontrou nome fantasia, usa o nome do arquivo como fallback
        if not nome_empresa:
            nome_empresa = sanitizar(pdf_path.stem)

        # Detecta softwares
        achados = encontrar_softwares_conhecidos(texto_completo) + encontrar_softwares_genericos(texto_completo)
        c = consolidar_achados(achados)

        resultados.append({
            "arquivo":            sanitizar(pdf_path.name),
            "empresa":            nome_empresa,
            "email_responsavel":  email_responsavel,
            "usa_software":       "SIM" if c["softwares"] else "NÃO",
            "softwares":          sanitizar(" | ".join(c["softwares"])),
            "categorias":         sanitizar(" | ".join(c["categorias"])),
            "contexto_1":         sanitizar(c["contextos"][0]) if len(c["contextos"]) > 0 else "",
            "contexto_2":         sanitizar(c["contextos"][1]) if len(c["contextos"]) > 1 else "",
            "contexto_3":         sanitizar(c["contextos"][2]) if len(c["contextos"]) > 2 else "",
            "total_softwares":    len(c["softwares"]),
            "paginas_analisadas": contar_paginas(pdf_path, max_paginas),
            "tamanho_texto":      len(texto_completo),
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
        "Arquivo PDF",
        "Nome Fantasia (Empresa)",
        "E-mail Responsável",
        "Usa Software Externo?",
        "Softwares / Plataformas",
        "Categorias",
        "Total",
        "Páginas",
        "Contexto 1",
        "Contexto 2",
        "Contexto 3",
    ]
    chaves = [
        "arquivo",
        "empresa",
        "email_responsavel",
        "usa_software",
        "softwares",
        "categorias",
        "total_softwares",
        "paginas_analisadas",
        "contexto_1",
        "contexto_2",
        "contexto_3",
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

    larguras = [25, 45, 35, 20, 50, 35, 8, 8, 80, 80, 80]
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
    com_email = sum(1 for r in resultados if r.get("email_responsavel", ""))

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
        ("Empresas com software externo identificado:", com_sw),
        ("% que usam software externo:", "=B4/B3" if total > 0 else "N/A"),
        ("Empresas sem software externo:", sem_sw),
        ("E-mails de responsáveis encontrados:", com_email),
        ("PDFs com erro de leitura:", com_erro),
        ("Data da análise:", datetime.now().strftime("%d/%m/%Y %H:%M")),
    ]

    for i, (label, valor) in enumerate(dados_resumo, 3):
        ws2.cell(row=i, column=1, value=label).font = Font(name="Arial", size=10)
        c = ws2.cell(row=i, column=2, value=valor)
        c.font = Font(name="Arial", size=10)
        if i == 5 and total > 0:
            c.number_format = "0.0%"

    ws2["A11"] = "SOFTWARES / PLATAFORMAS MAIS FREQUENTES"
    ws2["A11"].font = Font(name="Arial", bold=True, size=12, color="1F4E79")

    for col, title in [("A", "Software / Plataforma"), ("B", "Nº de Empresas")]:
        c = ws2[f"{col}12"]
        c.value = title
        c.font = Font(name="Arial", bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", start_color=COR_CAB)
        c.alignment = al_cen

    for i, (sw, cnt) in enumerate(sw_ord, 13):
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
    com_email = sum(1 for r in resultados if r.get("email_responsavel", ""))

    print(f"\n📊 Resumo:")
    print(f"   Total de PDFs:                {total}")
    print(f"   Com software externo:         {com_sw} ({100*com_sw/total:.1f}%)")
    print(f"   Sem software externo:         {total - com_sw - erros}")
    print(f"   E-mails encontrados:          {com_email} ({100*com_email/total:.1f}%)")
    print(f"   Erros de leitura:             {erros}")

    gerar_excel(resultados, args.saida)


if __name__ == "__main__":
    main()
