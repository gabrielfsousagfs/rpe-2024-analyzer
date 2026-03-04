# 📊 Analisador de Inventários GEE — Softwares e Plataformas

Identifica automaticamente quais softwares/plataformas foram utilizados na elaboração dos inventários de GEE publicados no Registro Público de Emissões (2024), gerando uma planilha Excel com os resultados.

---

## 🗂️ Estrutura do repositório

```
├── analisar_inventarios_gee.py   # Script principal
├── requirements.txt              # Dependências Python
├── .github/
│   └── workflows/
│       └── analisar_gee.yml      # Workflow do GitHub Actions
└── README.md
```

---

## ☁️ Onde armazenar os PDFs

> **❌ Não coloque os PDFs no repositório GitHub.**
> O GitHub tem limite de 100 MB por arquivo e recomenda repositórios abaixo de 1 GB. Com 1.300 PDFs e mais de 2 GB, o repositório ficaria inviável.

### ✅ Solução recomendada: Google Drive

1. Comprima todos os PDFs em **um único arquivo ZIP** (`inventarios_gee_2024.zip`)
2. Faça upload desse ZIP para o seu Google Drive
3. Clique com botão direito → **Compartilhar** → "Qualquer pessoa com o link pode ver"
4. Copie o **ID do arquivo** da URL:
   ```
   https://drive.google.com/file/d/  ➡️ 1aBcDeFgHiJkLmNoPqRsTuVwXyZ  ⬅️  /view
                                          └── este é o ID
   ```
5. Use esse ID ao disparar o workflow

---

## 🚀 Como executar no GitHub Actions

### Passo a passo

1. **Fork ou clone** este repositório para a sua conta GitHub
2. Faça o upload do ZIP no Google Drive (veja acima)
3. No repositório GitHub, clique em **Actions** (menu superior)
4. Selecione o workflow **"Analisar Inventários GEE"**
5. Clique em **"Run workflow"** (botão à direita)
6. Preencha os campos:
   | Campo | Descrição |
   |---|---|
   | **ID do arquivo ZIP no Google Drive** | O ID copiado da URL do Drive |
   | **Máximo de páginas por PDF** | `0` para analisar tudo; use `30` para uma análise mais rápida |
7. Clique em **"Run workflow"** e aguarde (pode levar de 30 min a 2h dependendo do tamanho)

### Baixar o resultado

Após a execução:
1. Clique no workflow concluído
2. Role até a seção **"Artifacts"** no final da página
3. Clique em **`resultados-gee-N`** para baixar a planilha `.xlsx`

---

## 📋 O que a planilha contém

### Aba "Resultados"
| Coluna | Descrição |
|---|---|
| Arquivo PDF | Nome do arquivo |
| Empresa | Nome da empresa (baseado no nome do arquivo) |
| Usa Software? | **SIM** / **NÃO** / **ERRO** |
| Softwares / Plataformas | Lista separada por `\|` |
| Categorias | Categoria de cada software encontrado |
| Total | Número de softwares distintos encontrados |
| Páginas Analisadas | Quantas páginas foram lidas |
| Contexto 1–3 | Trechos do texto onde o software foi mencionado |

### Aba "Resumo"
- Total de PDFs analisados
- % de empresas que usam software/plataforma
- Ranking dos softwares mais frequentes

---

## 🔍 Softwares detectados automaticamente

| Categoria | Exemplos |
|---|---|
| Plataforma Brasileira GEE | **WayCarbon**, **DEEP ESG**, **Climas**, **Ecosystem**, PBGHG, GVCes, SEEG |
| Plataforma Internacional GEE | Sphera, Watershed, Persefoni, Normative, Workiva, Measurabl |
| CDP | Carbon Disclosure Project |
| Ferramentas GHG Protocol | Cross-Sector Tool, Stationary Combustion Tool |
| ACV / LCA | SimaPro, OpenLCA, GaBi, ecoinvent |
| ERP | SAP, Oracle, TOTVS |
| Planilha | Excel, Google Sheets |
| BI / Dados | Power BI, Tableau |
| Identificação automática | Qualquer software mencionado próximo a termos de GEE |

---

## 💻 Execução local (opcional)

```bash
# Instalar dependências
pip install -r requirements.txt

# Executar
python analisar_inventarios_gee.py --pasta "./pdfs"

# Com limite de páginas (mais rápido)
python analisar_inventarios_gee.py --pasta "./pdfs" --paginas 30
```

---

## ⚙️ Customização

Para adicionar novos softwares, edite o dicionário `SOFTWARES_CONHECIDOS` em `analisar_inventarios_gee.py`:

```python
"Minha Categoria": [
    r"nome\s*do\s*software",   # aceita regex
    r"outro\s*software",
],
```
