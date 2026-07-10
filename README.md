# Dashboard SCM Semanal (WK)

Gerador de dashboard semanal de SCM (Suprimentos/Compras): lê a planilha de
Solicitações de Compra (SCs) e Pedidos de Compra (POs) da empresa e produz um
**dashboard HTML interativo e 100% autossuficiente** (dados + gráficos embutidos),
pronto para ser aberto no navegador ou enviado por e-mail/Teams sem depender de
internet, Python ou pastas externas.

## Sumário

- [Objetivo](#objetivo)
- [Como funciona](#como-funciona)
- [Uso rápido (sem instalar nada)](#uso-rápido-sem-instalar-nada)
- [Tecnologias](#tecnologias)
- [Instalação (fluxo dev / Python)](#instalação-fluxo-dev--python)
- [Execução](#execução)
- [Estrutura do projeto](#estrutura-do-projeto)
- [Configuração](#configuração)
- [Dados confidenciais](#dados-confidenciais)
- [Fluxo legado (Streamlit)](#fluxo-legado-streamlit)
- [Troubleshooting](#troubleshooting)
- [Roadmap](#roadmap)
- [Créditos](#créditos)

## Objetivo

Automatizar o relatório semanal de SCM, eliminando a digitação manual de números
num HTML/PowerPoint. Todo KPI, gráfico e ranking é calculado direto da planilha e
**tudo é clicável**: clicar em qualquer número/barra/fatia abre uma tabela com os
registros por trás daquele valor (buscar, ordenar, exportar CSV). Uma aba
"Comparar Semanas" acumula histórico automaticamente a cada geração.

## Como funciona

```
Relatório de SCs *.xlsx  --(gerar_dashboard.py)-->  WK/Dashboard SCM WK<nn>.html
        (abas SCM e SC7)                             (autossuficiente, abre sozinho)
```

1. `gerar_dashboard.py` lê as abas **SCM** (itens em cotação) e **SC7** (POs) da
   planilha mais recente.
2. Calcula KPIs, aging, ranking por departamento/comprador, status de POs etc.
3. Injeta os dados (JSON) e as libs de gráfico (Chart.js) **inline** dentro de
   `template_dashboard.html`, gerando um único arquivo HTML sem dependências
   externas.
4. Salva um resumo da semana em `WK/data/WK<nn>_<ano>.json` para alimentar a aba
   de comparação nas próximas gerações.

## Uso rápido (sem instalar nada)

Não precisa de Python, VSCode nem `pip install`. Baixe o executável na aba
[Releases](../../releases) do repositório:

1. Baixe `Gerar Dashboard SCM.exe` e coloque na mesma pasta da planilha da semana.
2. Dê duplo-clique (ou arraste o `.xlsx` para cima do `.exe`).
3. O dashboard é gerado e abre sozinho no navegador.

O `.exe` é gerado com [PyInstaller](https://pyinstaller.org/) a partir de
`gerar_dashboard.py`, com `template_dashboard.html` e `WK/assets/*.js` embutidos.
Para gerar um novo build depois de alterar o script:

```bash
pip install pyinstaller
pyinstaller --onefile --name "Gerar Dashboard SCM" ^
  --add-data "template_dashboard.html;." ^
  --add-data "WK/assets;WK/assets" ^
  gerar_dashboard.py
```

O executável fica em `dist/` (pasta ignorada pelo git — o binário (~70MB) é
distribuído via Release, não commitado no repositório).

## Tecnologias

- **Python 3.10+**, [pandas](https://pandas.pydata.org/) e
  [openpyxl](https://openpyxl.readthedocs.io/) para leitura/processamento do Excel.
- **HTML/CSS/JS puro** no template do dashboard.
- [Chart.js](https://www.chartjs.org/) 4.4.1 e
  [chartjs-plugin-datalabels](https://github.com/chartjs/chartjs-plugin-datalabels)
  2.2.0 vendorizados em `WK/assets/` (embutidos no HTML final, licença MIT).
- `tkinter` (biblioteca padrão do Python) para o seletor de arquivo.

## Instalação (fluxo dev / Python)

Só necessário se você for mexer no código-fonte (`gerar_dashboard.py`) — para uso
do dia a dia, prefira o [.exe pronto](#uso-rápido-sem-instalar-nada).

Pré-requisito: [Python 3.10+](https://www.python.org/downloads/) instalado e no
PATH (no Windows, marque "Add Python to PATH" durante a instalação).

```bash
git clone https://github.com/eholuwi/dashboard-SCM-wk.git
cd dashboard-SCM-wk
pip install -r requirements.txt
```

## Execução

1. Coloque a planilha da semana (`Relatório de SCs*.xlsx`) na raiz do projeto.
2. Rode um dos comandos abaixo:

```bash
# Abre um seletor de arquivo (ou usa a planilha "Relatório de SCs*.xlsx" mais
# recente da pasta se você cancelar a janela)
python gerar_dashboard.py

# Ou informe o arquivo direto
python gerar_dashboard.py "Relatório de SCs 01.07.xlsx"
```

No Windows, também é possível apenas dar **duplo-clique em `Gerar Dashboard.bat`**
(ou arrastar o `.xlsx` para cima dele) — veja `COMO USAR.txt` para o passo a passo
não-técnico usado no dia a dia.

O dashboard é salvo em `WK/Dashboard SCM WK<nn>.html` e abre sozinho no navegador
padrão. Para enviar a outra pessoa, basta compartilhar esse único arquivo `.html`.

## Estrutura do projeto

```
├── gerar_dashboard.py         # gerador principal (fluxo atual)
├── template_dashboard.html    # molde visual do dashboard
├── Gerar Dashboard.bat        # atalho de duplo-clique (Windows)
├── COMO USAR.txt              # guia prático para uso semanal
├── requirements.txt
├── WK/
│   ├── assets/                # Chart.js + datalabels (vendorizados, embutidos no HTML)
│   └── data/                  # histórico por semana (WK<nn>_<ano>.json), gerado em runtime
└── legacy/
    ├── app.py                 # versão Streamlit antiga, mantida só como referência de cálculo
    └── .streamlit/config.toml
```

`WK/Dashboard SCM WK<nn>.html` e o conteúdo de `WK/data/` são gerados em runtime e
**não fazem parte do repositório** (ver [Dados confidenciais](#dados-confidenciais)).

## Configuração

Não há variáveis de ambiente — todo ajuste é feito no bloco `CONFIG` no topo de
`gerar_dashboard.py`:

| Variável | O que faz | Padrão |
|---|---|---|
| `PERIODO_INI` / `PERIODO_FIM` | Período de análise | `2026-01-01` até hoje |
| `SLA_ATRASO_DIAS` | Limite (dias) para o ranking de aging por departamento | `15` |
| `COMPRADORES` | Mapa comprador → (rótulo individual, rótulo do time) | Miguel/Adrya/Davi/Luis Gabriel |
| `WK_OVERRIDE` / `ANO_OVERRIDE` | Força um número de semana/ano específico | `None` (automático) |
| `PO_DRILL_FULL_COLUMNS` | Incluir todas as colunas do SC7 no drill-down de POs | `True` |

## Dados confidenciais

As planilhas de entrada (`*.xlsx`), os dashboards `.html` gerados e o histórico em
`WK/data/*.json` contêm dados internos de compras/fornecedores da empresa e por
isso **estão no `.gitignore`** — nunca são versionados. Ao clonar o repositório,
essas pastas ficam vazias até você colocar sua própria planilha e gerar seu
próprio dashboard localmente.

## Fluxo legado (Streamlit)

`legacy/app.py` é a versão anterior (dashboard Streamlit interativo, sem geração
de HTML estático). Foi mantida apenas como referência de cálculo; **não é o fluxo
recomendado**. Para rodar:

```bash
streamlit run legacy/app.py
```

Rode a partir da raiz do projeto — o script lê `Solicitações.xlsx` e
`Relatório de Compras.xlsx` relativos à pasta onde o comando é executado.

## Troubleshooting

- **"Ocorreu um erro, ou o Python não foi encontrado"** (ao rodar o `.bat`): instale
  o Python e confirme que `python` ou `py` funcionam no terminal (`python --version`).
- **Planilha não encontrada**: confirme que o arquivo `Relatório de SCs*.xlsx` está
  na raiz do projeto (mesma pasta do `.bat`/`gerar_dashboard.py`).
- **Dashboard enviado abre sem gráficos**: foi gerado por uma versão antiga do
  gerador (libs carregadas por `<script src="assets/...">` externo). Regere o
  dashboard com a versão atual — o HTML passa a ser 100% autossuficiente.
- **`ModuleNotFoundError`**: rode `pip install -r requirements.txt` no mesmo
  ambiente Python usado para executar o script.
- **Não quero lidar com Python de jeito nenhum**: use o
  [`.exe` da aba Releases](#uso-rápido-sem-instalar-nada) em vez do `.bat`.

## Roadmap

- Nenhum item planejado no momento. Sugestões e ajustes de regra de negócio devem
  ser feitos diretamente no bloco `CONFIG` de `gerar_dashboard.py`.

## Créditos

- Chart.js — MIT License — https://www.chartjs.org/
- chartjs-plugin-datalabels — MIT License — https://github.com/chartjs/chartjs-plugin-datalabels

## Autor

Luis Gabriel de Oliveira
