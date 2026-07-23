# Princípios

Este documento existe para reduzir o consumo de contexto do Claude Code.

Objetivos:

- explicar arquitetura;
- explicar regras de negócio;
- explicar responsabilidades;
- registrar decisões importantes;
- evitar que futuras sessões precisem reler todo o repositório.

Este documento NÃO substitui o código.

Em caso de divergência:

Código > CLAUDE.md > README.

Sempre considere o código como fonte da verdade.

# Dashboard SCM Semanal — Contexto para Claude Code

> Documenta o estado **atual** do código (`gerar_dashboard.py`), não o README nem
> intenções futuras. Em caso de divergência entre este arquivo, o README e o
> código, **o código vence** — ver [Observações de inconsistência](#observações-de-inconsistência).

## O que é

Script Python que lê duas planilhas exportadas manualmente do Protheus (ERP) e
gera um dashboard HTML estático e autossuficiente para o time de Compras (SCM),
enviado semanalmente à gerente.

**A pasta se chama "Streamlit" mas o fluxo atual não usa Streamlit.**
`gerar_dashboard.py` é headless: usa `tkinter` para escolher o arquivo e escreve
um `.html` final. `legacy/app.py` é a versão anterior (Streamlit real, interativa),
mantida só como referência de cálculo — ver [decisões arquiteturais](#decisões-arquiteturais-e-porquês).

## Dois fluxos (o novo é o padrão)

**Fluxo API (`coletar.py`) — padrão desde 23/07/2026.** Puxa tudo da API do SCM;
nenhuma planilha é exportada à mão. Entrada: `Atualizar Dashboard (API).bat`.
Ver [`LEIA-ME_API.md`](./LEIA-ME_API.md) e [`../VALIDACAO_API.md`](../VALIDACAO_API.md).

**Fluxo manual (`gerar_dashboard.py`) — mantido como fallback.** Continua
funcionando exatamente como antes; é o plano B para quando a API estiver fora.
**Não remover.**

Os dois compartilham `processar()`, `salvar_snapshot()` e `render()` — a regra de
negócio existe uma vez só. O coletor apenas monta as três tabelas `(scm, sc7,
solic)` que antes vinham das planilhas, com o mesmo esquema de colunas.

```
Atualizar Dashboard (API).bat → coletar.py
  → scm_api.relatorio_solicitacoes_xlsx()  # .xlsx idêntico ao export manual
  → scm_api.pedidos_periodo() + pedido_itens()   # P.O. (cabeçalho + itens)
  → scm_api.centros_custo()                # Departamento (100% resolvido)
  → (scm, sc7, solic)  ─── mesmo contrato de carregar() ───┐
                                                            ▼
```

### Fluxo manual (fallback)

```
Gerar Dashboard.bat (ou python gerar_dashboard.py [*.xlsx ...])
  → resolver_arquivos()                 # classifica args por nome + auto-detecta
                                        #   as 3 planilhas na pasta; tkinter só p/
                                        #   fallback das fontes principais
  → carregar(scs, solic, compras)       # SCM(aux) + Solicitações(itens) + SC7(POs)
  → processar()                         # KPIs/aging/rankings/SLA → (dash, snap)
  → salvar_snapshot()                   # WK/data/WK<nn>_<ano>.json
  → render()                            # injeta dash+histórico em template_dashboard.html
  → WK/Dashboard SCM WK<nn>.html        # webbrowser.open(), a menos que DASH_NO_OPEN=1
```

**Fontes de dados (desde a reorganização):** as PRINCIPAIS são `Solicitações.xlsx`
(itens/SCs em aberto — mais atualizada) e `Relatório de Compras*.xlsx` (aba `SC7` =
Pedidos/P.O.). `Relatório de SCs*.xlsx` virou AUXILIAR: só a aba `SCM`, usada para
cruzar o **Departamento** por nº de SC e, como fallback, por **Solicitante** (cobre
SCs novas). A auto-detecção casa cada arquivo pelo nome (`Solicita*`, `Relat*rio de
Compras*`, `Relat*rio de SCs*`), então a ordem dos args de CLI não importa.

Tudo isso vive em **um único arquivo**, `gerar_dashboard.py` (~686 linhas): não há
módulos separados para load/processamento/render.

## Estrutura do projeto

```
Streamlit/
├── gerar_dashboard.py         # CONFIG + load + KPIs + render + CLI (fluxo manual)
├── scm_api.py                 # cliente da API do SCM (só GET + 1 POST de relatório)
├── coletar.py                 # adapter API → (scm, sc7, solic) → processar()  [fluxo padrão]
├── template_moderno.html      # layout atual (sidebar, filtros globais, drill) — usado pelo coletar.py
├── template_dashboard.html    # layout anterior — usado pelo fluxo manual e por `--classico`
├── Atualizar Dashboard (API).bat  # entrada do fluxo novo
├── Gerar Dashboard.bat        # entrada do fluxo manual (fallback)
├── LEIA-ME_API.md             # runbook do fluxo API (autoridade p/ uso e agendamento)
├── README.md / COMO USAR.txt  # docs do fluxo manual
├── requirements.txt
├── Gerar Dashboard SCM.spec   # build PyInstaller → dist/*.exe
├── WK/
│   ├── assets/                # Chart.js + datalabels vendorizados (embutidos no HTML)
│   ├── cache/                 # itens_pedidos.json + solicitacoes_api.xlsx (runtime, gitignored)
│   └── data/WK<nn>_<ano>.json # histórico semanal (gerado em runtime, gitignored)
└── legacy/app.py              # Streamlit antigo — referência de cálculo, não estender
```

`WK/cache/itens_pedidos.json` guarda a composição de cada P.O. indefinidamente —
ela não muda depois do pedido emitido. Apagar o arquivo só custa uma recoleta
longa (~15 min); não perde dado. O que é sempre recoletado é o STATUS de entrega.

Fora do escopo desta pasta (não confundir com o app): `Power BI/` (pasta irmã,
iniciativa separada), `graphify-out/` (cache de outra ferramenta).

**Artefatos ad hoc não suportados** (existem no diretório mas não são lidos por
nenhum código): `manutenção.xlsx`/`.csv`, `Dashboard SCM WK29 E Manutenção.html`
— parecem exports manuais filtrados para o depto Manutenção, feitos fora do fluxo
padrão. Não assumir que são inputs do pipeline.

## Contrato de entrada

- **Solicitações.xlsx** (fonte PRINCIPAL de itens/SCs): export mais atualizado,
  substitui `SCM` como fonte de itens abertos; colunas renomeadas via
  `RENAME_SOLIC`. Não traz `Departamento` — é cruzado contra a aba `SCM` por nº de
  SC e, como fallback, por `Solicitante`. Traz `Pedido`/`Numero PC` por linha de SC
  (é daí que sai o vínculo SC→Pedido, ver quirk abaixo).
- **Relatório de Compras *.xlsx** (fonte PRINCIPAL de Pedidos/P.O.): aba `SC7`
  (header autodetectado nas 10 primeiras linhas por conter "filial" +
  "pedido"/"numero pc"). Alimenta os KPIs/gráficos de POs e a DATA de emissão do
  Pedido no SLA. **Quirk:** nesta aba a coluna `Numero da SC` vem quase toda vazia
  — NÃO dá para ligar Pedido↔SC por ela. O link SC→Pedido vem da Solicitações
  (`Numero PC`); a emissão vem daqui, cruzada por `Numero PC`.
- **Relatório de SCs *.xlsx** (AUXILIAR, opcional): usa-se só a aba `SCM` (header
  linha 0) para o mapa `Departamento` (por SC e por Solicitante). Sem ela, itens
  ficam `SEM DEPARTAMENTO`. Fallback histórico: se não houver Relatório de Compras,
  a aba `SC7` desta planilha ainda serve de fonte de POs.
- Headers das planilhas-fonte têm mojibake (cp1252/utf-8) — é dos arquivos, não
  um bug para corrigir. Na aba `SCM` de alguns exports o `Departamento` vem com
  nomes; em cópias antigas pode vir com códigos numéricos (ex.: `1.0`) — é do
  arquivo, não do código.

## Regras de negócio (bloco `CONFIG`, `gerar_dashboard.py:29-97`)

Toda regra de negócio fica hardcoded nesse bloco — **é o único lugar a editar**
para ajustes (não existe `.env`/`.ini`). Principais:

- `PERIODO_INI`/`PERIODO_FIM` — janela de análise (filtro por data de Aprovação).
- `SLA_ATRASO_DIAS` (15) — acima disso, item conta como "atrasado" no ranking por depto.
- `COMPRADORES` — fuzzy-match por substring normalizada → (label individual, label do time).
- `FAIXAS_AGING`/`CORES_AGING` — bucketing e cores do aging (11 faixas; as 4 primeiras,
  dentro do SLA, ficam ocultas na UI).
- `FAIXAS_SLA`/`CORES_SLA`/`SLA_EMISSAO_DIAS` (15) — SLA de emissão do P.O.: 10 faixas
  (1-3 até >35 dias), reusam as cores do aging; nenhuma faixa é ocultada.
- `WK_OVERRIDE`/`ANO_OVERRIDE` — força rótulo de semana em vez do ISO week automático.

Fórmulas centrais (`processar()`, `gerar_dashboard.py:327`):
- Aging = `hoje − Aprovação`. "Aberto" = `Status` normalizado == "cotacao".
- PO "entregue" = pelo menos uma linha do Pedido com `Qtd.Entregue > 0` (parcial conta como entregue).
- Departamento é canonicalizado (`dep_canon`) para unir grafias divergentes de "Manutenção".
- SLA (dias) = emissão do Pedido (Relatório de Compras, 1ª emissão do `Numero PC`)
  − Aprovação (Solicitações). Calculado em **duas granularidades**, cada uma com
  seu gráfico: **por Item** (1 linha por item de SC, com o Pedido daquele item) e
  **por Pedido** (1 linha por P.O., usando a aprovação MAIS ANTIGA entre os itens
  que ele atende). Linhas sem Pedido emitido ficam fora do donut (faixa `None`).
  Classificações auxiliares só em colunas do drill (não viram gráfico):
  `Crítica` = necessidade a < 15d da aprovação; `Atrasada` = > 15d desde a
  aprovação e ainda sem Pedido emitido.

## Payload do dashboard (`dash`, retornado por `processar()`)

```python
dash = {
  "kpis": {  # únicos KPIs que existem hoje — não inventar outros
    "itens_abertos", "itens_abertos_pct", "scs_abertas", "scs_abertas_pct",
    "pos_emitidos", "itens_pos"
  },
  "charts": {  # cada chave alimenta um gráfico específico do template
    "itens_mes", "scs_mes", "deptos", "mes_por_depto",
    "comp_ind", "comp_time", "aging", "aging_depto", "po_pos", "po_itens",
    # SLA Aprovação→Emissão do P.O.; mesmo visual do aging, 10 faixas.
    # Duas granularidades, um donut para cada (centro: "Itens" / "Pedidos"):
    "sla",      # por Item
    "sla_ped"   # por Pedido (P.O.)
  },
  # drill-down (df_to_records)
  "records": {"cot": {...}, "po": {...}, "sla": {...}, "sla_ped": {...}},
  "meta": {"registros", "wk", "ano", "label", "gerado_em", "periodo_ini", "periodo_fim", "arquivo"},
}
```
`snap` (salvo no histórico) = `{"kpis", "atrasados_15", "aging_total"}` **e, quando
`salvar_snapshot(..., dash=dash)` recebe o dash, também `"charts"`** — só os
agregados (`labels`/`values`) de `CHARTS_NO_SNAPSHOT` (aging, deptos, comp_ind,
comp_time, sla, sla_ped), poucos KB. É o que permite comparar *gráficos* entre
semanas, e não só números. Snapshots antigos não têm essa chave; a UI de
comparação trata a ausência.

## Contrato do template

`render(dash, hist, wk, ano, label, gerado, arquivo, template=None, saida=None)`
faz `str.replace()` de:
- Placeholders de texto: `__WK_LABEL__`, `__PERIODO_INI__`, `__PERIODO_FIM__`,
  `__REGISTROS__`, `__GERADO__`, `__ARQUIVO__`, `__SLA__`.
- Blocos de script: `/*__CHARTJS_LIB__*/`, `/*__DATALABELS_LIB__*/` (libs inline)
  e `/*__DASH_JSON__*/` → `window.DASH = {...}; window.HIST = [...]`.

**Os dois templates obedecem ao mesmo contrato** — `template` escolhe qual usar.
Qualquer novo placeholder ou chave em `dash["charts"]`/`dash["kpis"]` precisa de
contraparte em `render()` **e nos dois templates**; os lados quebram
silenciosamente se ficarem fora de sincronia (não há schema/validação).

### `template_moderno.html` (layout atual)

Diferença central em relação ao anterior: **os gráficos e KPIs são agregados no
navegador**, a partir de `DASH.records`, e não lidos prontos de `DASH.charts`.
Isso é o que faz os filtros globais (período, departamento, comprador, aging)
valerem para a tela inteira.

Isso **não duplica regra de negócio**: os rótulos caros (`__faixa`, `__buyer`,
`__time`, `__mes`, `__faixa_sla`, `Aging (dias)`, `Departamento`) já vêm
calculados do Python como colunas dos registros — o JS só filtra e conta.
`DASH.charts` continua sendo usado para a *ordem* e as *cores* das categorias.
Se você mudar as colunas de `df_to_records`, confira o mapa `IX` no template.

Ao adicionar um gráfico, use `draw(id, cfg)` (destrói a instância anterior) e
`vazio(id, true)` para o estado sem dados — **nunca** substitua o `innerHTML` do
contêiner do canvas, senão o gráfico não volta quando o filtro é limpo.

O elemento `#__diag` (oculto) recebe um JSON com contagens e `window.__errs` ao
final da carga — é como validar o HTML sem abrir o navegador (ver "Como testar").

## Como rodar / testar

```bash
# Fluxo padrão (API) — não precisa de planilha nenhuma:
DASH_NO_OPEN=1 python coletar.py
DASH_NO_OPEN=1 python coletar.py --sem-itens --so-coletar   # ~45s, só os números

# Fluxo manual (fallback). Passar os arquivos como args evita os diálogos tkinter
# (a ordem não importa; são classificados pelo nome):
DASH_NO_OPEN=1 python gerar_dashboard.py "Solicitações.xlsx" "Relatório de Compras.xlsx"
```

⚠️ **A 1ª coleta leva ~15 min** (busca os itens de ~3.800 pedidos, um a um). As
seguintes levam ~1 min por causa do cache. Ao iterar, use `--sem-itens`.

Validação do HTML sem abrir o navegador — `#__diag` traz contagens e erros:

```bash
"/c/Program Files/Google/Chrome/Application/chrome.exe" --headless=new \
  --disable-gpu --virtual-time-budget=8000 --dump-dom \
  "file:///C:/.../WK/Dashboard SCM WK30.html" | grep -o '__diag[^<]*'
# esperado: "erros":[]  e contagens > 0
```
Sem suíte de testes — validar rodando com dado real e inspecionando o HTML
gerado em `WK/` (console do navegador expõe `window.DASH`/`window.HIST`). Dá para
validar o JS headless: Chrome `--headless=new --dump-dom` + checar `Chart.getChart('c_sla')`
e `window.__errs`. Ler os `.xlsx` reais é lento (~2-3 min); ao iterar em `processar()`,
vale cachear os DataFrames de `carregar()` num pickle e recarregar.

## Decisões arquiteturais e porquês

- **Streamlit abandonado como fluxo principal** → gera dashboard interativo mas
  não produz um arquivo único compartilhável por e-mail; o requisito real era
  "um HTML que a gerente abre sozinha", daí a reescrita para script + template
  estático. `legacy/app.py` foi deixado congelado como referência de cálculo, não
  para uso — duplica a lógica de negócio de forma independente, então **não
  reflita mudanças de KPI feitas em `gerar_dashboard.py` nele**.
- **Integração pela API do SCM (23/07/2026)** → substituiu a exportação manual.
  A premissa antiga ("o sistema não expõe API acessível") estava errada: a API em
  `mansrvapp03:5715` é anônima na rede interna e um dos endpoints **devolve o
  próprio .xlsx** que era exportado à mão. Medições em `../VALIDACAO_API.md`.
- **O coletor não recalcula nada** → ele produz `(scm, sc7, solic)` no mesmo
  esquema das planilhas e entrega a `processar()`. Foi a escolha deliberada para
  não criar uma segunda implementação da regra de negócio, e é o que permitiu
  validar a troca comparando KPIs lado a lado.
- **Fluxo manual preservado** → é o plano B quando a API estiver fora, e a
  referência de conferência. Não removê-lo é intencional.
- **Histórico em JSON simples (`WK/data/*.json`)**, um arquivo por semana → evita
  dependência de banco de dados para um relatório semanal de baixo volume;
  suficiente para alimentar a aba "Comparar Semanas".
- **Config centralizada no bloco `CONFIG` do próprio script** → time pequeno,
  único mantenedor; um arquivo de config externo (YAML/INI) adicionaria
  indireção sem benefício real hoje.
- **Sem testes automatizados** → volume de mudanças e usuários é baixo; validação
  manual (rodar + inspecionar HTML) é o método aceito até agora.
- **Migração para Power BI é iniciativa separada** → resolve o risco de "bus
  factor" (só uma pessoa entende o pipeline Python), mas é tratada como projeto
  à parte (`Power BI/`), não uma reforma deste código.

## Fora do escopo / não implementado

- Escrita na API do SCM. O cliente só faz GET (mais o POST do relatório, que é
  consulta). **Nunca** chamar os endpoints de escrita — criam SC, geram pedido no
  Protheus, disparam e-mail.
- Quantidade entregue por item vinda da API: não existe endpoint com esse campo.
  Vem da coluna `Qtd.Entregue` do relatório quando o item está lá; para os demais
  é derivada do `STATUS` do pedido (`ATENDIDO`/`PARCIALMENTE ATENDIDO`), que
  concordou com o critério antigo em 97% dos P.O. na validação.
- Suíte de testes automatizados.
- Qualquer uso do arquivo `manutenção.xlsx`/`.csv` dentro do pipeline — hoje é
  puramente manual, sem função associada no código.
- Migração/paridade com Power BI — projeto paralelo, não uma feature deste script.

## Observações de inconsistência

- `.claude/settings.local.json` (allowlist local, gitignored) só documenta o uso
  de `DASH_NO_OPEN=1` com o argumento do "Relatório de SCs" — nunca com o segundo
  argumento (Solicitações). Não assumir que o fluxo com 2 arquivos foi validado
  em sessão headless; testar antes de depender disso.
- README/COMO USAR.txt não mencionam `manutenção.xlsx`/`.csv` nem o HTML
  "E Manutenção" — são artefatos presentes no diretório sem documentação nem
  código associado; tratados aqui como ad hoc (ver seção acima).
