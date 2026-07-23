# Dashboard SCM — coleta automática pela API

Substitui as exportações manuais de planilha. Os dados vêm direto do SCM
(`http://mansrvapp03:5715/api`), somente leitura.

## Uso no dia a dia

**Duplo clique em `Atualizar Dashboard (API).bat`.** Pronto — nenhuma planilha
precisa ser exportada. O HTML sai em `WK\Dashboard SCM WK<nn>.html`.

Pela linha de comando:

```bat
python coletar.py                 :: coleta completa + gera o dashboard
python coletar.py --sem-itens     :: pula o detalhe item-a-item dos P.O. (rápido, ~45s)
python coletar.py --so-coletar    :: só baixa e mostra os números, não gera HTML
python coletar.py --classico      :: gera com o layout antigo (template_dashboard.html)
```

**Primeira execução demora ~18 min** porque busca os itens de cada pedido um a
um (medido: 3.809 pedidos). As seguintes levam **~45 s**: a composição de um
pedido não muda depois de emitido, então fica em cache
(`WK\cache\itens_pedidos.json`). O que é recoletado sempre é o **status de
entrega**, que vem do cabeçalho numa única chamada.

## De onde vem cada dado

| Antes (planilha exportada à mão) | Agora (API) |
|---|---|
| `Solicitações.xlsx` | `POST /Relatorios/GetSolicitacoesByDate/{comprador}/{ini}/{fim}` — devolve **o mesmo .xlsx**, com as mesmas 47 colunas |
| `Relatório de Compras.xlsx` (aba SC7) | `GET /Pedidos/{ini}/{fim}` (cabeçalho + status) + `GET /Pedidos/ByNumero/{filial}/{num}` (itens) |
| `Relatório de SCs.xlsx` (só o Departamento) | `GET /CentroCusto` — resolve **100%** das linhas, contra a cobertura parcial do cruzamento antigo |

Detalhes e medições em [`../VALIDACAO_API.md`](../VALIDACAO_API.md).

## Arquitetura

```
coletar.py  ──►  scm_api.py  ──GET──►  API do SCM
     │
     │  monta (scm, sc7, solic) — MESMO contrato de carregar()
     ▼
gerar_dashboard.processar()      ← regra de negócio inalterada
     ▼
salvar_snapshot() → WK\data\WK<nn>_<ano>.json
render(template_moderno.html) → WK\Dashboard SCM WK<nn>.html
```

O coletor **não reimplementa nenhuma regra de negócio**: ele só produz as três
tabelas que antes vinham das planilhas. Todo o cálculo de KPI, aging e SLA
continua em `processar()`, no `gerar_dashboard.py`.

## Configuração

| O quê | Onde |
|---|---|
| URL da API | variável de ambiente `SCM_API_URL` (padrão `http://mansrvapp03:5715/api`) |
| Código do usuário do relatório | variável `SCM_COD_USUARIO` (padrão `001054`) |
| Período, SLA, compradores do time | bloco `CONFIG` do `gerar_dashboard.py` — inalterado |

⚠️ O código do usuário precisa ser de um **comprador ou administrador**. Um
solicitante comum só enxerga as próprias SCs (o relatório sai com ~100 linhas em
vez de ~6.300).

## Agendamento (2x/dia)

Task Scheduler do Windows, duas tarefas apontando para o mesmo `.bat`:

```
Programa:      C:\...\Streamlit\Atualizar Dashboard (API).bat
Iniciar em:    C:\...\Streamlit
Gatilhos:      diariamente 07:30  e  13:30
Opções:        "Executar estando o usuário conectado ou não" (se possível)
               "Repetir a tarefa se falhar" a cada 15 min, até 3 vezes
```

Para publicar automaticamente numa pasta de rede, edite o `.bat` e descomente a
linha `set "PUBLICAR=\\servidor\Compras\DashboardSCM"` com o caminho real. O HTML
é autossuficiente (Chart.js embutido, sem internet), então funciona aberto direto
da pasta compartilhada — e continua podendo ser anexado num e-mail.

## Se a API cair

O `.bat` detecta e avisa. O **fluxo manual continua intacto**: exporte as
planilhas como antes e rode `Gerar Dashboard.bat`. Nada foi removido.

## O que muda no dashboard

Layout novo (`template_moderno.html`), com:
- sidebar que **recolhe** (só ícones) e modo **tela cheia** — estado lembrado no navegador;
- **filtros globais** (período, departamento, comprador, situação de aging) que
  recalculam todos os KPIs e gráficos no próprio navegador;
- navegação por seções: Visão Geral · Aging · SLA · Pedidos · Comparar Semanas;
- KPIs com **variação real** vs. a semana anterior (vem do histórico);
- drill-down por clique mantido (tabela + busca + export CSV).

Removidos a pedido: indicadores secundários, "Tabela Inteligente" (últimas
solicitações, busca, paginação) e os blocos inferiores (destaques, alertas,
timeline).

A série de cores dos Pedidos passou de laranja/**verde**/vermelho para
laranja/**azul**/vermelho: em protanopia o verde e o laranja ficavam a ΔE 3,3 —
indistinguíveis para quem tem daltonismo vermelho-verde.
