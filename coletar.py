# -*- coding: utf-8 -*-
"""
Coletor SCM — substitui as exportações manuais pela API
=======================================================
Monta as MESMAS três estruturas que `carregar()` produz a partir das planilhas
— `(scm, sc7, solic)` — só que puxando tudo da API. Assim `processar()` e todo
o resto do pipeline continuam valendo sem alteração.

    scm    (auxiliar)  Departamento por SC, vindo de /CentroCusto  (antes: 3ª planilha)
    solic  (principal) itens/SCs, vindo do relatório .xlsx da API  (antes: Solicitações.xlsx)
    sc7    (principal) Pedidos/P.O., vindo de /Pedidos             (antes: Relatório de Compras.xlsx)

Uso:
    python coletar.py                 # coleta e gera o dashboard da semana
    python coletar.py --sem-itens     # pula o detalhe item-a-item dos P.O. (mais rápido)
    python coletar.py --so-coletar    # só baixa e mostra o resumo, não gera HTML
"""

import io
import json
import os
import sys
import time
import unicodedata
from datetime import date, datetime

import pandas as pd

import scm_api
import gerar_dashboard as G

# Código de usuário usado no relatório de Solicitações. Precisa ser de alguém
# com permissão ampla (comprador/admin) — um solicitante comum só enxerga as
# próprias SCs. Ver FASE0_VALIDACAO_API.md §1.
COD_USUARIO = os.environ.get("SCM_COD_USUARIO", "001054")

CACHE_DIR = os.path.join(G.BASE_DIR, "WK", "cache")
CACHE_ITENS = os.path.join(CACHE_DIR, "itens_pedidos.json")


def _norm(s):
    return unicodedata.normalize("NFKD", str(s)).encode("ASCII", "ignore").decode().lower().strip()


def _col(df, alvo):
    """Acha a coluna pelo nome normalizado (os cabeçalhos vêm com acento/mojibake)."""
    for c in df.columns:
        if _norm(c) == alvo:
            return c
    return None


def _texto(v):
    return "" if v is None else str(v).strip()


# =====================================================================
# 1. Solicitações — o .xlsx que a API gera é o mesmo do export manual
# =====================================================================
def baixar_solicitacoes(ini, fim, salvar_em=None):
    bruto = scm_api.relatorio_solicitacoes_xlsx(COD_USUARIO, ini, fim)
    if salvar_em:
        os.makedirs(os.path.dirname(salvar_em), exist_ok=True)
        with open(salvar_em, "wb") as f:
            f.write(bruto)
    solic = pd.read_excel(io.BytesIO(bruto), sheet_name=0, header=0)
    solic.columns = [str(c).strip() for c in solic.columns]
    solic = solic.rename(columns=G.RENAME_SOLIC)
    if "SC" in solic.columns:
        solic = solic[solic["SC"].astype(str).str.strip() != "Número da Solicitação"]
    return solic, bruto


# =====================================================================
# 2. Departamento — mapa por centro de custo (substitui a 3ª planilha)
# =====================================================================
def montar_scm_aux(solic):
    """DataFrame no formato da aba SCM: uma linha por SC com o Departamento.
    Antes o Departamento vinha de um cruzamento SC→Solicitante contra o
    "Relatório de SCs" (cobertura parcial); o centro de custo resolve 100%."""
    mapa = scm_api.centros_custo()
    c_cc = _col(solic, "centro custo")
    if c_cc is None:
        return pd.DataFrame(columns=["SC", "Solicitante", "Departamento"])

    codigos = solic[c_cc].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    depto = codigos.map(mapa).fillna("SEM DEPARTAMENTO")

    c_sol = _col(solic, "solicitante")
    aux = pd.DataFrame({
        "SC": solic["SC"] if "SC" in solic.columns else "",
        "Solicitante": solic[c_sol] if c_sol else "",
        "Departamento": depto,
    })
    nao_resolvidos = int((depto == "SEM DEPARTAMENTO").sum())
    if nao_resolvidos:
        print(f"  (aviso: {nao_resolvidos:,} linhas sem centro de custo reconhecido)")
    return aux.drop_duplicates(subset=["SC"])


# =====================================================================
# 3. Pedidos (P.O.) — /Pedidos + itens, no esquema da aba SC7
# =====================================================================
def _carregar_cache_itens():
    try:
        with open(CACHE_ITENS, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _salvar_cache_itens(cache):
    os.makedirs(CACHE_DIR, exist_ok=True)
    with open(CACHE_ITENS, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False)


def _num_item(v):
    """Normaliza o número do item para cruzar '0001' (Protheus) com 1 (planilha)."""
    try:
        return int(float(str(v).strip()))
    except (TypeError, ValueError):
        return None


def _entregas_do_xlsx(solic):
    """Mapa {(Numero PC, item): Qtd.Entregue} e {Numero PC: SC} vindos do
    relatório de Solicitações — é a única fonte com a quantidade entregue REAL
    (o endpoint de itens do pedido não expõe esse campo)."""
    c_pc = _col(solic, "numero pc")
    c_qe = _col(solic, "qtd.entregue")
    c_it = _col(solic, "item.1")
    if c_pc is None or c_qe is None:
        return {}, {}
    sub = solic[solic[c_pc].notna() & (solic[c_pc].astype(str).str.strip() != "")]
    entregues, sc_por_pc = {}, {}
    itens = sub[c_it] if c_it else [None] * len(sub)
    scs = sub["SC"] if "SC" in sub.columns else [None] * len(sub)
    for pc, it, qe, sc in zip(sub[c_pc], itens, pd.to_numeric(sub[c_qe], errors="coerce"), scs):
        chave_pc = G.sc_key(pc)
        if not chave_pc:
            continue
        n = _num_item(it)
        if n is not None and pd.notna(qe):
            entregues[(chave_pc, n)] = float(qe)
        if sc is not None and chave_pc not in sc_por_pc:
            sc_por_pc[chave_pc] = G.sc_key(sc)
    return entregues, sc_por_pc


def _itens_dos_pedidos(pedidos, entregues, sc_por_pc, com_itens=True, msg_a_cada=250):
    """Monta as linhas de item de cada P.O. a partir do Protheus (mesma origem da
    aba SC7 do export manual). Usa cache em disco: a composição de um pedido não
    muda depois de emitido — só o STATUS de entrega, que vem do cabeçalho a cada
    coleta.

    Pedidos cujo endpoint de itens devolve lista vazia (eliminados/resíduo, ~10%)
    entram com UMA linha sintética montada do cabeçalho: o pedido existe e precisa
    ser contado em "P.O. emitidos", ainda que sem detalhe de item."""
    cache = _carregar_cache_itens()
    linhas, novos, sem_itens = [], 0, 0
    total = len(pedidos)

    for i, ped in enumerate(pedidos, 1):
        chave = f"{ped['filial']}/{ped['numero']}"
        if chave in cache:
            itens = cache[chave]
        elif com_itens:
            try:
                itens = scm_api.pedido_itens(ped["filial"], ped["numero"])
            except scm_api.SCMApiError as e:
                print(f"    (sem detalhe de {chave}: {e})")
                itens = []
            cache[chave] = itens
            novos += 1
            if novos % msg_a_cada == 0:
                print(f"    ... {i}/{total} pedidos ({novos} novos da API)", flush=True)
                _salvar_cache_itens(cache)
        else:
            itens = []

        entregue_po = ped["status"] in ("ATENDIDO", "PARCIALMENTE ATENDIDO")
        sc_vinculada = sc_por_pc.get(ped["numero"], "")

        if not itens:
            sem_itens += 1
            linhas.append({
                "Filial": ped["filial"], "Item": "", "Numero PC": ped["numero"],
                "DT Emissao": ped["emissao"], "Produto": "", "Descricao": "",
                "Unidade": "", "Quantidade": 0, "Prc Unitario": 0,
                "Vlr.Total": ped.get("total", 0),
                "Qtd.Entregue": 1 if entregue_po else 0,
                "Fornecedor": ped.get("fornecedor", ""), "Razão Social": ped.get("razao", ""),
                "Comprador": ped["comprador"], "Numero da SC": sc_vinculada,
                "Status PO": ped["status"],
            })
            continue

        for it in itens:
            qtd = it.get("C7_QUANT") or 0
            num_pc = _texto(it.get("C7_NUM")) or ped["numero"]
            n_item = _num_item(it.get("C7_ITEM"))
            # Qtd.Entregue real quando o relatório cobre esse item; senão, o
            # STATUS do pedido decide (aproximação documentada na Fase 0).
            qe = entregues.get((num_pc, n_item))
            if qe is None:
                qe = qtd if entregue_po else 0
            linhas.append({
                "Filial":       _texto(it.get("C7_FILIAL")) or ped["filial"],
                "Item":         _texto(it.get("C7_ITEM")),
                "Numero PC":    num_pc,
                "DT Emissao":   ped["emissao"],
                "Produto":      _texto(it.get("C7_PRODUTO")),
                "Descricao":    _texto(it.get("C7_DESCRI")),
                "Unidade":      _texto(it.get("C7_UM")),
                "Quantidade":   qtd,
                "Prc Unitario": it.get("C7_PRECO") or 0,
                "Vlr.Total":    it.get("C7_TOTAL") or 0,
                "Qtd.Entregue": qe,
                "Fornecedor":   _texto(it.get("C7_FORNECE")),
                "Razão Social": _texto(it.get("A2_NOME")),
                "Comprador":    ped["comprador"],
                "Numero da SC": _texto(it.get("C7_XPEDSCM")) or sc_vinculada,
                "Status PO":    ped["status"],
            })

    if novos:
        _salvar_cache_itens(cache)
    return pd.DataFrame(linhas), novos, sem_itens


def montar_sc7(solic, ini, fim, com_itens=True):
    """Monta a tabela de Pedidos no esquema da aba SC7.

    `/Pedidos` é a lista autoritativa de P.O. (na Fase 0 cobriu 100% dos pedidos
    que o dashboard conta) e traz o comprador e o STATUS de entrega. Os itens vêm
    do Protheus via `/Pedidos/ByNumero` — a mesma origem da aba SC7 exportada à
    mão. A Qtd.Entregue real é cruzada do relatório de Solicitações."""
    cabecalhos = scm_api.pedidos_periodo(ini, fim)
    print(f"  /Pedidos: {len(cabecalhos):,} pedidos no período")

    pedidos = []
    for p in cabecalhos:
        num = _texto(p.get("C7_NUM"))
        if not num:
            continue
        pedidos.append({
            "numero": num,
            "filial": _texto(p.get("C7_FILIAL")) or "01",
            "status": _texto(p.get("STATUS")).upper(),
            "comprador": _texto(p.get("COMPRADOR")),
            "emissao": pd.to_datetime(_texto(p.get("C7_EMISSAO")), dayfirst=True, errors="coerce"),
            "total": p.get("TOTAL") or 0,
            "fornecedor": _texto(p.get("C7_FORNECE")),
            "razao": _texto(p.get("A2_NOME")),
        })

    entregues, sc_por_pc = _entregas_do_xlsx(solic)
    print(f"  Qtd.Entregue real disponível para {len(entregues):,} itens do relatório")

    t0 = time.time()
    if com_itens:
        print(f"  montando itens de {len(pedidos):,} pedidos "
              f"(1ª execução é lenta; depois vem do cache) ...", flush=True)
    sc7, novos, sem_itens = _itens_dos_pedidos(pedidos, entregues, sc_por_pc, com_itens=com_itens)
    print(f"  {len(sc7):,} linhas ({novos:,} pedidos baixados agora, resto do cache) "
          f"em {time.time()-t0:.0f}s")
    if sem_itens:
        print(f"  ({sem_itens:,} pedidos sem detalhe de item — contados como P.O., "
              f"1 linha sintética cada)")
    return sc7


# =====================================================================
# Orquestração
# =====================================================================
def coletar(ini=None, fim=None, com_itens=True, salvar_xlsx=True):
    """Devolve `(scm, sc7, solic)` — mesmo contrato de `gerar_dashboard.carregar()`."""
    ini = ini or G.PERIODO_INI.strftime("%Y%m%d")
    fim = fim or G.PERIODO_FIM.strftime("%Y%m%d")

    destino = os.path.join(CACHE_DIR, "solicitacoes_api.xlsx") if salvar_xlsx else None
    print(f"Baixando relatório de Solicitações ({ini}-{fim}, usuário {COD_USUARIO}) ...", flush=True)
    solic, _ = baixar_solicitacoes(ini, fim, salvar_em=destino)
    print(f"  {len(solic):,} linhas")

    print("Resolvendo Departamento por centro de custo ...", flush=True)
    scm = montar_scm_aux(solic)
    print(f"  {len(scm):,} SCs mapeadas")

    print("Coletando Pedidos (P.O.) ...", flush=True)
    sc7 = montar_sc7(solic, ini, fim, com_itens=com_itens)
    print(f"  SC7 montado: {len(sc7):,} linhas")

    return scm, sc7, solic


def main():
    args = sys.argv[1:]
    com_itens = "--sem-itens" not in args
    so_coletar = "--so-coletar" in args

    print("=" * 62)
    print(" Coletor SCM — dados direto da API (sem exportação manual)")
    print("=" * 62)

    if not scm_api.disponivel():
        print(f"ERRO: API não respondeu em {scm_api.BASE_URL}.")
        print("Verifique a rede/VPN, ou gere o dashboard pelo fluxo manual "
              "(Gerar Dashboard.bat com as planilhas).")
        return 1

    t0 = time.time()
    scm, sc7, solic = coletar(com_itens=com_itens)

    print("Calculando métricas ...", flush=True)
    dash, snap = G.processar(scm, sc7, solic)
    k = dash["kpis"]
    print(f"  Itens abertos: {k['itens_abertos']} | SCs abertas: {k['scs_abertas']} "
          f"| POs: {k['pos_emitidos']} | Itens c/ PO: {k['itens_pos']}")

    if so_coletar:
        print(f"\nConcluído em {time.time()-t0:.0f}s (--so-coletar: nenhum HTML gerado).")
        return 0

    hoje = date.today()
    iso = hoje.isocalendar()
    wk = G.WK_OVERRIDE or iso[1]
    ano = G.ANO_OVERRIDE or iso[0]
    label = f"WK{wk:02d}"
    gerado = datetime.now().strftime("%d/%m/%Y %H:%M")
    periodo = f"{G.PERIODO_INI.strftime('%d/%m/%Y')} a {G.PERIODO_FIM.strftime('%d/%m/%Y')}"

    G.salvar_snapshot(wk, ano, label, gerado, periodo, snap, dash=dash)
    hist = G.carregar_historico()

    template = os.path.join(G.RES_DIR, "template_moderno.html")
    if "--classico" in args or not os.path.exists(template):
        template = None
    destino = os.path.join(G.WK_DIR, f"Dashboard SCM {label}.html")
    saida = G.render(dash, hist, wk, ano, label, gerado, "API do SCM",
                     template=template, saida=destino)
    print(f"Dashboard gerado: {saida}  ({os.path.getsize(saida)/1e6:.1f} MB)")
    print(f"Concluído em {time.time()-t0:.0f}s.")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:
        import traceback
        traceback.print_exc()
        input("\nOcorreu um erro. Pressione Enter para sair...")
