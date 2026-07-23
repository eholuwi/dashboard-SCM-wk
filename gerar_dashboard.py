# -*- coding: utf-8 -*-
"""
Gerador de Dashboard SCM Semanal (HTML interativo)
==================================================
Lê a planilha "Relatório de SCs" (abas SCM e SC7), calcula todas as métricas
(mesma lógica do app.py) e gera um dashboard HTML moderno, autossuficiente e
interativo (clicar em qualquer dado abre a tabela dos registros por trás dele),
além de acumular um histórico para comparar as semanas (WK).

Uso:
    python gerar_dashboard.py                -> abre janela p/ escolher o Excel
    python gerar_dashboard.py "arquivo.xlsx" -> usa o arquivo informado
    (se cancelar a janela, usa o "Relatório de SCs*.xlsx" mais recente da pasta)
"""

import os
import sys
import re
import csv
import glob
import json
import unicodedata
import webbrowser
from datetime import date, datetime

import pandas as pd
import numpy as np

# =====================================================================
# CONFIG  (ajuste aqui se precisar)
# =====================================================================
# Quando empacotado com PyInstaller (--onefile), os arquivos ficam em dois
# lugares diferentes: os "recursos" (template/assets) são extraídos para uma
# pasta temporária (sys._MEIPASS) embutida no .exe; já a planilha, a saída
# (WK/*.html) e o histórico (WK/data) precisam ficar ao lado do .exe de fato,
# para persistir entre execuções e para o usuário achar o arquivo escolhido.
FROZEN    = getattr(sys, "frozen", False)
APP_DIR   = os.path.dirname(sys.executable) if FROZEN else os.path.dirname(os.path.abspath(__file__))
RES_DIR   = sys._MEIPASS if FROZEN else APP_DIR  # type: ignore[attr-defined]

BASE_DIR   = APP_DIR
WK_DIR     = os.path.join(BASE_DIR, "WK")
DATA_DIR   = os.path.join(WK_DIR, "data")
TEMPLATE   = os.path.join(RES_DIR, "template_dashboard.html")

# Período de análise
PERIODO_INI = date(2026, 1, 1)
PERIODO_FIM = date.today()

# Semana (WK). None = calcula automaticamente pela semana ISO de hoje.
WK_OVERRIDE  = None      # ex.: 27
ANO_OVERRIDE = None      # ex.: 2026

# SLA alvo (dias) para o ranking de aging por departamento
SLA_ATRASO_DIAS = 15

# Compradores do time (para KPIs e Seção 6 de POs).
# chave = substring normalizada encontrada no nome do comprador
# valor = (rótulo individual, rótulo do time)
COMPRADORES = {
    "miguel": ("Miguel",       "Time Miguel"),
    "adrya":  ("Adrya",        "Time Miguel"),
    "davi":   ("Davi",         "Time Davi"),
    "luis":   ("Luis Gabriel", "Time Davi"),
}

# Incluir TODAS as colunas do SC7 no drill-down de POs (botão "ver tudo").
# True = mais completo porém arquivo maior; False = só colunas principais.
PO_DRILL_FULL_COLUMNS = True

# Padroniza as colunas da planilha "Solicitações" (export cru do SCM) para o
# esquema interno usado nos itens em aberto. As chaves são os nomes ORIGINAIS;
# renomeamos a "Descrição" (linha do PC) antes de promover "Descrição Detalhada"
# a "Descrição" (descrição do item) — assim não há colisão de nomes de coluna.
RENAME_SOLIC = {
    "Número da Solicitação": "SC",
    "Descrição": "Descrição PC",
    "Descrição Detalhada": "Descrição",
}

MESES = {1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio",
         6: "Junho", 7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro",
         11: "Novembro", 12: "Dezembro"}
MESES_ABREV = {1: "Jan", 2: "Fev", 3: "Mar", 4: "Abr", 5: "Mai", 6: "Jun",
               7: "Jul", 8: "Ago", 9: "Set", 10: "Out", 11: "Nov", 12: "Dez"}

FAIXAS_AGING = ["1-3 dias", "4-7 dias", "8-11 dias", "12-15 dias", "16-30 dias",
                "31-60 dias", "61-70 dias", "71-80 dias", "81-90 dias",
                "91-100 dias", "> 100 dias"]
# Faixas dentro do SLA (<= SLA_ATRASO_DIAS): não aparecem no gráfico/lista de aging.
FAIXAS_AGING_OCULTAS = ["1-3 dias", "4-7 dias", "8-11 dias", "12-15 dias"]
CORES_AGING = {
    "1-3 dias": "#16A34A", "4-7 dias": "#86EFAC", "8-11 dias": "#FDE68A",
    "12-15 dias": "#FB923C", "16-30 dias": "#FCD34D", "31-60 dias": "#F97316",
    "61-70 dias": "#EF4444", "71-80 dias": "#DC2626", "81-90 dias": "#991B1B",
    "91-100 dias": "#7F1D1D", "> 100 dias": "#450A0A",
}

# SLA "Aprovação do gestor -> Emissão do Pedido (P.O.)": mesmas 10 faixas/cores
# visuais do Aging (verde -> vermelho), só muda a regra de agrupamento.
# Prazo padrão de emissão do Pedido após a aprovação da SC.
SLA_EMISSAO_DIAS = 15
FAIXAS_SLA = ["1-3 dias", "4-7 dias", "8-11 dias", "12-15 dias", "16-19 dias",
              "20-23 dias", "24-27 dias", "28-31 dias", "32-35 dias", "> 35 dias"]
CORES_SLA = {
    "1-3 dias": "#16A34A", "4-7 dias": "#86EFAC", "8-11 dias": "#FDE68A",
    "12-15 dias": "#FB923C", "16-19 dias": "#FCD34D", "20-23 dias": "#F97316",
    "24-27 dias": "#EF4444", "28-31 dias": "#DC2626", "32-35 dias": "#991B1B",
    "> 35 dias": "#7F1D1D",
}


def categorizar_sla(dias):
    """Faixa de SLA (dias corridos entre aprovação da SC e emissão do Pedido)."""
    if pd.isna(dias) or dias < 0:
        return None
    if dias <= 3:   return "1-3 dias"
    if dias <= 7:   return "4-7 dias"
    if dias <= 11:  return "8-11 dias"
    if dias <= 15:  return "12-15 dias"
    if dias <= 19:  return "16-19 dias"
    if dias <= 23:  return "20-23 dias"
    if dias <= 27:  return "24-27 dias"
    if dias <= 31:  return "28-31 dias"
    if dias <= 35:  return "32-35 dias"
    return "> 35 dias"

# =====================================================================
# Helpers
# =====================================================================
def norm(s):
    if not isinstance(s, str):
        s = str(s)
    return unicodedata.normalize("NFKD", s).encode("ASCII", "ignore").decode().lower().strip()


def dep_canon(v):
    """Rótulo canônico de departamento: sem acento, MAIÚSCULO, espaços colapsados.
    Une variações de grafia (Manutenção / MANUTENÇÃO / MANUTENCAO) num único
    balde — usado no ranking, no aging por depto, no drill e no seletor por mês,
    garantindo que os números batam entre si."""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "SEM DEPARTAMENTO"
    s = unicodedata.normalize("NFKD", str(v)).encode("ASCII", "ignore").decode()
    s = " ".join(s.split()).upper().strip()
    return s if s and s.lower() != "nan" else "SEM DEPARTAMENTO"


def sc_key(v):
    """Chave de junção por número de SC, robusta a int/float/str (28844 == '28844')."""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return ""
    if isinstance(v, (int, np.integer)):
        return str(int(v))
    if isinstance(v, float):
        return str(int(v)) if v == int(v) else str(v)
    s = str(v).strip()
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s


def buyer_label(v, idx=0):
    """idx=0 -> individual, idx=1 -> time."""
    n = norm(v)
    for key, labels in COMPRADORES.items():
        if key in n:
            return labels[idx]
    return "Sem comprador"


def categorizar_aging(dias):
    if pd.isna(dias) or dias <= 0:
        return None
    if dias <= 3:   return "1-3 dias"
    if dias <= 7:   return "4-7 dias"
    if dias <= 11:  return "8-11 dias"
    if dias <= 15:  return "12-15 dias"
    if dias <= 30:  return "16-30 dias"
    if dias <= 60:  return "31-60 dias"
    if dias <= 70:  return "61-70 dias"
    if dias <= 80:  return "71-80 dias"
    if dias <= 90:  return "81-90 dias"
    if dias <= 100: return "91-100 dias"
    return "> 100 dias"


def _json_default(o):
    if isinstance(o, np.integer):
        return int(o)
    if isinstance(o, np.floating):
        f = float(o)
        return None if np.isnan(f) else f
    if isinstance(o, np.bool_):
        return bool(o)
    if isinstance(o, (pd.Timestamp, datetime, date)):
        return o.strftime("%d/%m/%Y")
    return str(o)


def cell(v):
    """Converte um valor de célula para algo serializável em JSON."""
    if v is None:
        return ""
    if isinstance(v, (pd.Timestamp, datetime)):
        if pd.isna(v):
            return ""
        return v.strftime("%d/%m/%Y")
    if isinstance(v, float):
        if np.isnan(v):
            return ""
        if v == int(v):
            return int(v)
        return round(v, 2)
    if isinstance(v, (np.integer,)):
        return int(v)
    if isinstance(v, (np.floating,)):
        f = float(v)
        return "" if np.isnan(f) else (int(f) if f == int(f) else round(f, 2))
    if pd.isna(v):
        return ""
    return str(v).strip()


def df_to_records(df, specs):
    """specs = list of dicts {key, label, principal, hidden}. Retorna {cols, rows}."""
    cols_meta = []
    used = []
    for sp in specs:
        if sp["key"] in df.columns:
            cols_meta.append({"key": sp["key"], "label": sp.get("label", sp["key"]),
                              "principal": sp.get("principal", False),
                              "hidden": sp.get("hidden", False)})
            used.append(sp["key"])
    sub = df[used]
    rows = [[cell(v) for v in row] for row in sub.itertuples(index=False, name=None)]
    return {"cols": cols_meta, "rows": rows}


def _mais_recente(patterns, excluir=None):
    """Retorna o .xlsx mais recente que casa com algum dos padrões (glob),
    ignorando temporários do Excel (~$) e um caminho a excluir."""
    excluir = os.path.abspath(excluir) if excluir else None
    cands = []
    for pat in patterns:
        cands += glob.glob(os.path.join(BASE_DIR, pat))
    cands = [c for c in cands
             if not os.path.basename(c).startswith("~$")
             and (excluir is None or os.path.abspath(c) != excluir)]
    return max(cands, key=os.path.getmtime) if cands else None


def find_solicitacoes(excluir=None):
    """Fonte principal de itens/SCs (mais atualizada)."""
    return _mais_recente(["Solicita*.xlsx"], excluir=excluir)


def find_compras(excluir=None):
    """Fonte principal de Pedidos/P.O. (aba SC7 do Relatório de Compras)."""
    return _mais_recente(["Relat*rio de Compras*.xlsx"], excluir=excluir)


def find_scs(excluir=None):
    """Fonte AUXILIAR: só fornece o Departamento (aba SCM do Relatório de SCs)."""
    return _mais_recente(["Relat*rio de SCs*.xlsx"], excluir=excluir)


def _classificar_xlsx(paths):
    """Classifica caminhos .xlsx passados na CLI pelo nome do arquivo, para que
    a ordem dos argumentos não importe (arrastar qualquer um sobre o .bat)."""
    b = {"solic": None, "compras": None, "scs": None}
    for p in paths:
        n = norm(os.path.basename(p))
        if "solicita" in n:
            b["solic"] = b["solic"] or p
        elif "compras" in n:
            b["compras"] = b["compras"] or p
        elif "scs" in n:
            b["scs"] = b["scs"] or p
    return b


def _ask_open(titulo):
    """Abre uma janela de seleção de .xlsx e devolve o caminho (ou None)."""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk(); root.withdraw(); root.attributes("-topmost", True)
        chosen = filedialog.askopenfilename(
            title=titulo, initialdir=BASE_DIR,
            filetypes=[("Planilhas Excel", "*.xlsx"), ("Todos", "*.*")])
        root.destroy()
        return chosen or None
    except Exception as e:
        print(f"(aviso: janela de seleção indisponível: {e})")
        return None


def resolver_arquivos():
    """Localiza as 3 planilhas. Fontes PRINCIPAIS: Solicitações (itens/SCs) e
    Relatório de Compras (Pedidos/P.O.). Fonte AUXILIAR: Relatório de SCs (só o
    Departamento). Prioridade por arquivo: CLI (classificada pelo nome) >
    auto-detecção na pasta > janela de seleção (fallback só p/ as principais).
    Retorna (solic, compras, scs) — 'scs' pode ser None (segue sem Departamento).
    """
    args = [a for a in sys.argv[1:] if a.lower().endswith(".xlsx") and os.path.exists(a)]
    cls = _classificar_xlsx(args)

    solic   = cls["solic"]   or find_solicitacoes()
    compras = cls["compras"] or find_compras()
    scs     = cls["scs"]     or find_scs()

    if solic:
        print(f"  Solicitações (itens/SCs): {os.path.basename(solic)}")
    else:
        solic = _ask_open("Selecione a planilha 'Solicitações' (fonte principal de SCs)")

    if compras:
        print(f"  Relatório de Compras (P.O.): {os.path.basename(compras)}")
    else:
        compras = _ask_open("Selecione a planilha 'Relatório de Compras' (Pedidos/P.O.)")

    if scs:
        print(f"  Relatório de SCs (aux. Departamento): {os.path.basename(scs)}")
    else:
        scs = _ask_open("Selecione 'Relatório de SCs' (auxiliar de Departamento — "
                        "opcional; cancele para seguir sem)")
        if not scs:
            print("  Sem 'Relatório de SCs' — itens ficarão SEM DEPARTAMENTO.")

    return solic, compras, scs


# =====================================================================
# Leitura da planilha
# =====================================================================
def _ler_sc7(path):
    """Lê a aba SC7 de uma planilha (Relatório de Compras ou de SCs). O cabeçalho
    pode ficar após algumas linhas de título — detecta pelas 10 primeiras linhas."""
    raw = pd.read_excel(path, sheet_name="SC7", header=None, nrows=10)
    hdr = 0
    for i in range(len(raw)):
        vals = [norm(x) for x in raw.iloc[i].tolist()]
        if "filial" in vals and ("pedido" in vals or "numero pc" in vals):
            hdr = i
            break
    sc7 = pd.read_excel(path, sheet_name="SC7", header=hdr)
    sc7.columns = [str(c).strip() for c in sc7.columns]
    return sc7


def carregar(scs_path=None, solic_path=None, compras_path=None):
    # aba SCM — fonte AUXILIAR (só o mapa de Departamento). Vem do Relatório de
    # SCs; se ele não existir, segue com SCM vazio (itens sem Departamento).
    if scs_path and os.path.exists(scs_path):
        scm = pd.read_excel(scs_path, sheet_name="SCM", header=0)
        scm.columns = [str(c).strip() for c in scm.columns]
    else:
        scm = pd.DataFrame()

    # aba SC7 (Pedidos/P.O.) — fonte PRINCIPAL: Relatório de Compras. Faz
    # fallback para a aba SC7 do Relatório de SCs (compatibilidade com o fluxo
    # antigo) caso o Relatório de Compras não seja informado.
    sc7_src = compras_path if (compras_path and os.path.exists(compras_path)) else scs_path
    sc7 = _ler_sc7(sc7_src) if (sc7_src and os.path.exists(sc7_src)) else pd.DataFrame()

    # Planilha "Solicitações" — fonte PRINCIPAL dos itens/SCs (export cru do SCM,
    # mais atualizado). Não traz Departamento — ele é cruzado por SC/Solicitante
    # contra a aba SCM (auxiliar) lá em processar().
    solic = None
    if solic_path and os.path.exists(solic_path):
        sx = pd.ExcelFile(solic_path)
        # a 1ª aba costuma ser "Solicitações"; usa a que tiver "Número da Solicitação".
        alvo = sx.sheet_names[0]
        for sn in sx.sheet_names:
            cols = pd.read_excel(solic_path, sheet_name=sn, header=0, nrows=0).columns
            if any(norm(c) == "numero da solicitacao" for c in cols):
                alvo = sn
                break
        solic = pd.read_excel(solic_path, sheet_name=alvo, header=0)
        solic.columns = [str(c).strip() for c in solic.columns]
        solic = solic.rename(columns=RENAME_SOLIC)
        # remove eventual linha-cabeçalho repetida no corpo
        if "SC" in solic.columns:
            solic = solic[solic["SC"].astype(str).str.strip() != "Número da Solicitação"]

    return scm, sc7, solic


# =====================================================================
# Cálculo das métricas + montagem do payload
# =====================================================================
def processar(scm, sc7, solic=None):
    # ---------------- Fonte dos itens em aberto ----------------
    # Mapas de Departamento (canônico) aprendidos na aba SCM (auxiliar): por nº
    # de SC e por Solicitante. Solicitações não traz Departamento — o cruzamento
    # é por SC (exato) e, quando a SC é nova/ausente, por Solicitante (cobre SCs
    # que ainda não existiam no Relatório de SCs).
    scm = scm.copy()
    if "Departamento" in scm.columns:
        scm["Departamento"] = scm["Departamento"].apply(dep_canon)
    dep_map, dep_by_solic = {}, {}
    if "Departamento" in scm.columns:
        if "SC" in scm.columns:
            for sc_val, dep_val in zip(scm["SC"], scm["Departamento"]):
                k = sc_key(sc_val)
                if k and k not in dep_map:
                    dep_map[k] = dep_val
        if "Solicitante" in scm.columns:
            for sol_val, dep_val in zip(scm["Solicitante"], scm["Departamento"]):
                k = norm(sol_val)
                if k and k not in dep_by_solic and dep_val != "SEM DEPARTAMENTO":
                    dep_by_solic[k] = dep_val

    def _resolver_depto(sc_val, sol_val):
        d = dep_map.get(sc_key(sc_val))
        if d and d != "SEM DEPARTAMENTO":
            return d
        return dep_by_solic.get(norm(sol_val), "SEM DEPARTAMENTO")

    # Solicitações (mais atualizada) é a fonte dos itens em aberto; sobrepõe a
    # aba SCM. O Departamento vem dos mapas acima (SC > Solicitante).
    if solic is not None:
        src = solic.copy()
        if "SC" in src.columns:
            _sol = src["Solicitante"] if "Solicitante" in src.columns else pd.Series([None] * len(src), index=src.index)
            src["Departamento"] = [_resolver_depto(sc, sl) for sc, sl in zip(src["SC"], _sol)]
            _match = src["Departamento"] != "SEM DEPARTAMENTO"
            print(f"  Fonte de itens: Solicitações ({len(src):,} linhas) | "
                  f"com depto cruzado: {int(_match.sum()):,} | "
                  f"sem depto: {int((~_match).sum()):,}")
        else:
            src["Departamento"] = "SEM DEPARTAMENTO"
    else:
        src = scm.copy()
        if "Departamento" not in src.columns:
            src["Departamento"] = "SEM DEPARTAMENTO"
        print(f"  Fonte de itens: aba SCM do Relatório ({len(src):,} linhas)")

    src["_aprov"] = pd.to_datetime(src.get("Aprovação"), errors="coerce")
    src["_emiss"] = pd.to_datetime(src.get("Emissão"), errors="coerce")
    src["_status"] = src.get("Status").astype(str).apply(norm)

    mask = (src["_aprov"].dt.date >= PERIODO_INI) & (src["_aprov"].dt.date <= PERIODO_FIM)
    src_f = src[mask].copy()
    registros = len(src_f)

    total_itens = len(src_f)
    total_scs = src_f["SC"].nunique() if "SC" in src_f else 0

    cot = src_f[src_f["_status"] == "cotacao"].copy()

    hoje = pd.Timestamp.now().normalize()
    cot["Aging (dias)"] = (hoje - cot["_aprov"]).dt.days
    cot["__faixa"] = cot["Aging (dias)"].apply(categorizar_aging)
    cot["__mes"] = cot["_emiss"].dt.month.fillna(0).astype(int)
    cot["__buyer"] = cot.get("Comprador").apply(lambda v: buyer_label(v, 0))
    cot["__time"] = cot.get("Comprador").apply(lambda v: buyer_label(v, 1))

    qtd_itens_abertos = len(cot)
    qtd_scs_abertas = cot["SC"].nunique() if "SC" in cot else 0
    pct_itens = round(qtd_itens_abertos / total_itens * 100, 1) if total_itens else 0
    pct_scs = round(qtd_scs_abertas / total_scs * 100, 1) if total_scs else 0

    # meses presentes no período (dentro do ano analisado)
    cot_mes = cot[(cot["_emiss"].dt.year == PERIODO_INI.year)]
    meses_presentes = sorted([m for m in cot_mes["__mes"].unique() if m > 0])

    itens_mes = [int((cot_mes["__mes"] == m).sum()) for m in meses_presentes]
    scs_mes = [int(cot_mes[cot_mes["__mes"] == m]["SC"].nunique()) for m in meses_presentes]
    labels_mes = [MESES_ABREV[m] for m in meses_presentes]

    # departamentos (top 10)
    dep = (cot.groupby("Departamento").size().reset_index(name="q")
           .sort_values("q", ascending=False).head(10)) if "Departamento" in cot else pd.DataFrame(columns=["Departamento", "q"])
    dep_labels = [str(x) for x in dep["Departamento"].tolist()]
    dep_values = [int(x) for x in dep["q"].tolist()]

    # Itens/SCs por mês COM seletor de departamento — derivado do MESMO `cot`
    # que alimenta o ranking, para que a soma dos meses (incl. balde "S/ emissão")
    # de cada depto bata exatamente com a barra do ranking. Nada é descartado.
    meses_mpd = sorted([m for m in cot["__mes"].unique() if m > 0])
    tem_sem_emissao = bool((cot["__mes"] == 0).any())
    mpd_meses = meses_mpd + ([0] if tem_sem_emissao else [])
    mpd_labels = [MESES_ABREV[m] for m in meses_mpd] + (["S/ emissão"] if tem_sem_emissao else [])
    # ordem dos deptos = por qtd de itens desc (mesma lógica do ranking, todos)
    dep_ordem = (cot.groupby("Departamento").size().sort_values(ascending=False).index.tolist()
                 if "Departamento" in cot else [])
    mpd_deptos = ["Todos"] + [str(d) for d in dep_ordem]
    mpd_itens, mpd_scs = {}, {}

    def _serie_itens(frame):
        return [int((frame["__mes"] == m).sum()) for m in mpd_meses]

    def _serie_scs(frame):
        return [int(frame[frame["__mes"] == m]["SC"].nunique()) if "SC" in frame else 0
                for m in mpd_meses]

    mpd_itens["Todos"] = _serie_itens(cot)
    mpd_scs["Todos"] = _serie_scs(cot)
    for d in dep_ordem:
        sub = cot[cot["Departamento"] == d]
        mpd_itens[str(d)] = _serie_itens(sub)
        mpd_scs[str(d)] = _serie_scs(sub)

    # compradores (individual e time)
    ordem_ind = ["Davi", "Miguel", "Luis Gabriel", "Adrya", "Sem comprador"]
    comp_ind = cot["__buyer"].value_counts()
    ci_labels = [l for l in ordem_ind if comp_ind.get(l, 0) > 0]
    ci_values = [int(comp_ind.get(l, 0)) for l in ci_labels]

    ordem_time = ["Time Davi", "Time Miguel", "Sem comprador"]
    comp_time = cot["__time"].value_counts()
    ct_labels = [l for l in ordem_time if comp_time.get(l, 0) > 0]
    ct_values = [int(comp_time.get(l, 0)) for l in ct_labels]

    # aging
    fx = cot["__faixa"].value_counts()
    aging_labels, aging_values, aging_colors = [], [], []
    for f in FAIXAS_AGING:
        if f in FAIXAS_AGING_OCULTAS:
            continue
        q = int(fx.get(f, 0))
        if q > 0:
            aging_labels.append(f); aging_values.append(q); aging_colors.append(CORES_AGING[f])
    aging_total = sum(aging_values)
    aging_pct = [round(v / aging_total * 100, 1) if aging_total else 0 for v in aging_values]

    # aging por departamento (> SLA dias)
    atr = cot[(cot["Aging (dias)"] > SLA_ATRASO_DIAS) & (cot["Aging (dias)"].notna())]
    qtd_atrasados = len(atr)
    dep_atr = (atr.groupby("Departamento").size().reset_index(name="q")
               .sort_values("q", ascending=False).head(10)) if "Departamento" in atr else pd.DataFrame(columns=["Departamento", "q"])
    depatr_labels = [str(x) for x in dep_atr["Departamento"].tolist()]
    depatr_values = [int(x) for x in dep_atr["q"].tolist()]

    # ---------------- SC7 (POs) ----------------
    sc7 = sc7.copy()
    pc_col = "Pedido" if "Pedido" in sc7.columns else ("Numero PC" if "Numero PC" in sc7.columns else None)
    sc7["_dt"] = pd.to_datetime(sc7.get("DT Emissao"), errors="coerce")
    sc7["_buyer"] = sc7.get("Comprador").apply(lambda v: buyer_label(v, 0))
    sc7["_time"] = sc7.get("Comprador").apply(lambda v: buyer_label(v, 1))
    m7 = (sc7["_dt"].dt.date >= PERIODO_INI) & (sc7["_dt"].dt.date <= PERIODO_FIM)
    po = sc7[m7 & (sc7["_buyer"] != "Sem comprador")].copy()

    po["_qent"] = pd.to_numeric(po.get("Qtd.Entregue"), errors="coerce").fillna(0)
    po["__mes"] = po["_dt"].dt.month.fillna(0).astype(int)
    po["__item_entregue"] = (po["_qent"] > 0).astype(int)
    # PO entregue = algum item daquele Pedido foi entregue
    pedidos_entregues = set(po[po["_qent"] > 0][pc_col].astype(str))
    po["__po_entregue"] = po[pc_col].astype(str).isin(pedidos_entregues).astype(int)
    po["Aging (dias)"] = ""  # placeholder p/ manter esquema (não usado em POs)

    pos_emitidos = po[pc_col].nunique()
    itens_pos = len(po)

    po_meses = sorted([m for m in po["__mes"].unique() if m > 0 and (PERIODO_INI.month <= m <= 12)])
    po_meses = [m for m in po_meses if m >= PERIODO_INI.month or True]  # todos no período
    po_meses = sorted([m for m in po["__mes"].unique() if m > 0])
    po_labels = [MESES_ABREV[m] for m in po_meses]

    pos_emit_s, pos_entr_s, pos_ag_s = [], [], []
    it_emit_s, it_entr_s, it_ag_s = [], [], []
    for m in po_meses:
        g = po[po["__mes"] == m]
        ge = g[g["_qent"] > 0]
        pe = g[pc_col].nunique(); pen = ge[pc_col].nunique()
        pos_emit_s.append(int(pe)); pos_entr_s.append(int(pen)); pos_ag_s.append(int(max(0, pe - pen)))
        ie = len(g); ien = len(ge)
        it_emit_s.append(int(ie)); it_entr_s.append(int(ien)); it_ag_s.append(int(max(0, ie - ien)))

    # ---------------- SLA: Aprovação do gestor -> Emissão do P.O. ----------------
    # Mede os dias corridos entre a aprovação da SC e a emissão do Pedido (P.O.).
    # Mesma regra, duas granularidades (o dashboard tem um gráfico para cada):
    #   - por ITEM   -> 1 linha por item de SC, usando o Pedido daquele item.
    #   - por PEDIDO -> 1 linha por Pedido (P.O.); a aprovação considerada é a
    #                   MAIS ANTIGA entre os itens que ele atende (pior espera).
    # Origem de cada peça:
    #   - Data de Aprovação -> Solicitações (src), por item
    #   - vínculo item -> Pedido -> Solicitações (o Relatório de Compras não traz
    #                             o nº da SC preenchido na aba SC7)
    #   - Data de Emissão do Pedido -> Relatório de Compras (autoritativa),
    #                             cruzada pelo nº do Pedido (Numero PC)
    #   - SLA (dias) = Emissão do Pedido - Aprovação
    pc_solic = "Numero PC" if "Numero PC" in src_f.columns else (
               "Pedido" if "Pedido" in src_f.columns else None)

    # Pedido -> emissão mais antiga (Relatório de Compras)
    if pc_solic and "Numero PC" in sc7.columns:
        comp = sc7[["Numero PC", "DT Emissao"]].copy()
        comp["__pc"] = comp["Numero PC"].map(sc_key)
        comp["_emiss_po"] = pd.to_datetime(comp["DT Emissao"], errors="coerce")
        comp = comp.dropna(subset=["_emiss_po"])
        comp = comp[comp["__pc"] != ""]
        comp_min = comp.groupby("__pc", as_index=False)["_emiss_po"].min()
    else:
        comp_min = pd.DataFrame({"__pc": pd.Series(dtype="object"),
                                 "_emiss_po": pd.Series(dtype="datetime64[ns]")})

    # Base por ITEM: cada linha de Solicitações com o Pedido dela.
    _keep = [c for c in ["SC", "Item", "Descrição", "_aprov", "Departamento",
                         "Comprador", "Solicitante", "Data Necessidade"]
             if c in src_f.columns]
    if pc_solic and pc_solic not in _keep:
        _keep.append(pc_solic)
    sla = src_f[_keep].dropna(subset=["_aprov"]).copy()
    sla["__k"] = sla["SC"].map(sc_key)
    sla = sla[sla["__k"] != ""]
    sla["Pedido"] = sla[pc_solic] if pc_solic else ""
    sla["__pc"] = sla["Pedido"].map(sc_key) if pc_solic else ""
    sla = sla.merge(comp_min, on="__pc", how="left")

    # Base por PEDIDO: 1 linha por P.O. (aprovação mais antiga entre seus itens).
    _ped = sla[sla["__pc"] != ""]
    sla_ped = _ped.sort_values("_aprov").groupby("__pc", as_index=False).first()
    if len(sla_ped):
        _cnt = (_ped.groupby("__pc")
                    .agg(**{"Itens no Pedido": ("__pc", "size"),
                            "SCs no Pedido": ("__k", "nunique")})
                    .reset_index())
        sla_ped = sla_ped.merge(_cnt, on="__pc", how="left")
    else:
        sla_ped["Itens no Pedido"] = 0
        sla_ped["SCs no Pedido"] = 0

    def _derivar_sla(df):
        """Colunas de SLA + classificações auxiliares (comuns às duas visões)."""
        df["Data de Aprovação"] = df["_aprov"]
        df["Data de Emissão do Pedido"] = df["_emiss_po"]
        df["SLA (dias)"] = (df["_emiss_po"] - df["_aprov"]).dt.days
        df["__faixa_sla"] = df["SLA (dias)"].apply(categorizar_sla)
        # Classificações auxiliares (para futuras análises; não viram indicador):
        #  - Crítica  = Data de Necessidade a menos de SLA dias da aprovação.
        #  - Atrasada = passou o SLA desde a aprovação e ainda não há Pedido emitido.
        if "Data Necessidade" in df.columns:
            _nec = pd.to_datetime(df["Data Necessidade"], errors="coerce")
            df["Crítica"] = np.where((_nec - df["_aprov"]).dt.days < SLA_EMISSAO_DIAS, "Sim", "")
        else:
            df["Crítica"] = ""
        _sem_po = df["_emiss_po"].isna()
        _dias_aberto = (hoje - df["_aprov"]).dt.days
        df["Atrasada"] = np.where(_sem_po & (_dias_aberto > SLA_EMISSAO_DIAS), "Sim", "")
        return df

    def _faixas_sla(df):
        """Série do donut: só as faixas com quantidade > 0 (linhas com Pedido)."""
        fxs = df["__faixa_sla"].value_counts()
        labels, values, colors = [], [], []
        for f in FAIXAS_SLA:
            q = int(fxs.get(f, 0))
            if q > 0:
                labels.append(f); values.append(q); colors.append(CORES_SLA[f])
        total = sum(values)
        pct = [round(v / total * 100, 1) if total else 0 for v in values]
        return {"labels": labels, "values": values, "pct": pct,
                "colors": colors, "total": total}

    sla = _derivar_sla(sla)
    sla_ped = _derivar_sla(sla_ped)

    # =====================================================================
    # Registros para drill-down
    # =====================================================================
    cot_principais = ["SC", "Item", "Descrição", "Status", "Departamento", "Comprador",
                      "Emissão", "Aprovação", "Aging (dias)", "Numero PC", "Vlr.Total",
                      "Qtd.Entregue", "Fornecedor"]
    cot_specs = []
    for k in cot_principais:
        cot_specs.append({"key": k, "label": k, "principal": True})
    for k in src.columns:
        if k in cot_principais or str(k).startswith("_"):
            continue
        cot_specs.append({"key": k, "label": k, "principal": False})
    for k, lab in [("__mes", "mês"), ("__faixa", "faixa"), ("__buyer", "comprador"), ("__time", "time")]:
        cot_specs.append({"key": k, "label": lab, "principal": False, "hidden": True})
    cot_records = df_to_records(cot, cot_specs)

    po_principais = ["Numero da SC", "Pedido", "Item", "Produto", "Descricao", "Comprador",
                     "DT Emissao", "Quantidade", "Qtd.Entregue", "Vlr.Total", "Fornecedor"]
    po_specs = [{"key": k, "label": k, "principal": True} for k in po_principais]
    if PO_DRILL_FULL_COLUMNS:
        for k in sc7.columns:
            if k in po_principais or str(k).startswith("_"):
                continue
            po_specs.append({"key": k, "label": k, "principal": False})
    for k in ["__mes", "__item_entregue", "__po_entregue"]:
        po_specs.append({"key": k, "label": k, "principal": False, "hidden": True})
    po_records = df_to_records(po, po_specs)

    # Drill-down do SLA: colunas principais na ordem pedida; o resto aparece só
    # com "Ver todos os campos". As duas visões (item/pedido) usam o mesmo
    # esquema, mudando só as colunas principais.
    def _sla_specs(df, principais, extras):
        specs = [{"key": k, "label": k, "principal": True} for k in principais]
        for k in extras:
            if k in df.columns and k not in principais:
                specs.append({"key": k, "label": k, "principal": False})
        specs.append({"key": "__faixa_sla", "label": "faixa",
                      "principal": False, "hidden": True})
        return specs

    _sla_extras = ["Data Necessidade", "Solicitante", "Departamento", "Comprador",
                   "Crítica", "Atrasada"]
    sla_records = df_to_records(sla, _sla_specs(
        sla, ["SC", "Item", "Descrição", "Data de Aprovação", "Pedido",
              "Data de Emissão do Pedido", "SLA (dias)"], _sla_extras))
    sla_ped_records = df_to_records(sla_ped, _sla_specs(
        sla_ped, ["Pedido", "Data de Emissão do Pedido", "SC", "Data de Aprovação",
                  "SLA (dias)", "Itens no Pedido"],
        ["SCs no Pedido"] + _sla_extras))

    # =====================================================================
    # Payload
    # =====================================================================
    dash = {
        "kpis": {
            "itens_abertos": qtd_itens_abertos, "itens_abertos_pct": pct_itens,
            "scs_abertas": qtd_scs_abertas, "scs_abertas_pct": pct_scs,
            "pos_emitidos": int(pos_emitidos), "itens_pos": int(itens_pos),
        },
        "charts": {
            "itens_mes": {"labels": labels_mes, "values": itens_mes, "meses": meses_presentes},
            "scs_mes": {"labels": labels_mes, "values": scs_mes, "meses": meses_presentes},
            "deptos": {"labels": dep_labels, "values": dep_values},
            "mes_por_depto": {"deptos": mpd_deptos, "labels": mpd_labels, "meses": mpd_meses,
                              "itens": mpd_itens, "scs": mpd_scs},
            "comp_ind": {"labels": ci_labels, "values": ci_values},
            "comp_time": {"labels": ct_labels, "values": ct_values},
            "aging": {"labels": aging_labels, "values": aging_values, "pct": aging_pct,
                      "colors": aging_colors, "total": aging_total},
            "aging_depto": {"labels": depatr_labels, "values": depatr_values, "atrasados": qtd_atrasados},
            "po_pos": {"labels": po_labels, "meses": po_meses,
                       "emitidos": pos_emit_s, "entregues": pos_entr_s, "aguardando": pos_ag_s},
            "po_itens": {"labels": po_labels, "meses": po_meses,
                         "emitidos": it_emit_s, "entregues": it_entr_s, "aguardando": it_ag_s},
            "sla": _faixas_sla(sla),            # por Item
            "sla_ped": _faixas_sla(sla_ped),    # por Pedido (P.O.)
        },
        "records": {"cot": cot_records, "po": po_records,
                    "sla": sla_records, "sla_ped": sla_ped_records},
        "meta": {"registros": registros},
    }
    snap = {
        "kpis": dash["kpis"],
        "atrasados_15": qtd_atrasados,
        "aging_total": aging_total,
    }
    return dash, snap


# =====================================================================
# Histórico / snapshots
# =====================================================================
# Séries guardadas no snapshot além dos KPIs. Só os agregados (labels/values),
# alguns KB — é o que permite comparar GRÁFICOS entre semanas, e não apenas
# números. Os registros do drill-down nunca vão para o histórico.
CHARTS_NO_SNAPSHOT = ["aging", "deptos", "comp_ind", "comp_time", "sla", "sla_ped"]


def _resumir_charts(dash):
    resumo = {}
    for chave in CHARTS_NO_SNAPSHOT:
        ch = dash.get("charts", {}).get(chave)
        if not ch or "labels" not in ch:
            continue
        resumo[chave] = {"labels": ch["labels"], "values": ch.get("values", [])}
    return resumo


def salvar_snapshot(wk, ano, label, gerado, periodo, snap, dash=None):
    os.makedirs(DATA_DIR, exist_ok=True)
    payload = {"wk": wk, "ano": ano, "label": label, "gerado_em": gerado,
               "periodo": periodo, **snap}
    if dash is not None:
        payload["charts"] = _resumir_charts(dash)
    with open(os.path.join(DATA_DIR, f"WK{wk:02d}_{ano}.json"), "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2, default=_json_default)


def carregar_historico():
    hist = []
    for fp in glob.glob(os.path.join(DATA_DIR, "WK*_*.json")):
        try:
            with open(fp, "r", encoding="utf-8") as f:
                hist.append(json.load(f))
        except Exception:
            pass
    hist.sort(key=lambda x: (x.get("ano", 0), x.get("wk", 0)))
    return hist


# =====================================================================
# Render
# =====================================================================
def _ler_asset(nome):
    """Lê um arquivo de WK/assets/ como texto (para embutir inline no HTML)."""
    caminho = os.path.join(RES_DIR, "WK", "assets", nome)
    with open(caminho, "r", encoding="utf-8") as f:
        return f.read()


def render(dash, hist, wk, ano, label, gerado, arquivo, template=None, saida=None):
    """Injeta os dados no template e grava o HTML.

    `template` permite gerar com o layout novo (template_moderno.html) sem
    mexer no fluxo manual, que continua usando template_dashboard.html."""
    caminho_template = template or TEMPLATE
    with open(caminho_template, "r", encoding="utf-8") as f:
        html = f.read()

    # Embute as bibliotecas Chart.js DENTRO do HTML (autossuficiente).
    # Assim o .html funciona sozinho, sem a pasta assets/ e sem internet.
    html = html.replace("/*__CHARTJS_LIB__*/", _ler_asset("chart.umd.min.js"))
    html = html.replace("/*__DATALABELS_LIB__*/", _ler_asset("chartjs-plugin-datalabels.min.js"))

    dash["meta"].update({
        "wk": wk, "ano": ano, "label": label, "gerado_em": gerado,
        "periodo_ini": PERIODO_INI.strftime("%d/%m/%Y"),
        "periodo_fim": PERIODO_FIM.strftime("%d/%m/%Y"),
        "arquivo": os.path.basename(arquivo),
    })
    blob = "window.DASH = %s;\nwindow.HIST = %s;" % (
        json.dumps(dash, ensure_ascii=False, default=_json_default),
        json.dumps(hist, ensure_ascii=False, default=_json_default),
    )
    # evita fechar o <script> caso algum texto contenha "</..."
    blob = blob.replace("</", "<\\/")
    html = html.replace("/*__DASH_JSON__*/", blob)
    html = html.replace("__WK_LABEL__", label)
    html = html.replace("__PERIODO_INI__", dash["meta"]["periodo_ini"])
    html = html.replace("__PERIODO_FIM__", dash["meta"]["periodo_fim"])
    html = html.replace("__REGISTROS__", f"{dash['meta']['registros']:,}".replace(",", "."))
    html = html.replace("__GERADO__", gerado)
    html = html.replace("__ARQUIVO__", dash["meta"]["arquivo"])
    html = html.replace("__SLA__", str(SLA_ATRASO_DIAS))

    out = saida or os.path.join(WK_DIR, f"Dashboard SCM {label}.html")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with open(out, "w", encoding="utf-8") as f:
        f.write(html)
    return out


# =====================================================================
# Main
# =====================================================================
def main():
    print("=" * 60)
    print(" Gerador de Dashboard SCM Semanal")
    print("=" * 60)

    print("Localizando planilhas ...")
    solic_arq, compras_arq, scs_arq = resolver_arquivos()
    if not solic_arq and not scs_arq:
        print("ERRO: nenhuma fonte de itens (Solicitações/Relatório de SCs) encontrada.")
        input("Pressione Enter para sair..."); return

    print("Lendo planilhas ...")
    scm, sc7, solic = carregar(scs_arq, solic_arq, compras_arq)
    print(f"  SCM: {len(scm):,} linhas | SC7: {len(sc7):,} linhas"
          + (f" | Solicitações: {len(solic):,} linhas" if solic is not None else ""))

    print("Calculando métricas ...")
    dash, snap = processar(scm, sc7, solic)
    k = dash["kpis"]
    print(f"  Itens abertos: {k['itens_abertos']} | SCs abertas: {k['scs_abertas']} "
          f"| POs: {k['pos_emitidos']} | Itens c/ PO: {k['itens_pos']}")

    today = date.today()
    iso = today.isocalendar()
    wk = WK_OVERRIDE if WK_OVERRIDE else iso[1]
    ano = ANO_OVERRIDE if ANO_OVERRIDE else iso[0]
    label = f"WK{wk:02d}"
    gerado = datetime.now().strftime("%d/%m/%Y %H:%M")
    periodo = f"{PERIODO_INI.strftime('%d/%m/%Y')} a {PERIODO_FIM.strftime('%d/%m/%Y')}"

    print(f"Semana: {label} ({ano})")
    salvar_snapshot(wk, ano, label, gerado, periodo, snap)
    hist = carregar_historico()

    if not os.path.exists(TEMPLATE):
        print(f"ERRO: template não encontrado: {TEMPLATE}")
        input("Pressione Enter para sair..."); return

    fontes = [os.path.basename(p) for p in (solic_arq, compras_arq, scs_arq) if p]
    fonte = " + ".join(fontes) if fontes else "—"
    out = render(dash, hist, wk, ano, label, gerado, fonte)
    size_mb = os.path.getsize(out) / 1e6
    print(f"Dashboard gerado: {out}  ({size_mb:.1f} MB)")

    if not os.environ.get("DASH_NO_OPEN"):
        try:
            webbrowser.open("file:///" + out.replace("\\", "/"))
        except Exception as e:
            print(f"(abra manualmente: {out})  [{e}]")
    print("Concluído.")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback
        traceback.print_exc()
        input("\nOcorreu um erro. Pressione Enter para sair...")
