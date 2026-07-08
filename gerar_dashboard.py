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
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
WK_DIR     = os.path.join(BASE_DIR, "WK")
DATA_DIR   = os.path.join(WK_DIR, "data")
TEMPLATE   = os.path.join(BASE_DIR, "template_dashboard.html")

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

# =====================================================================
# Helpers
# =====================================================================
def norm(s):
    if not isinstance(s, str):
        s = str(s)
    return unicodedata.normalize("NFKD", s).encode("ASCII", "ignore").decode().lower().strip()


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


def find_xlsx():
    cands = glob.glob(os.path.join(BASE_DIR, "Relat*rio de SCs*.xlsx"))
    cands = [c for c in cands if not os.path.basename(c).startswith("~$")]
    if not cands:
        cands = [c for c in glob.glob(os.path.join(BASE_DIR, "*.xlsx"))
                 if not os.path.basename(c).startswith("~$")]
    if not cands:
        return None
    return max(cands, key=os.path.getmtime)


def pick_file():
    """Argumento de linha de comando > janela de seleção > mais recente."""
    if len(sys.argv) > 1 and os.path.exists(sys.argv[1]):
        return sys.argv[1]
    chosen = None
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk(); root.withdraw(); root.attributes("-topmost", True)
        chosen = filedialog.askopenfilename(
            title="Selecione a planilha 'Relatório de SCs' da semana",
            initialdir=BASE_DIR,
            filetypes=[("Planilhas Excel", "*.xlsx"), ("Todos", "*.*")])
        root.destroy()
    except Exception as e:
        print(f"(aviso: janela de seleção indisponível: {e})")
    if chosen:
        return chosen
    auto = find_xlsx()
    if auto:
        print(f"Nenhum arquivo escolhido — usando o mais recente: {os.path.basename(auto)}")
    return auto


# =====================================================================
# Leitura da planilha
# =====================================================================
def carregar(path):
    xls = pd.ExcelFile(path)

    # aba SCM (header linha 0)
    scm = pd.read_excel(path, sheet_name="SCM", header=0)
    scm.columns = [str(c).strip() for c in scm.columns]

    # aba SC7 — o cabeçalho fica após algumas linhas de título; detecta.
    raw = pd.read_excel(path, sheet_name="SC7", header=None, nrows=10)
    hdr = 0
    for i in range(len(raw)):
        vals = [norm(x) for x in raw.iloc[i].tolist()]
        if "filial" in vals and ("pedido" in vals or "numero pc" in vals):
            hdr = i
            break
    sc7 = pd.read_excel(path, sheet_name="SC7", header=hdr)
    sc7.columns = [str(c).strip() for c in sc7.columns]

    return scm, sc7


# =====================================================================
# Cálculo das métricas + montagem do payload
# =====================================================================
def processar(scm, sc7):
    # ---------------- SCM ----------------
    scm = scm.copy()
    scm["_aprov"] = pd.to_datetime(scm.get("Aprovação"), errors="coerce")
    scm["_emiss"] = pd.to_datetime(scm.get("Emissão"), errors="coerce")
    scm["_status"] = scm.get("Status").astype(str).apply(norm)

    mask = (scm["_aprov"].dt.date >= PERIODO_INI) & (scm["_aprov"].dt.date <= PERIODO_FIM)
    scm_f = scm[mask].copy()
    registros = len(scm_f)

    total_itens = len(scm_f)
    total_scs = scm_f["SC"].nunique() if "SC" in scm_f else 0

    cot = scm_f[scm_f["_status"] == "cotacao"].copy()

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

    # =====================================================================
    # Registros para drill-down
    # =====================================================================
    cot_principais = ["SC", "Item", "Descrição", "Status", "Departamento", "Comprador",
                      "Emissão", "Aprovação", "Aging (dias)", "Numero PC", "Vlr.Total",
                      "Qtd.Entregue", "Fornecedor"]
    cot_specs = []
    for k in cot_principais:
        cot_specs.append({"key": k, "label": k, "principal": True})
    for k in scm.columns:
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
            "comp_ind": {"labels": ci_labels, "values": ci_values},
            "comp_time": {"labels": ct_labels, "values": ct_values},
            "aging": {"labels": aging_labels, "values": aging_values, "pct": aging_pct,
                      "colors": aging_colors, "total": aging_total},
            "aging_depto": {"labels": depatr_labels, "values": depatr_values, "atrasados": qtd_atrasados},
            "po_pos": {"labels": po_labels, "meses": po_meses,
                       "emitidos": pos_emit_s, "entregues": pos_entr_s, "aguardando": pos_ag_s},
            "po_itens": {"labels": po_labels, "meses": po_meses,
                         "emitidos": it_emit_s, "entregues": it_entr_s, "aguardando": it_ag_s},
        },
        "records": {"cot": cot_records, "po": po_records},
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
def salvar_snapshot(wk, ano, label, gerado, periodo, snap):
    os.makedirs(DATA_DIR, exist_ok=True)
    payload = {"wk": wk, "ano": ano, "label": label, "gerado_em": gerado,
               "periodo": periodo, **snap}
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
    caminho = os.path.join(WK_DIR, "assets", nome)
    with open(caminho, "r", encoding="utf-8") as f:
        return f.read()


def render(dash, hist, wk, ano, label, gerado, arquivo):
    with open(TEMPLATE, "r", encoding="utf-8") as f:
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

    out = os.path.join(WK_DIR, f"Dashboard SCM {label}.html")
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

    arquivo = pick_file()
    if not arquivo:
        print("ERRO: nenhuma planilha encontrada/selecionada.")
        input("Pressione Enter para sair..."); return

    print(f"Lendo: {os.path.basename(arquivo)} ...")
    scm, sc7 = carregar(arquivo)
    print(f"  SCM: {len(scm):,} linhas | SC7: {len(sc7):,} linhas")

    print("Calculando métricas ...")
    dash, snap = processar(scm, sc7)
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

    out = render(dash, hist, wk, ano, label, gerado, arquivo)
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
