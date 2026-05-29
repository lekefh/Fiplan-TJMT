import streamlit as st
import pandas as pd
import sqlite3
import plotly.graph_objects as go
import unicodedata
import re
import io
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from reportlab.lib import colors as rl_c
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                 Paragraph, Spacer, Image as RLImage,
                                 HRFlowable)
from reportlab.lib.enums import TA_CENTER

DB_NAME = 'dados_gestao_integrada.db'
st.set_page_config(page_title="FIPLAN - GESTAO INTEGRADA", layout="wide")
st.markdown(
    "<h2 style='text-align:center;margin-bottom:0'>UO 03101 - TJMT</h2>"
    "<p style='text-align:center;color:#888;margin-top:0'>"
    "Gestao Financeira Integrada - FIPLAN</p>",
    unsafe_allow_html=True
)

MESES_NOMES = ["Jan", "Fev", "Mar", "Abr", "Maio", "Jun",
               "Jul", "Ago", "Set", "Out", "Nov", "Dez"]

MESES_SEM_ACENTO = {
    "JANEIRO": 1, "FEVEREIRO": 2, "MARCO": 3, "ABRIL": 4,
    "MAIO": 5, "JUNHO": 6, "JULHO": 7, "AGOSTO": 8,
    "SETEMBRO": 9, "OUTUBRO": 10, "NOVEMBRO": 11, "DEZEMBRO": 12
}

CATEGORIAS_REC = [
    "Receita Tributaria", "Receita Patrimonial", "Receita de Servicos",
    "Receita Corrente", "Demais Receitas"
]

BIMESTRES = {
    "1 Bimestre (Jan-Fev)": [1, 2],
    "2 Bimestre (Mar-Abr)": [3, 4],
    "3 Bimestre (Mai-Jun)": [5, 6],
    "4 Bimestre (Jul-Ago)": [7, 8],
    "5 Bimestre (Set-Out)": [9, 10],
    "6 Bimestre (Nov-Dez)": [11, 12]
}

st.markdown(
    "<style>[data-testid='stMetricValue']"
    "{font-size:1.4rem!important;font-weight:700}</style>",
    unsafe_allow_html=True
)


# ---------------------------------------------------------------------------
# AUXILIARES
# ---------------------------------------------------------------------------
def sem_acento(txt):
    return "".join(
        c for c in unicodedata.normalize("NFD", txt)
        if unicodedata.category(c) != "Mn"
    )


def detectar_mes(arquivo):
    m_final = 1
    try:
        df_scan = pd.read_excel(arquivo, nrows=12, header=None)
        for r in range(len(df_scan)):
            for celula in df_scan.iloc[r]:
                txt = sem_acento(str(celula)).upper()
                for nome, num in MESES_SEM_ACENTO.items():
                    if nome in txt:
                        m_final = num
    except Exception:
        pass
    return m_final


def limpar_f(v):
    if pd.isna(v) or str(v).strip() in ("", "-", "nan"):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace('"', "").replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0


def norm(v):
    if pd.isna(v):
        return ""
    s = str(v).strip().replace('"', "").replace("\xa0", "")
    try:
        f = float(s)
        if f == int(f):
            return str(int(f))
    except Exception:
        pass
    return s


# ---------------------------------------------------------------------------
# AUXILIARES LRF
# ---------------------------------------------------------------------------
def safe_div(n, d):
    return (n / d) if d not in [0, None] else 0.0


def periodo_bimestre_extenso(meses_bim):
    meses_bim = sorted(meses_bim)
    nomes = {
        1: "JANEIRO", 2: "FEVEREIRO", 3: "MARCO", 4: "ABRIL",
        5: "MAIO", 6: "JUNHO", 7: "JULHO", 8: "AGOSTO",
        9: "SETEMBRO", 10: "OUTUBRO", 11: "NOVEMBRO", 12: "DEZEMBRO"
    }
    if len(meses_bim) == 1:
        return nomes.get(meses_bim[0], "")
    if len(meses_bim) == 2:
        return "{} E {}".format(nomes.get(meses_bim[0], ""), nomes.get(meses_bim[1], ""))
    return " A ".join([nomes.get(m, str(m)) for m in meses_bim])


def natureza_para_str(v):
    return re.sub(r"\D", "", str(v)) if pd.notna(v) else ""


def modalidade_da_natureza(natureza):
    s = natureza_para_str(natureza)
    return s[2:4] if len(s) >= 4 else ""


def grupo_natureza(natureza):
    s = natureza_para_str(natureza)
    return s[0] if s else ""


def criar_formatos_excel(workbook):
    base = {"font_name": "Arial", "font_size": 8}
    fmt_header = workbook.add_format({**base, "bold": True, "align": "center", "valign": "vcenter", "border": 1, "bg_color": "#BFBFBF", "text_wrap": True})
    fmt_group = workbook.add_format({**base, "bold": True, "border": 1, "bg_color": "#D9D9D9"})
    fmt_subgroup = workbook.add_format({**base, "bold": True, "border": 1, "bg_color": "#EDEDED"})
    fmt_item = workbook.add_format({**base, "border": 1, "indent": 1})
    fmt_subitem = workbook.add_format({**base, "border": 1, "indent": 2})
    fmt_total_text = workbook.add_format({**base, "bold": True, "border": 1, "bg_color": "#EDEDED"})
    fmt_money = workbook.add_format({**base, "border": 1, "num_format": "#,##0.00"})
    fmt_money_bold = workbook.add_format({**base, "bold": True, "border": 1, "bg_color": "#EDEDED", "num_format": "#,##0.00"})
    fmt_money_total = workbook.add_format({**base, "bold": True, "border": 1, "bg_color": "#EDEDED", "num_format": "#,##0.00"})
    fmt_pct = workbook.add_format({**base, "border": 1, "num_format": "0.00%"})
    fmt_pct_bold = workbook.add_format({**base, "bold": True, "border": 1, "bg_color": "#EDEDED", "num_format": "0.00%"})
    fmt_pct_total = workbook.add_format({**base, "bold": True, "border": 1, "bg_color": "#EDEDED", "num_format": "0.00%"})
    return {
        "fmt_header": fmt_header, "fmt_group": fmt_group, "fmt_subgroup": fmt_subgroup,
        "fmt_item": fmt_item, "fmt_subitem": fmt_subitem, "fmt_total_text": fmt_total_text,
        "fmt_money": fmt_money, "fmt_money_bold": fmt_money_bold, "fmt_money_total": fmt_money_total,
        "fmt_pct": fmt_pct, "fmt_pct_bold": fmt_pct_bold, "fmt_pct_total": fmt_pct_total,
    }


def preparar_base_receitas_lrf(df_rec, meses_bim, meses_ate_agora):
    if df_rec.empty:
        return pd.DataFrame(columns=["categoria", "natureza", "previsao_inicial", "previsao_atualizada", "no_bimestre", "ate_bimestre", "saldo", "perc_bim", "perc_ate"])
    df_base = df_rec[~df_rec["codigo_full"].astype(str).str.startswith("9")].copy()
    chaves = ["categoria", "natureza"]
    df_orcado = df_base[df_base["mes"].isin(meses_ate_agora)].groupby(chaves, as_index=False).agg({"orcado": "max"}).rename(columns={"orcado": "previsao_atualizada"})
    df_orcado["previsao_inicial"] = df_orcado["previsao_atualizada"]
    df_bim = df_base[df_base["mes"].isin(meses_bim)].groupby(chaves, as_index=False)["realizado"].sum().rename(columns={"realizado": "no_bimestre"})
    df_ate = df_base[df_base["mes"].isin(meses_ate_agora)].groupby(chaves, as_index=False)["realizado"].sum().rename(columns={"realizado": "ate_bimestre"})
    base = df_orcado.merge(df_bim, on=chaves, how="left").merge(df_ate, on=chaves, how="left").fillna(0)
    base["saldo"] = base["previsao_atualizada"] - base["ate_bimestre"]
    base["perc_bim"] = base.apply(lambda r: safe_div(r["no_bimestre"], r["previsao_atualizada"]), axis=1)
    base["perc_ate"] = base.apply(lambda r: safe_div(r["ate_bimestre"], r["previsao_atualizada"]), axis=1)
    return base


def preparar_deducoes_receitas_lrf(df_rec, meses_bim, meses_ate_agora):
    zero = {"previsao_inicial": 0.0, "previsao_atualizada": 0.0, "no_bimestre": 0.0, "ate_bimestre": 0.0, "saldo": 0.0, "perc_bim": 0.0, "perc_ate": 0.0}
    if df_rec.empty:
        return zero
    df_ded = df_rec[df_rec["codigo_full"].astype(str).str.startswith("9")].copy()
    if df_ded.empty:
        return zero
    prev_atu = float(df_ded[df_ded["mes"].isin(meses_ate_agora)].groupby("codigo_full")["orcado"].max().sum())
    no_bim = float(df_ded[df_ded["mes"].isin(meses_bim)]["realizado"].sum())
    ate_bim = float(df_ded[df_ded["mes"].isin(meses_ate_agora)]["realizado"].sum())
    return {"previsao_inicial": prev_atu, "previsao_atualizada": prev_atu, "no_bimestre": no_bim, "ate_bimestre": ate_bim, "saldo": prev_atu - ate_bim, "perc_bim": safe_div(no_bim, prev_atu), "perc_ate": safe_div(ate_bim, prev_atu)}


def preparar_base_despesas_lrf(df_orc, df_exec, meses_bim, meses_ate_agora):
    if df_orc.empty and df_exec.empty:
        return pd.DataFrame(columns=["natureza", "orcado_inicial", "cred_autorizado", "emp_no_bim", "emp_ate", "liq_no_bim", "liq_ate", "pago_ate", "modalidade", "grupo"])
    meses_orc = sorted(set(df_orc["mes"].tolist()).intersection(set(meses_ate_agora))) if not df_orc.empty else []
    m_ref = max(meses_orc) if meses_orc else max(meses_ate_agora)
    if not df_orc.empty and m_ref in df_orc["mes"].values:
        df_last = df_orc[df_orc["mes"] == m_ref].groupby(["natureza"], as_index=False).agg({"orcado_inicial": "sum", "cred_autorizado": "sum"})
    else:
        df_last = pd.DataFrame(columns=["natureza", "orcado_inicial", "cred_autorizado"])
    if not df_exec.empty:
        df_bim = df_exec[df_exec["mes"].isin(meses_bim)].groupby(["natureza"], as_index=False).agg({"empenhado": "sum", "liquidado": "sum"}).rename(columns={"empenhado": "emp_no_bim", "liquidado": "liq_no_bim"})
        df_ate = df_exec[df_exec["mes"].isin(meses_ate_agora)].groupby(["natureza"], as_index=False).agg({"empenhado": "sum", "liquidado": "sum", "pago": "sum"}).rename(columns={"empenhado": "emp_ate", "liquidado": "liq_ate", "pago": "pago_ate"})
    else:
        df_bim = pd.DataFrame(columns=["natureza", "emp_no_bim", "liq_no_bim"])
        df_ate = pd.DataFrame(columns=["natureza", "emp_ate", "liq_ate", "pago_ate"])
    base = df_last.merge(df_bim, on="natureza", how="outer").merge(df_ate, on="natureza", how="outer").fillna(0)
    base["modalidade"] = base["natureza"].apply(modalidade_da_natureza)
    base["grupo"] = base["natureza"].apply(grupo_natureza)
    return base


def preparar_base_funcional_lrf(df_orc, df_exec, meses_bim, meses_ate_agora):
    if df_orc.empty and df_exec.empty:
        return pd.DataFrame(columns=["funcao", "subfuncao", "orcado_inicial", "cred_autorizado", "emp_no_bim", "emp_ate", "liq_no_bim", "liq_ate"])
    meses_orc = sorted(set(df_orc["mes"].tolist()).intersection(set(meses_ate_agora))) if not df_orc.empty else []
    m_ref = max(meses_orc) if meses_orc else max(meses_ate_agora)
    if not df_orc.empty and m_ref in df_orc["mes"].values:
        df_last = df_orc[df_orc["mes"] == m_ref].groupby(["funcao", "subfuncao"], as_index=False).agg({"orcado_inicial": "sum", "cred_autorizado": "sum"})
    else:
        df_last = pd.DataFrame(columns=["funcao", "subfuncao", "orcado_inicial", "cred_autorizado"])
    if not df_exec.empty:
        df_bim = df_exec[df_exec["mes"].isin(meses_bim)].groupby(["funcao", "subfuncao"], as_index=False).agg({"empenhado": "sum", "liquidado": "sum"}).rename(columns={"empenhado": "emp_no_bim", "liquidado": "liq_no_bim"})
        df_ate = df_exec[df_exec["mes"].isin(meses_ate_agora)].groupby(["funcao", "subfuncao"], as_index=False).agg({"empenhado": "sum", "liquidado": "sum"}).rename(columns={"empenhado": "emp_ate", "liquidado": "liq_ate"})
    else:
        df_bim = pd.DataFrame(columns=["funcao", "subfuncao", "emp_no_bim", "liq_no_bim"])
        df_ate = pd.DataFrame(columns=["funcao", "subfuncao", "emp_ate", "liq_ate"])
    for df_ in [df_last, df_bim, df_ate]:
        for col in ["funcao", "subfuncao"]:
            if col in df_.columns:
                df_[col] = df_[col].astype(str)
    return df_last.merge(df_bim, on=["funcao", "subfuncao"], how="outer").merge(df_ate, on=["funcao", "subfuncao"], how="outer").fillna(0)


def gerar_excel_anexo1(df_rec, meses_bim, meses_ate_agora):
    base = preparar_base_receitas_lrf(df_rec, meses_bim, meses_ate_agora)
    deducoes = preparar_deducoes_receitas_lrf(df_rec, meses_bim, meses_ate_agora)
    periodo = periodo_bimestre_extenso(meses_bim)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet("Anexo_I")
        writer.sheets["Anexo_I"] = worksheet
        f = criar_formatos_excel(workbook)
        worksheet.set_column("A:A", 42); worksheet.set_column("B:C", 18); worksheet.set_column("D:H", 18)
        worksheet.merge_range(0, 0, 1, 0, "RECEITAS", f["fmt_header"])
        worksheet.merge_range(0, 1, 1, 1, "PREVISAO INICIAL", f["fmt_header"])
        worksheet.merge_range(0, 2, 1, 2, "PREVISAO ATUALIZADA (A)", f["fmt_header"])
        worksheet.merge_range(0, 3, 0, 7, "RECEITAS REALIZADAS", f["fmt_header"])
        worksheet.write(1, 3, "NO BIMESTRE\n" + periodo, f["fmt_header"])
        worksheet.write(1, 4, "%\n(B/A)", f["fmt_header"])
        worksheet.write(1, 5, "ATE O BIMESTRE\n" + periodo, f["fmt_header"])
        worksheet.write(1, 6, "%\n(C/A)", f["fmt_header"])
        worksheet.write(1, 7, "SALDO A\nREALIZAR\n(A-C)", f["fmt_header"])
        ordem_categorias = ["Receita Tributaria", "Receita Patrimonial", "Receita de Servicos", "Receita Corrente", "Demais Receitas"]
        grupos = {"RECEITAS CORRENTES": ["Receita Tributaria", "Receita Patrimonial", "Receita de Servicos", "Receita Corrente"], "DEMAIS RECEITAS CORRENTES": ["Demais Receitas"]}
        row = 2
        def write_row(desc, vals, fd, fn, fp):
            nonlocal row
            worksheet.write(row, 0, desc, fd)
            worksheet.write_number(row, 1, vals.get("previsao_inicial", 0), fn)
            worksheet.write_number(row, 2, vals.get("previsao_atualizada", 0), fn)
            worksheet.write_number(row, 3, vals.get("no_bimestre", 0), fn)
            worksheet.write_number(row, 4, vals.get("perc_bim", 0), fp)
            worksheet.write_number(row, 5, vals.get("ate_bimestre", 0), fn)
            worksheet.write_number(row, 6, vals.get("perc_ate", 0), fp)
            worksheet.write_number(row, 7, vals.get("saldo", 0), fn)
            row += 1
        total_geral = {"previsao_inicial": 0, "previsao_atualizada": 0, "no_bimestre": 0, "ate_bimestre": 0, "saldo": 0}
        for nome_grupo, cats in grupos.items():
            df_g = base[base["categoria"].isin(cats)].copy()
            if df_g.empty:
                continue
            soma_g = {"previsao_inicial": float(df_g["previsao_inicial"].sum()), "previsao_atualizada": float(df_g["previsao_atualizada"].sum()), "no_bimestre": float(df_g["no_bimestre"].sum()), "ate_bimestre": float(df_g["ate_bimestre"].sum()), "saldo": float(df_g["saldo"].sum())}
            soma_g["perc_bim"] = safe_div(soma_g["no_bimestre"], soma_g["previsao_atualizada"])
            soma_g["perc_ate"] = safe_div(soma_g["ate_bimestre"], soma_g["previsao_atualizada"])
            write_row(nome_grupo, soma_g, f["fmt_group"], f["fmt_money"], f["fmt_pct"])
            for cat in [c for c in ordem_categorias if c in cats]:
                df_c = df_g[df_g["categoria"] == cat].copy()
                if df_c.empty:
                    continue
                soma_c = {"previsao_inicial": float(df_c["previsao_inicial"].sum()), "previsao_atualizada": float(df_c["previsao_atualizada"].sum()), "no_bimestre": float(df_c["no_bimestre"].sum()), "ate_bimestre": float(df_c["ate_bimestre"].sum()), "saldo": float(df_c["saldo"].sum())}
                soma_c["perc_bim"] = safe_div(soma_c["no_bimestre"], soma_c["previsao_atualizada"])
                soma_c["perc_ate"] = safe_div(soma_c["ate_bimestre"], soma_c["previsao_atualizada"])
                write_row(cat.upper(), soma_c, f["fmt_subgroup"], f["fmt_money"], f["fmt_pct"])
                for _, r in df_c.sort_values("natureza").iterrows():
                    write_row(str(r["natureza"]), {"previsao_inicial": float(r["previsao_inicial"]), "previsao_atualizada": float(r["previsao_atualizada"]), "no_bimestre": float(r["no_bimestre"]), "ate_bimestre": float(r["ate_bimestre"]), "saldo": float(r["saldo"]), "perc_bim": float(r["perc_bim"]), "perc_ate": float(r["perc_ate"])}, f["fmt_item"], f["fmt_money"], f["fmt_pct"])
            for k in ["previsao_inicial", "previsao_atualizada", "no_bimestre", "ate_bimestre", "saldo"]:
                total_geral[k] += soma_g[k]
        total_geral["perc_bim"] = safe_div(total_geral["no_bimestre"], total_geral["previsao_atualizada"])
        total_geral["perc_ate"] = safe_div(total_geral["ate_bimestre"], total_geral["previsao_atualizada"])
        total_final = {k: total_geral.get(k, 0) + deducoes.get(k, 0) for k in ["previsao_inicial", "previsao_atualizada", "no_bimestre", "ate_bimestre", "saldo"]}
        total_final["perc_bim"] = safe_div(total_final["no_bimestre"], total_final["previsao_atualizada"])
        total_final["perc_ate"] = safe_div(total_final["ate_bimestre"], total_final["previsao_atualizada"])
        zero = {"previsao_inicial": 0, "previsao_atualizada": 0, "no_bimestre": 0, "ate_bimestre": 0, "saldo": 0, "perc_bim": 0, "perc_ate": 0}
        for desc, vals in [("SUBTOTAL DA RECEITA (I)", total_geral), ("DEFICIT (II)", zero), ("TOTAL (III) = I + II", total_geral), ("DEDUCOES DA RECEITA (CODIGOS INICIADOS POR 9)", deducoes), ("SALDO DE EXERCICIOS ANTERIORES", zero), ("SUPERAVIT FINANCEIRO", zero), ("TOTAL DA RECEITA (IV)", total_final)]:
            write_row(desc, vals, f["fmt_total_text"], f["fmt_money_bold"], f["fmt_pct_bold"])
        worksheet.freeze_panes(2, 1)
    return output.getvalue()


def gerar_excel_anexo1a(df_orc, df_exec, df_rec, meses_bim, meses_ate_agora):
    base = preparar_base_despesas_lrf(df_orc, df_exec, meses_bim, meses_ate_agora)
    periodo = periodo_bimestre_extenso(meses_bim)
    receita_bim = float(df_rec[df_rec["mes"].isin(meses_bim)]["realizado"].sum()) if not df_rec.empty else 0.0
    receita_ate = float(df_rec[df_rec["mes"].isin(meses_ate_agora)]["realizado"].sum()) if not df_rec.empty else 0.0

    def somar(mask):
        df = base[mask].copy() if not base.empty and len(mask) > 0 else pd.DataFrame()
        v = {c: float(df[c].sum()) if not df.empty and c in df.columns else 0.0 for c in ["orcado_inicial", "cred_autorizado", "emp_no_bim", "emp_ate", "liq_no_bim", "liq_ate", "pago_ate"]}
        v["saldo_emp"] = v["cred_autorizado"] - v["emp_ate"]
        v["saldo_liq"] = v["cred_autorizado"] - v["liq_ate"]
        v["restos"] = 0.0
        return v

    if not base.empty:
        mask_cor = (base["grupo"] == "3") & (base["modalidade"] != "91")
        mask_cor50 = mask_cor & (base["modalidade"] == "50")
        mask_cor90 = mask_cor & (base["modalidade"] == "90")
        mask_cap = (base["grupo"] == "4") & (base["modalidade"] != "91")
        mask_cap90 = mask_cap & (base["modalidade"] == "90")
        mask_intra = (base["modalidade"] == "91")
        mask_all = pd.Series([True] * len(base), index=base.index)
    else:
        mask_cor = mask_cor50 = mask_cor90 = mask_cap = mask_cap90 = mask_intra = mask_all = pd.Series(dtype=bool)

    v_cor = somar(mask_cor); v_cor50 = somar(mask_cor50); v_cor90 = somar(mask_cor90)
    v_cap = somar(mask_cap); v_cap90 = somar(mask_cap90); v_intra = somar(mask_intra)
    v_sub = somar(mask_all); v_zero = {k: 0.0 for k in v_sub}
    v_sup = {**v_zero, "emp_no_bim": receita_bim - v_sub["emp_no_bim"], "emp_ate": receita_ate - v_sub["emp_ate"], "liq_no_bim": max(receita_bim - v_sub["liq_no_bim"], 0), "liq_ate": max(receita_ate - v_sub["liq_ate"], 0), "pago_ate": max(receita_ate - v_sub["pago_ate"], 0)}

    linhas = [
        ("DESPESAS (EXCETO INTRA-ORCAMENTARIAS) (VIII)", somar(mask_cor | mask_cap) if not base.empty else v_zero, "total"),
        ("DESPESAS CORRENTES", v_cor, "grupo"),
        ("Inst. privadas sem fins lucrativos (mod. 50)", v_cor50, "item"),
        ("Outras Desp.Correntes (modalidade 90)", v_cor90, "item"),
        ("DESPESAS DE CAPITAL", v_cap, "grupo"),
        ("Investimentos (modalidade 90)", v_cap90, "item"),
        ("DESPESAS (INTRA-ORCAMENTARIAS) (IX) (91)", v_intra, "grupo"),
        ("SUBTOTAL DESPESAS (X) = (VIII+IX)", v_sub, "total"),
        ("AMORTIZACAO DA DIVIDA / REFINANCIAMENTO (XI)", v_zero, "grupo"),
        ("Amortizacao da Divida Interna", v_zero, "item"),
        ("Amortizacao da Divida Externa", v_zero, "item"),
        ("TOTAL DAS DESPESAS (XII) = (X+XI)", v_sub, "total"),
        ("SUPERAVIT (XIII)", v_sup, "total"),
    ]

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet("Anexo_IA")
        writer.sheets["Anexo_IA"] = worksheet
        f = criar_formatos_excel(workbook)
        worksheet.set_column("A:A", 48); worksheet.set_column("B:K", 16)
        worksheet.merge_range(0, 0, 2, 0, "DESPESAS", f["fmt_header"])
        worksheet.merge_range(0, 1, 2, 1, "DOTACAO INICIAL\n(a)", f["fmt_header"])
        worksheet.merge_range(0, 2, 2, 2, "DOTACAO\nATUALIZADA\n(c)", f["fmt_header"])
        worksheet.merge_range(0, 3, 0, 5, "DESPESAS EMPENHADAS", f["fmt_header"])
        worksheet.write(1, 3, "NO BIMESTRE\n" + periodo, f["fmt_header"]); worksheet.write(1, 4, "ATE O BIMESTRE\n" + periodo, f["fmt_header"]); worksheet.write(1, 5, "Saldo\n(g)=c-f", f["fmt_header"])
        worksheet.merge_range(0, 6, 0, 8, "DESPESAS EXECUTADAS", f["fmt_header"])
        worksheet.write(1, 6, "LIQUIDADAS", f["fmt_header"]); worksheet.write(1, 7, "LIQUIDADAS", f["fmt_header"]); worksheet.write(1, 8, "Saldo\n(i)=c-h", f["fmt_header"])
        worksheet.write(2, 3, ""); worksheet.write(2, 4, ""); worksheet.write(2, 5, "")
        worksheet.write(2, 6, "NO BIMESTRE\n" + periodo, f["fmt_header"]); worksheet.write(2, 7, "ATE O BIMESTRE\n" + periodo, f["fmt_header"]); worksheet.write(2, 8, "", f["fmt_header"])
        worksheet.merge_range(0, 9, 2, 9, "Despesas pagas\nate o mes\n(j)", f["fmt_header"])
        worksheet.merge_range(0, 10, 2, 10, "INSCRITAS EM\nRESTOS A\nPAGAR NAO\nPROCESSADOS (k)", f["fmt_header"])
        row = 3
        for desc, vals, tipo in linhas:
            fd = f["fmt_total_text"] if tipo == "total" else (f["fmt_group"] if tipo == "grupo" else f["fmt_item"])
            fn = f["fmt_money_total"] if tipo == "total" else f["fmt_money"]
            worksheet.write(row, 0, desc, fd)
            for ci, key in enumerate(["orcado_inicial", "cred_autorizado", "emp_no_bim", "emp_ate", "saldo_emp", "liq_no_bim", "liq_ate", "saldo_liq", "pago_ate", "restos"]):
                worksheet.write_number(row, ci + 1, vals.get(key, 0), fn)
            row += 1
        worksheet.freeze_panes(3, 1)
    return output.getvalue()


def gerar_excel_anexo2(df_orc, df_exec, meses_bim, meses_ate_agora):
    base = preparar_base_funcional_lrf(df_orc, df_exec, meses_bim, meses_ate_agora)
    periodo = periodo_bimestre_extenso(meses_bim)
    total_emp = float(base["emp_ate"].sum()) if not base.empty else 0.0
    total_liq = float(base["liq_ate"].sum()) if not base.empty else 0.0
    linhas = []
    if not base.empty:
        for funcao in sorted(base["funcao"].astype(str).unique()):
            df_f = base[base["funcao"].astype(str) == str(funcao)].copy()
            vf = {"orcado_inicial": float(df_f["orcado_inicial"].sum()), "cred_autorizado": float(df_f["cred_autorizado"].sum()), "emp_no_bim": float(df_f["emp_no_bim"].sum()), "emp_ate": float(df_f["emp_ate"].sum()), "perc_emp": safe_div(float(df_f["emp_ate"].sum()), total_emp), "saldo_emp": float(df_f["cred_autorizado"].sum() - df_f["emp_ate"].sum()), "liq_no_bim": float(df_f["liq_no_bim"].sum()), "liq_ate": float(df_f["liq_ate"].sum()), "perc_liq": safe_div(float(df_f["liq_ate"].sum()), total_liq), "saldo_liq": float(df_f["cred_autorizado"].sum() - df_f["liq_ate"].sum()), "restos": 0.0}
            linhas.append(("FUNCAO " + str(funcao), vf, "grupo"))
            for subf in sorted(df_f["subfuncao"].astype(str).unique()):
                df_s = df_f[df_f["subfuncao"].astype(str) == str(subf)].copy()
                vs = {"orcado_inicial": float(df_s["orcado_inicial"].sum()), "cred_autorizado": float(df_s["cred_autorizado"].sum()), "emp_no_bim": float(df_s["emp_no_bim"].sum()), "emp_ate": float(df_s["emp_ate"].sum()), "perc_emp": safe_div(float(df_s["emp_ate"].sum()), total_emp), "saldo_emp": float(df_s["cred_autorizado"].sum() - df_s["emp_ate"].sum()), "liq_no_bim": float(df_s["liq_no_bim"].sum()), "liq_ate": float(df_s["liq_ate"].sum()), "perc_liq": safe_div(float(df_s["liq_ate"].sum()), total_liq), "saldo_liq": float(df_s["cred_autorizado"].sum() - df_s["liq_ate"].sum()), "restos": 0.0}
                linhas.append(("Subfuncao " + str(subf), vs, "item"))
    vt = {"orcado_inicial": float(base["orcado_inicial"].sum()) if not base.empty else 0.0, "cred_autorizado": float(base["cred_autorizado"].sum()) if not base.empty else 0.0, "emp_no_bim": float(base["emp_no_bim"].sum()) if not base.empty else 0.0, "emp_ate": total_emp, "perc_emp": 1.0 if total_emp > 0 else 0.0, "saldo_emp": float(base["cred_autorizado"].sum() - base["emp_ate"].sum()) if not base.empty else 0.0, "liq_no_bim": float(base["liq_no_bim"].sum()) if not base.empty else 0.0, "liq_ate": total_liq, "perc_liq": 1.0 if total_liq > 0 else 0.0, "saldo_liq": float(base["cred_autorizado"].sum() - base["liq_ate"].sum()) if not base.empty else 0.0, "restos": 0.0}
    linhas.append(("TOTAL", vt, "total"))
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet("Anexo_II")
        writer.sheets["Anexo_II"] = worksheet
        f = criar_formatos_excel(workbook)
        worksheet.set_column("A:A", 32); worksheet.set_column("B:L", 14)
        worksheet.merge_range(0, 0, 2, 0, "FUNCAO/\nSUBFUNCAO", f["fmt_header"])
        worksheet.merge_range(0, 1, 2, 1, "DOTACAO\nINICIAL", f["fmt_header"])
        worksheet.merge_range(0, 2, 2, 2, "DOTACAO\nATUALIZADA\n(a)", f["fmt_header"])
        worksheet.merge_range(0, 3, 0, 6, "DESPESA EMPENHADA", f["fmt_header"])
        worksheet.write(1, 3, "NO BIMESTRE\n" + periodo, f["fmt_header"]); worksheet.write(1, 4, "ATE O BIMESTRE\n" + periodo + "\n(b)", f["fmt_header"]); worksheet.write(1, 5, "%\n(b/total b)", f["fmt_header"]); worksheet.write(1, 6, "SALDO\n(c)=(a-b)", f["fmt_header"])
        worksheet.merge_range(0, 7, 0, 10, "DESPESA LIQUIDADA", f["fmt_header"])
        worksheet.write(1, 7, "NO BIMESTRE\n" + periodo, f["fmt_header"]); worksheet.write(1, 8, "ATE O BIMESTRE\n" + periodo + "\n(d)", f["fmt_header"]); worksheet.write(1, 9, "%\n(d/total d)", f["fmt_header"]); worksheet.write(1, 10, "SALDO\n(e)=(a-d)", f["fmt_header"])
        worksheet.merge_range(0, 11, 2, 11, "INSCRITAS EM\nRESTOS A\nPAGAR NAO\nPROCESSADOS (f)", f["fmt_header"])
        for c in range(3, 11):
            worksheet.write(2, c, "", f["fmt_header"])
        row = 3
        for desc, vals, tipo in linhas:
            fd = f["fmt_total_text"] if tipo == "total" else (f["fmt_group"] if tipo == "grupo" else f["fmt_item"])
            fn = f["fmt_money_total"] if tipo == "total" else f["fmt_money"]
            fp = f["fmt_pct_total"] if tipo == "total" else f["fmt_pct"]
            worksheet.write(row, 0, desc, fd)
            worksheet.write_number(row, 1, vals.get("orcado_inicial", 0), fn)
            worksheet.write_number(row, 2, vals.get("cred_autorizado", 0), fn)
            worksheet.write_number(row, 3, vals.get("emp_no_bim", 0), fn)
            worksheet.write_number(row, 4, vals.get("emp_ate", 0), fn)
            worksheet.write_number(row, 5, vals.get("perc_emp", 0), fp)
            worksheet.write_number(row, 6, vals.get("saldo_emp", 0), fn)
            worksheet.write_number(row, 7, vals.get("liq_no_bim", 0), fn)
            worksheet.write_number(row, 8, vals.get("liq_ate", 0), fn)
            worksheet.write_number(row, 9, vals.get("perc_liq", 0), fp)
            worksheet.write_number(row, 10, vals.get("saldo_liq", 0), fn)
            worksheet.write_number(row, 11, vals.get("restos", 0), fn)
            row += 1
        worksheet.freeze_panes(3, 1)
    return output.getvalue()


# ---------------------------------------------------------------------------
# CORES
# ---------------------------------------------------------------------------
COR_AZUL  = "#1B3A6B"
COR_VERDE = "#2E7D32"
COR_LARAN = "#E65100"
COR_CINZA = "#F5F7FA"

_RL_AZUL  = rl_c.HexColor("#1B3A6B")
_RL_VERDE = rl_c.HexColor("#2E7D32")
_RL_CINZA = rl_c.HexColor("#F0F2F5")


# ---------------------------------------------------------------------------
# HELPERS PARA GRÁFICOS MATPLOTLIB (usados nos PDFs)
# ---------------------------------------------------------------------------
def _fmt_val(v, _=None):
    if abs(v) >= 1e9:
        return f"R$ {v/1e9:.2f}Bi"
    if abs(v) >= 1e6:
        return f"R$ {v/1e6:.2f}M"
    if abs(v) >= 1e3:
        return f"R$ {v/1e3:.1f}k"
    return f"R$ {v:.0f}"


def _grafico_barras_bytes(labels, vals_a, vals_b, label_a="Orçado", label_b="Realizado"):
    fig, ax = plt.subplots(figsize=(10, 4))
    x = range(len(labels))
    w = 0.38
    b1 = ax.bar([i - w/2 for i in x], vals_a, w, label=label_a,
                color=COR_AZUL, alpha=0.85)
    b2 = ax.bar([i + w/2 for i in x], vals_b, w, label=label_b,
                color=COR_VERDE, alpha=0.85)
    ax.set_xticks(list(x))
    ax.set_xticklabels(labels, rotation=30, ha='right', fontsize=9)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(_fmt_val))
    ax.legend(fontsize=9)
    ax.grid(axis='y', alpha=0.3, linestyle='--')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    for bar in [*b1, *b2]:
        h = bar.get_height()
        if h > 0:
            ax.annotate(_fmt_val(h), xy=(bar.get_x() + bar.get_width() / 2, h),
                        xytext=(0, 2), textcoords="offset points",
                        ha='center', va='bottom', fontsize=6.5)
    plt.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    buf.seek(0)
    return buf


def _grafico_h_bytes(labels, valores, cor=COR_AZUL, pcts=None):
    n = max(len(labels), 1)
    fig, ax = plt.subplots(figsize=(10, max(3.5, n * 0.6)))
    y = list(range(n))
    ax.barh(y, valores, color=cor, alpha=0.85)
    ax.set_yticks(y)
    ax.set_yticklabels(labels, fontsize=8)
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(_fmt_val))
    ax.grid(axis='x', alpha=0.3, linestyle='--')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    for i, v in enumerate(valores):
        txt = _fmt_val(v)
        if pcts and i < len(pcts):
            txt += f"  ({pcts[i]:.1f}%)"
        ax.annotate(txt, xy=(v, i), xytext=(5, 0),
                    textcoords="offset points", va='center', fontsize=7.5)
    plt.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    buf.seek(0)
    return buf


# ---------------------------------------------------------------------------
# HELPERS PARA PDF (reportlab)
# ---------------------------------------------------------------------------
def _estilos_pdf():
    s = getSampleStyleSheet()
    st_titulo = ParagraphStyle('T', parent=s['Heading1'], fontSize=14,
                               textColor=_RL_AZUL, alignment=TA_CENTER,
                               spaceAfter=3, leading=17)
    st_sub    = ParagraphStyle('S', parent=s['Normal'],   fontSize=8.5,
                               textColor=rl_c.grey, alignment=TA_CENTER, spaceAfter=6)
    st_secao  = ParagraphStyle('SE', parent=s['Heading2'], fontSize=11,
                               textColor=_RL_AZUL, spaceBefore=10, spaceAfter=4)
    st_corpo  = ParagraphStyle('C', parent=s['Normal'],   fontSize=8, spaceAfter=2)
    return st_titulo, st_sub, st_secao, st_corpo


def _tabela_pdf(dados, col_w):
    """Cria tabela ReportLab com word-wrap automático em todas as células."""
    s = getSampleStyleSheet()
    _st_h = ParagraphStyle(
        "_th", parent=s["Normal"], fontSize=7.5, leading=10,
        textColor=rl_c.white, fontName="Helvetica-Bold", wordWrap="LTR"
    )
    _st_b = ParagraphStyle(
        "_tb", parent=s["Normal"], fontSize=7.5, leading=10, wordWrap="LTR"
    )
    _st_br = ParagraphStyle(
        "_tbr", parent=s["Normal"], fontSize=7.5, leading=10,
        wordWrap="LTR", alignment=2   # TA_RIGHT
    )
    dados_fmt = []
    for i, row in enumerate(dados):
        linha = []
        for j, cell in enumerate(row):
            if i == 0:
                linha.append(Paragraph(str(cell), _st_h))
            elif j >= 1:
                linha.append(Paragraph(str(cell), _st_br))
            else:
                linha.append(Paragraph(str(cell), _st_b))
        dados_fmt.append(linha)

    t = Table(dados_fmt, colWidths=col_w, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1,  0), _RL_AZUL),
        ("ROWBACKGROUNDS",(0, 1), (-1, -1), [rl_c.white, _RL_CINZA]),
        ("GRID",          (0, 0), (-1, -1), 0.35, rl_c.lightgrey),
        ("TOPPADDING",    (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING",   (0, 0), (-1, -1), 4),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 4),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
    ]))
    return t


def _header_pdf(elements, st1, st2, titulo, periodo, filtros):
    elements.append(Paragraph(
        "UO 03101 — TRIBUNAL DE JUSTIÇA DE MATO GROSSO", st1))
    elements.append(Paragraph(
        "Gestão Financeira Integrada — FIPLAN / 2026", st2))
    elements.append(HRFlowable(
        width="100%", thickness=2, color=_RL_AZUL, spaceAfter=3))
    elements.append(Paragraph(titulo, st1))
    elements.append(Spacer(1, 3))
    elements.append(Paragraph(
        f"<b>Período:</b> {periodo}   |   <b>Filtros:</b> {filtros}", st2))
    elements.append(HRFlowable(
        width="100%", thickness=0.5, color=rl_c.lightgrey, spaceAfter=6))


def gerar_pdf_receitas(df_rf, df_rec_total, mes_sel, filtros_str):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
          topMargin=1.3*cm, bottomMargin=1.3*cm,
          leftMargin=1.5*cm, rightMargin=1.5*cm)
    st1, st2, stS, stC = _estilos_pdf()
    el = []
    periodo = " / ".join(MESES_NOMES[m - 1] for m in sorted(mes_sel))
    _header_pdf(el, st1, st2, "RELATÓRIO DE RECEITAS", periodo,
                filtros_str or "Sem filtros adicionais")

    v_real     = float(df_rf["realizado"].sum())
    v_orc      = float(df_rf["orcado"].sum())
    v_real_tot = float(df_rec_total["realizado"].sum()) if not df_rec_total.empty else 0
    v_orc_tot  = float(df_rec_total["orcado"].sum())   if not df_rec_total.empty else 0
    pct_sel    = v_real / v_real_tot * 100 if v_real_tot  > 0 else 0
    pct_ating  = v_real / v_orc     * 100 if v_orc       > 0 else 0

    el.append(_tabela_pdf([
        ["Métrica", "Filtro Selecionado (R$)", "Total Banco (R$)", "% Filtro / Total"],
        ["Orçado",     f"{v_orc:,.2f}",  f"{v_orc_tot:,.2f}",
         f"{v_orc/v_orc_tot*100 if v_orc_tot>0 else 0:.1f}%"],
        ["Realizado",  f"{v_real:,.2f}", f"{v_real_tot:,.2f}", f"{pct_sel:.1f}%"],
        ["Atingimento",f"{pct_ating:.1f}%", "—", "—"],
    ], [4*cm, 5*cm, 5*cm, 3*cm]))
    el.append(Spacer(1, 10))

    df_g = df_rf.groupby("mes").agg({"realizado": "sum", "orcado": "sum"}).reset_index()
    if not df_g.empty:
        el.append(Paragraph("Orçado vs Realizado por Mês", stS))
        img = _grafico_barras_bytes(
            [MESES_NOMES[m - 1] for m in df_g["mes"]],
            list(df_g["orcado"]), list(df_g["realizado"]))
        el.append(RLImage(img, width=17*cm, height=7*cm))
        el.append(Spacer(1, 8))

    if "categoria" in df_rf.columns:
        df_cat = (df_rf.groupby("categoria")["realizado"]
                  .sum().reset_index().sort_values("realizado"))
        tot_c = float(df_cat["realizado"].sum())
        pcts  = [v / tot_c * 100 if tot_c > 0 else 0 for v in df_cat["realizado"]]
        el.append(Paragraph("Realizado por Categoria", stS))
        el.append(RLImage(
            _grafico_h_bytes(list(df_cat["categoria"]),
                             list(df_cat["realizado"]),
                             cor=COR_VERDE, pcts=pcts),
            width=17*cm, height=max(3.5*cm, len(df_cat) * 1.3*cm)))
        el.append(Spacer(1, 8))

    el.append(Paragraph("Detalhamento por Natureza", stS))
    df_tab = (df_rf.groupby(["categoria", "natureza"])
              .agg({"orcado": "sum", "realizado": "sum"}).reset_index())
    tot_t = float(df_tab["realizado"].sum())
    rows  = [["Categoria", "Natureza", "Orçado (R$)", "Realizado (R$)", "% s/Total"]]
    for _, r in df_tab.sort_values(["categoria", "realizado"], ascending=[True, False]).iterrows():
        pct = r["realizado"] / tot_t * 100 if tot_t > 0 else 0
        rows.append([r["categoria"], str(r["natureza"])[:45],
                     f"{r['orcado']:,.2f}", f"{r['realizado']:,.2f}", f"{pct:.1f}%"])
    if len(rows) > 1:
        el.append(_tabela_pdf(rows, [3.5*cm, 6*cm, 3.5*cm, 3.5*cm, 2*cm]))

    doc.build(el)
    buf.seek(0)
    return buf.getvalue()


def gerar_pdf_despesas(df_ef, cred_banco, mes_sel, filtros_str):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
          topMargin=1.3*cm, bottomMargin=1.3*cm,
          leftMargin=1.5*cm, rightMargin=1.5*cm)
    st1, st2, stS, stC = _estilos_pdf()
    el = []
    periodo = " / ".join(MESES_NOMES[m - 1] for m in sorted(mes_sel))
    _header_pdf(el, st1, st2, "RELATÓRIO DE DESPESAS", periodo,
                filtros_str or "Sem filtros adicionais")

    emp  = float(df_ef["empenhado"].sum()) if not df_ef.empty else 0
    liq  = float(df_ef["liquidado"].sum()) if not df_ef.empty else 0
    pago = float(df_ef["pago"].sum())      if not df_ef.empty else 0

    el.append(_tabela_pdf([
        ["Métrica", "Valor (R$)", "% Cred. Autorizado"],
        ["Crédito Autorizado", f"{cred_banco:,.2f}", "100,00%"],
        ["Empenhado",  f"{emp:,.2f}",  f"{emp /cred_banco*100 if cred_banco>0 else 0:.2f}%"],
        ["Liquidado",  f"{liq:,.2f}",  f"{liq /cred_banco*100 if cred_banco>0 else 0:.2f}%"],
        ["Pago",       f"{pago:,.2f}", f"{pago/cred_banco*100 if cred_banco>0 else 0:.2f}%"],
    ], [5*cm, 6*cm, 5*cm]))
    el.append(Spacer(1, 10))

    if not df_ef.empty:
        df_g = (df_ef.groupby("mes")[["empenhado", "liquidado"]]
                .sum().reset_index())
        el.append(Paragraph("Empenhado vs Liquidado por Mês", stS))
        el.append(RLImage(
            _grafico_barras_bytes(
                [MESES_NOMES[m - 1] for m in df_g["mes"]],
                list(df_g["empenhado"]), list(df_g["liquidado"]),
                "Empenhado", "Liquidado"),
            width=17*cm, height=7*cm))
        el.append(Spacer(1, 8))

        df_nat = (df_ef.groupby("natureza")[["empenhado", "liquidado"]]
                  .sum().reset_index()
                  .sort_values("liquidado", ascending=True).tail(15))
        tot_n = float(df_nat["liquidado"].sum())
        pcts_n = [v / tot_n * 100 if tot_n > 0 else 0 for v in df_nat["liquidado"]]
        el.append(Paragraph("Liquidado por Natureza (Top 15)", stS))
        el.append(RLImage(
            _grafico_h_bytes(list(df_nat["natureza"].astype(str)),
                             list(df_nat["liquidado"]), pcts=pcts_n),
            width=17*cm, height=max(4*cm, len(df_nat) * 1.1*cm)))
        el.append(Spacer(1, 8))

        el.append(Paragraph("Detalhamento por Natureza", stS))
        df_tab = (df_ef.groupby("natureza")[["empenhado", "liquidado", "pago"]]
                  .sum().reset_index())
        tot_t = float(df_tab["liquidado"].sum())
        rows  = [["Natureza", "Empenhado (R$)", "Liquidado (R$)", "Pago (R$)", "% Liq."]]
        for _, r in df_tab.sort_values("liquidado", ascending=False).iterrows():
            pct = r["liquidado"] / tot_t * 100 if tot_t > 0 else 0
            rows.append([str(r["natureza"])[:35],
                         f"{r['empenhado']:,.2f}", f"{r['liquidado']:,.2f}",
                         f"{r['pago']:,.2f}", f"{pct:.1f}%"])
        if len(rows) > 1:
            el.append(_tabela_pdf(rows, [4.5*cm, 3.5*cm, 3.5*cm, 3.5*cm, 2.5*cm]))

    doc.build(el)
    buf.seek(0)
    return buf.getvalue()


def _limpar_labels(series, max_len=40):
    """Remove acentos e trunca labels para evitar falhas no matplotlib."""
    resultado = []
    for v in series:
        s = sem_acento(str(v).replace("\xa0", " ").strip())
        resultado.append(s[:max_len])
    return resultado


def _h_bytes_seguro(labels, valores, cor=None, pcts=None):
    """Wrapper de _grafico_h_bytes com sanitização e cap de altura."""
    labels_limpos = _limpar_labels(labels)
    cor = cor or COR_AZUL
    try:
        return _grafico_h_bytes(labels_limpos, valores, cor=cor, pcts=pcts)
    except Exception:
        # Fallback: gráfico sem labels
        return _grafico_h_bytes(
            [str(i + 1) for i in range(len(valores))], valores, cor=cor)


# Altura máxima de gráfico que cabe em A4 com margens (cm)
_MAX_H_GRAF = 20 * cm


def gerar_pdf_701(df_sv, df_sub_total, mes_sel, filtros_str):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
          topMargin=1.3*cm, bottomMargin=1.3*cm,
          leftMargin=1.5*cm, rightMargin=1.5*cm)
    st1, st2, stS, stC = _estilos_pdf()
    el = []
    periodo = (" / ".join(MESES_NOMES[m - 1] for m in sorted(mes_sel))
               if mes_sel else "—")
    _header_pdf(el, st1, st2, "RELATÓRIO DE SUB-ELEMENTOS (FIP 701)", periodo,
                filtros_str or "Sem filtros adicionais")

    liq_sel  = float(df_sv["liquidado"].sum())        if not df_sv.empty        else 0
    pago_sel = float(df_sv["pago"].sum())             if not df_sv.empty        else 0
    liq_tot  = float(df_sub_total["liquidado"].sum()) if not df_sub_total.empty else 0
    pago_tot = float(df_sub_total["pago"].sum())      if not df_sub_total.empty else 0

    el.append(_tabela_pdf([
        ["Métrica", "Filtro (R$)", "Total Banco (R$)", "% Filtro / Total"],
        ["Liquidado", f"{liq_sel:,.2f}",  f"{liq_tot:,.2f}",
         f"{liq_sel / liq_tot * 100 if liq_tot > 0 else 0:.1f}%"],
        ["Pago",      f"{pago_sel:,.2f}", f"{pago_tot:,.2f}",
         f"{pago_sel / pago_tot * 100 if pago_tot > 0 else 0:.1f}%"],
    ], [4*cm, 5*cm, 5*cm, 3*cm]))
    el.append(Spacer(1, 10))

    if not df_sv.empty:
        # --- Gráfico por sub-elemento (Top 20) ---
        df_sub_g = (df_sv.groupby("subelemento_desc")[["liquidado"]]
                    .sum().reset_index()
                    .sort_values("liquidado", ascending=True).tail(20))
        tot_s  = float(df_sub_g["liquidado"].sum())
        pcts_s = [v / tot_s * 100 if tot_s > 0 else 0 for v in df_sub_g["liquidado"]]
        h_sub  = min(_MAX_H_GRAF, max(5*cm, len(df_sub_g) * 1.1*cm))
        el.append(Paragraph(f"Liquidado por Sub-elemento (Top {len(df_sub_g)})", stS))
        el.append(RLImage(
            _h_bytes_seguro(list(df_sub_g["subelemento_desc"]), list(df_sub_g["liquidado"]),
                            pcts=pcts_s),
            width=17*cm, height=h_sub))
        el.append(Spacer(1, 8))

        # --- Gráfico por natureza (Top 15) ---
        df_nat = (df_sv.groupby(["natureza_cod", "natureza_desc"])[["liquidado"]]
                  .sum().reset_index()
                  .sort_values("liquidado", ascending=True).tail(15))
        tot_n  = float(df_nat["liquidado"].sum())
        pcts_n = [v / tot_n * 100 if tot_n > 0 else 0 for v in df_nat["liquidado"]]
        h_nat  = min(_MAX_H_GRAF, max(3.5*cm, len(df_nat) * 1.3*cm))
        el.append(Paragraph(f"Liquidado por Natureza de Despesa (Top {len(df_nat)})", stS))
        el.append(RLImage(
            _h_bytes_seguro(list(df_nat["natureza_desc"]), list(df_nat["liquidado"]),
                            cor=COR_VERDE, pcts=pcts_n),
            width=17*cm, height=h_nat))
        el.append(Spacer(1, 8))

        # --- Tabela de detalhamento ---
        el.append(Paragraph("Detalhamento Completo", stS))
        df_tab = (df_sv.groupby(
            ["subelemento_cod", "subelemento_desc", "natureza_cod", "natureza_desc"])
            [["liquidado", "pago"]].sum().reset_index())
        tot_t = float(df_tab["liquidado"].sum())
        rows  = [["Sub-elemento", "Natureza", "Liquidado (R$)", "Pago (R$)", "% s/Total"]]
        for _, r in df_tab.sort_values("liquidado", ascending=False).iterrows():
            pct = r["liquidado"] / tot_t * 100 if tot_t > 0 else 0
            rows.append([
                sem_acento(str(r["subelemento_desc"])),
                sem_acento(str(r["natureza_desc"])),
                f"{r['liquidado']:,.2f}", f"{r['pago']:,.2f}", f"{pct:.1f}%"
            ])
        if len(rows) > 1:
            el.append(_tabela_pdf(rows, [5.5*cm, 5*cm, 3*cm, 2.5*cm, 1.5*cm]))

    doc.build(el)
    buf.seek(0)
    return buf.getvalue()


def gerar_pdf_comparativo(tr, te, tl, tp, v_orc, mes_sel, df_rec, df_exec):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
          topMargin=1.3*cm, bottomMargin=1.3*cm,
          leftMargin=1.5*cm, rightMargin=1.5*cm)
    st1, st2, stS, stC = _estilos_pdf()
    el = []
    periodo = (" / ".join(MESES_NOMES[m - 1] for m in sorted(mes_sel))
               if mes_sel else "—")
    _header_pdf(el, st1, st2, "RELATÓRIO COMPARATIVO — RECEITA × DESPESA",
                periodo, "Período selecionado")

    sup_fin  = tr - tp
    sup_orc  = tr - te
    pct_emp  = te / tr * 100 if tr > 0 else 0
    pct_liq  = tl / tr * 100 if tr > 0 else 0
    pct_pag  = tp / tr * 100 if tr > 0 else 0
    pct_cred = te / v_orc * 100 if v_orc > 0 else 0

    el.append(_tabela_pdf([
        ["Indicador", "Valor (R$)", "% s/ Receita Arrecadada"],
        ["Receita Arrecadada",  f"{tr:,.2f}", "100,00%"],
        ["Cred. Autorizado",    f"{v_orc:,.2f}", "—"],
        ["Desp. Empenhada",     f"{te:,.2f}", f"{pct_emp:.2f}%"],
        ["Desp. Liquidada",     f"{tl:,.2f}", f"{pct_liq:.2f}%"],
        ["Desp. Paga",          f"{tp:,.2f}", f"{pct_pag:.2f}%"],
        ["Superávit Financeiro (Rec. - Pago)",     f"{sup_fin:,.2f}", "—"],
        ["Superávit Orçamentário (Rec. - Emp.)",   f"{sup_orc:,.2f}", "—"],
    ], [7*cm, 5.5*cm, 5*cm]))
    el.append(Spacer(1, 12))

    # Gráfico mensal receita vs despesa
    if not df_rec.empty and not df_exec.empty and mes_sel:
        meses_comuns = sorted(set(mes_sel))
        rec_m = (df_rec[df_rec["mes"].isin(meses_comuns)]
                 .groupby("mes")["realizado"].sum().reindex(meses_comuns, fill_value=0))
        emp_m = (df_exec[df_exec["mes"].isin(meses_comuns)]
                 .groupby("mes")["empenhado"].sum().reindex(meses_comuns, fill_value=0))
        liq_m = (df_exec[df_exec["mes"].isin(meses_comuns)]
                 .groupby("mes")["liquidado"].sum().reindex(meses_comuns, fill_value=0))
        labels_m = [MESES_NOMES[m - 1] for m in meses_comuns]
        el.append(Paragraph("Receita vs Despesas por Mês", stS))
        el.append(RLImage(
            _grafico_barras_bytes(labels_m, list(rec_m), list(emp_m),
                                  "Receita Arrecadada", "Desp. Empenhada"),
            width=17*cm, height=7*cm))
        el.append(Spacer(1, 8))

        # Liquidado vs Pago
        el.append(Paragraph("Desp. Liquidada vs Paga por Mês", stS))
        el.append(RLImage(
            _grafico_barras_bytes(labels_m, list(liq_m),
                                  list(df_exec[df_exec["mes"].isin(meses_comuns)]
                                       .groupby("mes")["pago"].sum()
                                       .reindex(meses_comuns, fill_value=0)),
                                  "Liquidado", "Pago"),
            width=17*cm, height=7*cm))

    doc.build(el)
    buf.seek(0)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# RESTOS A PAGAR — PDF
# ---------------------------------------------------------------------------
def gerar_pdf_rp(df_rp_ref, mes_ref, meses_selecionados):
    """Gera PDF de Restos a Pagar com base na FIP 226 do mês de referência."""
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
          topMargin=1.3*cm, bottomMargin=1.3*cm,
          leftMargin=1.5*cm, rightMargin=1.5*cm)
    st1, st2, stS, stC = _estilos_pdf()
    el = []
    periodo = " / ".join(MESES_NOMES[m - 1] for m in sorted(meses_selecionados))
    _header_pdf(el, st1, st2, "RELATÓRIO DE RESTOS A PAGAR", periodo,
                "FIP 226 — Referência: " + MESES_NOMES[mes_ref - 1])

    proc_ins = float(df_rp_ref["proc_inscrito"].sum())
    proc_pag = float(df_rp_ref["proc_pagos"].sum())
    proc_can = float(df_rp_ref["proc_cancelados"].sum())
    proc_apa = float(df_rp_ref["proc_a_pagar"].sum())
    np_ins   = float(df_rp_ref["np_inscrito"].sum())
    np_pag   = float(df_rp_ref["np_pagos"].sum())
    np_can   = float(df_rp_ref["np_cancelados"].sum())
    np_aliq  = float(df_rp_ref["np_a_liquidar"].sum())

    tot_ins  = proc_ins + np_ins
    tot_pag  = proc_pag + np_pag
    tot_can  = proc_can + np_can
    tot_aliq = proc_apa + np_aliq
    tot_liq  = max(0.0, tot_ins - tot_aliq - tot_can)

    el.append(_tabela_pdf([
        ["Tipo", "Inscrito (R$)", "Pagos (R$)", "Cancelados (R$)", "A Liquidar (R$)", "Liquidados (R$)"],
        ["Processados",
         f"{proc_ins:,.2f}", f"{proc_pag:,.2f}", f"{proc_can:,.2f}",
         f"{proc_apa:,.2f}", f"{max(0.0, proc_ins - proc_apa - proc_can):,.2f}"],
        ["Não Processados",
         f"{np_ins:,.2f}", f"{np_pag:,.2f}", f"{np_can:,.2f}",
         f"{np_aliq:,.2f}", f"{max(0.0, np_ins - np_aliq - np_can):,.2f}"],
        ["TOTAL",
         f"{tot_ins:,.2f}", f"{tot_pag:,.2f}", f"{tot_can:,.2f}",
         f"{tot_aliq:,.2f}", f"{tot_liq:,.2f}"],
    ], [4*cm, 3*cm, 3*cm, 3*cm, 3*cm, 3*cm]))
    el.append(Spacer(1, 10))

    # Gráfico por tipo
    el.append(Paragraph("Visão por Tipo de Resto a Pagar", stS))
    labels_tipo = ["Processados", "Nao Processados"]
    vals_ins  = [proc_ins, np_ins]
    vals_pag  = [proc_pag, np_pag]
    el.append(RLImage(
        _grafico_barras_bytes(labels_tipo, vals_ins, vals_pag, "Inscrito", "Pago"),
        width=14*cm, height=6*cm))
    el.append(Spacer(1, 8))

    # Top credores por inscrito
    df_cred = (df_rp_ref.assign(inscrito_total=df_rp_ref["proc_inscrito"] + df_rp_ref["np_inscrito"])
               .groupby("nome_credor", as_index=False)["inscrito_total"].sum()
               .sort_values("inscrito_total", ascending=False).head(15))
    if not df_cred.empty:
        el.append(Paragraph("Top Credores — Total Inscrito", stS))
        labels_c = _limpar_labels(df_cred["nome_credor"])
        h_buf = _h_bytes_seguro(labels_c, list(df_cred["inscrito_total"]), cor=COR_AZUL)
        h_img = min(_MAX_H_GRAF, max(3.5*cm, len(df_cred) * 1.3*cm))
        el.append(RLImage(h_buf, width=17*cm, height=h_img))
        el.append(Spacer(1, 8))

    # Tabela por empenho (top 30)
    df_det = (df_rp_ref[["num_empenho", "nome_credor", "proc_inscrito", "np_inscrito",
                          "proc_pagos", "np_pagos", "proc_cancelados", "np_cancelados",
                          "np_a_liquidar", "proc_a_pagar"]]
              .copy())
    df_det["inscrito"] = df_det["proc_inscrito"] + df_det["np_inscrito"]
    df_det = df_det.sort_values("inscrito", ascending=False).head(30)
    if not df_det.empty:
        el.append(Paragraph("Detalhamento por Empenho (Top 30 por Inscrito)", stS))
        rows_tab = [["Empenho", "Credor", "Inscrito", "Pagos", "Cancelados", "A Liquidar"]]
        for _, r in df_det.iterrows():
            rows_tab.append([
                str(r["num_empenho"]),
                str(r["nome_credor"])[:40],
                f"{r['inscrito']:,.2f}",
                f"{r['proc_pagos'] + r['np_pagos']:,.2f}",
                f"{r['proc_cancelados'] + r['np_cancelados']:,.2f}",
                f"{r['proc_a_pagar'] + r['np_a_liquidar']:,.2f}",
            ])
        el.append(_tabela_pdf(rows_tab, [4*cm, 5.5*cm, 2.5*cm, 2.5*cm, 2.5*cm, 2.5*cm]))

    doc.build(el)
    buf.seek(0)
    return buf.getvalue()


def gerar_excel_rp(df_rp, meses_bim, meses_ate_agora):
    """Gera Anexo RP no formato LRF Anexo VII por fonte, com split exerc.ant./(b)."""
    def extrair_fonte(dotacao):
        try:
            partes = str(dotacao).split(".")
            # UG.FONTE = partes[-3] + "." + partes[-2] → ex: "1.760" e "2.760"
            return (partes[-3] + "." + partes[-2]).strip() if len(partes) >= 3 else "N/D"
        except Exception:
            return "N/D"

    def extrair_ano_emp(empenho):
        try:
            # Empenho: "03601.0001.23.006821-7" → 3º segmento = ano (2 dígitos)
            return int(str(empenho).split(".")[2])
        except Exception:
            return 0

    mes_ref = max(m for m in df_rp["mes"].unique() if m in meses_ate_agora) if not df_rp.empty else None
    if mes_ref is None:
        return io.BytesIO().getvalue()

    # Ano do exercício anterior em 2 dígitos (ex: 2026 → 25)
    ano_ref = int(df_rp["ano"].max()) if "ano" in df_rp.columns and not df_rp["ano"].isnull().all() else 2026
    ano_ant_2d = (ano_ref - 1) % 100

    df_ref = df_rp[df_rp["mes"] == mes_ref].copy()
    df_ref["fonte"]   = df_ref["dotacao"].apply(extrair_fonte)
    df_ref["ano_emp"] = df_ref["num_empenho"].apply(extrair_ano_emp)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        ws = workbook.add_worksheet("RP_LRF")
        writer.sheets["RP_LRF"] = ws
        f = criar_formatos_excel(workbook)

        ws.set_column(0, 0, 34)
        ws.set_column(1, 12, 15)

        # ---- Linha 0: super-cabeçalhos ----------------------------------
        ws.merge_range(0, 0, 2, 0, "PODER/ÓRGÃO", f["fmt_header"])
        ws.merge_range(0, 1, 0, 5, "RP PROCESSADOS", f["fmt_header"])
        ws.merge_range(0, 6, 0, 11, "RP NÃO-PROCESSADOS", f["fmt_header"])
        ws.merge_range(0, 12, 2, 12, "SALDO TOTAL\nL=(e+k)", f["fmt_header"])

        # ---- Linha 1: sub-grupos ----------------------------------------
        ws.merge_range(1, 1, 1, 2, "INSCRITOS", f["fmt_header"])
        ws.merge_range(1, 3, 2, 3, "PAGOS\n(c)", f["fmt_header"])
        ws.merge_range(1, 4, 2, 4, "CANCELADOS\n(d)", f["fmt_header"])
        ws.merge_range(1, 5, 2, 5, "SALDO\ne=(a+b)-(c+d)", f["fmt_header"])
        ws.merge_range(1, 6, 1, 7, "INSCRITOS", f["fmt_header"])
        ws.merge_range(1, 8, 2, 8, "LIQUIDADOS\n(h)", f["fmt_header"])
        ws.merge_range(1, 9, 2, 9, "PAGOS\n(i)", f["fmt_header"])
        ws.merge_range(1, 10, 2, 10, "CANCELADOS\n(j)", f["fmt_header"])
        ws.merge_range(1, 11, 2, 11, "SALDO\nk=(f+g)-(i+j)", f["fmt_header"])

        # ---- Linha 2: sub-colunas de inscritos --------------------------
        ws.write(2, 1, "EM EXERC.\nANTERIORES\n(a)", f["fmt_header"])
        ws.write(2, 2, "EM 31/12\nEXERC.ANT.\n(b)", f["fmt_header"])
        ws.write(2, 6, "EM EXERC.\nANTERIORES\n(f)", f["fmt_header"])
        ws.write(2, 7, "EM 31/12\nEXERC.ANT.\n(g)", f["fmt_header"])

        def soma(df, col):
            return float(df[col].sum()) if not df.empty and col in df.columns else 0.0

        def calcular_linha(df_fonte):
            # Split inscrito por ano do empenho
            df_ant   = df_fonte[df_fonte["ano_emp"] < ano_ant_2d]   # exerc. anteriores
            df_31dez = df_fonte[df_fonte["ano_emp"] >= ano_ant_2d]  # 31/12 exerc.ant.
            proc_a = soma(df_ant,   "proc_inscrito")   # (a)
            proc_b = soma(df_31dez, "proc_inscrito")   # (b)
            np_f   = soma(df_ant,   "np_inscrito")     # (f)
            np_g   = soma(df_31dez, "np_inscrito")     # (g)
            # Pagos, cancelados e saldos — totais do grupo
            pp  = soma(df_fonte, "proc_pagos")
            pc  = soma(df_fonte, "proc_cancelados")
            pa  = soma(df_fonte, "proc_a_pagar")
            np_ = soma(df_fonte, "np_pagos")
            nc  = soma(df_fonte, "np_cancelados")
            na  = soma(df_fonte, "np_a_liquidar")
            ni  = np_f + np_g
            saldo_proc = pa                         # e = (a+b) - c - d
            liq_np     = max(0.0, ni - na - nc)     # h: NP liquidados
            saldo_np   = max(0.0, ni - np_ - nc)    # k = (f+g) - i - j
            return (proc_a, proc_b, pp, pc, saldo_proc,
                    np_f, np_g, liq_np, np_, nc, saldo_np, saldo_proc + saldo_np)

        fontes = sorted(df_ref["fonte"].unique())
        row = 3
        tot = [0.0] * 12

        for fonte in fontes:
            df_f = df_ref[df_ref["fonte"] == fonte]
            vals = calcular_linha(df_f)
            ws.write(row, 0, "JUDICIÁRIO/TJMT/" + fonte, f["fmt_group"])
            for c, v in enumerate(vals, 1):
                ws.write_number(row, c, v, f["fmt_money"])
                tot[c - 1] += v
            row += 1

        # ---- TOTAL GERAL -----------------------------------------------
        ws.write(row, 0, "TOTAL GERAL", f["fmt_total_text"])
        for c, v in enumerate(tot, 1):
            ws.write_number(row, c, v, f["fmt_money_total"])

        ws.freeze_panes(3, 1)
    return output.getvalue()


# ---------------------------------------------------------------------------
# BANCO DE DADOS
# ---------------------------------------------------------------------------
def inicializar_banco():
    conn = sqlite3.connect(DB_NAME)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA busy_timeout=5000")

    conn.execute(
        "CREATE TABLE IF NOT EXISTS receitas ("
        "mes INTEGER, ano INTEGER, codigo_full TEXT, natureza TEXT, "
        "orcado REAL, realizado REAL, previsao REAL, "
        "categoria TEXT DEFAULT 'Nao Classificada')"
    )
    conn.execute(
        "CREATE TABLE IF NOT EXISTS orcamento ("
        "mes INTEGER, ano INTEGER, uo TEXT, ug TEXT, funcao TEXT, subfuncao TEXT, "
        "programa TEXT, projeto TEXT, natureza TEXT, fonte TEXT, "
        "orcado_inicial REAL, cred_autorizado REAL)"
    )
    conn.execute(
        "CREATE TABLE IF NOT EXISTS execucao ("
        "mes INTEGER, ano INTEGER, uo TEXT, ug TEXT, funcao TEXT, subfuncao TEXT, "
        "programa TEXT, projeto TEXT, regional TEXT, natureza TEXT, fonte TEXT, "
        "iduso TEXT, tipo_rec TEXT, "
        "empenhado REAL, liquidado REAL, pago REAL)"
    )
    conn.execute(
        "CREATE TABLE IF NOT EXISTS sub_elementos ("
        "mes INTEGER, ano INTEGER, ug TEXT, paoe TEXT, natureza_cod TEXT, natureza_desc TEXT, "
        "subelemento_cod TEXT, subelemento_desc TEXT, fonte TEXT, "
        "liquidado REAL, pago REAL)"
    )
    conn.execute(
        "CREATE TABLE IF NOT EXISTS anexo_v ("
        "mes INTEGER, ano INTEGER, data TEXT, "
        "entidade_repassadora TEXT, valor REAL, "
        "finalidade TEXT, fundamento_legal TEXT)"
    )
    conn.execute(
        "CREATE TABLE IF NOT EXISTS restos_a_pagar ("
        "mes INTEGER, ano INTEGER, "
        "cod_credor TEXT, nome_credor TEXT, num_empenho TEXT, data_empenho TEXT, "
        "dotacao TEXT, "
        "proc_inscrito REAL, proc_pagos REAL, proc_cancelados REAL, proc_a_pagar REAL, "
        "np_inscrito REAL, np_pagos REAL, np_cancelados REAL, np_a_pagar REAL, "
        "np_em_liquidacao REAL, np_a_liquidar REAL)"
    )
    conn.commit()

    for stmt in [
        "ALTER TABLE receitas ADD COLUMN categoria TEXT DEFAULT 'Nao Classificada'",
        "ALTER TABLE sub_elementos ADD COLUMN ug TEXT DEFAULT ''",
    ]:
        try:
            conn.execute(stmt)
            conn.commit()
        except Exception:
            conn.rollback()

    try:
        conn.execute(
            "UPDATE receitas SET categoria='Receita Corrente' "
            "WHERE categoria='Repasses Correntes'"
        )
        conn.commit()
    except Exception:
        conn.rollback()

    conn.close()


def limpar_todos_dados():
    conn = sqlite3.connect(DB_NAME)
    conn.execute("DELETE FROM receitas")
    conn.execute("DELETE FROM orcamento")
    conn.execute("DELETE FROM execucao")
    conn.execute("DELETE FROM sub_elementos")
    conn.execute("DELETE FROM anexo_v")
    conn.execute("DELETE FROM restos_a_pagar")
    try:
        conn.execute("DELETE FROM despesas")
    except Exception:
        pass
    conn.commit()
    conn.close()


inicializar_banco()


# ---------------------------------------------------------------------------
# SIDEBAR
# ---------------------------------------------------------------------------
with st.sidebar:
    st.subheader("Importar Dados")
    tipo_dado = st.radio(
        "Tipo:", [
            "Receita (FIP 729)",
            "Orcamento (FIP 616)",
            "Execucao (FIP 613)",
            "Sub-elemento (FIP 701)",
            "Repasses Recebidos (ANEXO V)",
            "Restos a Pagar (FIP 226)"
        ]
    )
    arquivo = st.file_uploader("Arquivo Excel", type=["xlsx"])

    if arquivo and st.button("Processar Dados"):
        try:
            m_final = detectar_mes(arquivo)
            conn = sqlite3.connect(DB_NAME)

            # ----------------------------------------------------------------
            # RECEITA (FIP 729)
            # ----------------------------------------------------------------
            if tipo_dado == "Receita (FIP 729)":
                df = pd.read_excel(arquivo, skiprows=7, header=0)
                dados = []
                for _, row in df.iterrows():
                    try:
                        cod = str(row.iloc[0]).strip().replace('"', "")
                        if not re.match(r"^\d", cod) or cod.endswith(".00"):
                            continue
                        real = limpar_f(row.iloc[6])
                        if cod.startswith("9"):
                            real = -abs(real)
                        cur = conn.execute(
                            "SELECT categoria FROM receitas WHERE codigo_full=?", (cod,)
                        )
                        r_cat = cur.fetchone()
                        cat = r_cat[0] if r_cat else "Nao Classificada"
                        dados.append((
                            m_final, 2026, cod,
                            str(row.iloc[1]).replace('"', ""),
                            limpar_f(row.iloc[3]), real, limpar_f(row.iloc[5]), cat
                        ))
                    except Exception:
                        continue
                conn.execute(
                    "DELETE FROM receitas WHERE ano=2026 AND mes=?", (m_final,)
                )
                conn.executemany(
                    "INSERT INTO receitas (mes, ano, codigo_full, natureza, orcado, realizado, previsao, categoria) VALUES (?,?,?,?,?,?,?,?)", dados
                )
                conn.commit()
                st.success(
                    "Receita " + MESES_NOMES[m_final - 1]
                    + "/2026: " + str(len(dados)) + " registros"
                )

            # ----------------------------------------------------------------
            # ORCAMENTO (FIP 616)
            # ----------------------------------------------------------------
            elif tipo_dado == "Orcamento (FIP 616)":
                df = pd.read_excel(arquivo, skiprows=6, header=0)
                n = len(df.columns)

                def gc616(row, i, default=0):
                    return row.iloc[i] if i < n else default

                linhas = []
                for _, row in df.iterrows():
                    try:
                        uo = norm(gc616(row, 1))
                        if not uo or uo in ("nan", "") or uo != "3101":
                            continue
                        ug        = norm(gc616(row, 2))
                        funcao    = norm(gc616(row, 3))
                        subfuncao = norm(gc616(row, 4))
                        programa  = norm(gc616(row, 5))
                        projeto   = norm(gc616(row, 6))
                        natureza  = norm(gc616(row, 7))
                        fonte     = norm(gc616(row, 10))
                        orc_ini   = limpar_f(gc616(row, 11, 0))
                        cred_aut  = limpar_f(gc616(row, 12, 0))
                        if orc_ini == 0 and cred_aut == 0:
                            continue
                        linhas.append((
                            m_final, 2026, uo, ug, funcao, subfuncao,
                            programa, projeto, natureza, fonte,
                            orc_ini, cred_aut
                        ))
                    except Exception:
                        continue

                conn.execute("DELETE FROM orcamento WHERE ano=2026")
                conn.executemany(
                    "INSERT INTO orcamento VALUES (?,?,?,?,?,?,?,?,?,?,?,?)", linhas
                )
                conn.commit()
                cred_total = sum(r[11] for r in linhas)
                st.success(
                    "Orcamento " + MESES_NOMES[m_final - 1] + "/2026: "
                    + str(len(linhas)) + " linhas | "
                    + "Cred. Autorizado R$ {:,.0f}".format(cred_total)
                )

            # ----------------------------------------------------------------
            # EXECUCAO (FIP 613)
            # ----------------------------------------------------------------
            elif tipo_dado == "Execucao (FIP 613)":
                df = pd.read_excel(arquivo, skiprows=10, header=0)
                n = len(df.columns)

                def gc613(row, i, default=0):
                    return row.iloc[i] if i < n else default

                linhas = []
                for _, row in df.iterrows():
                    try:
                        uo = norm(gc613(row, 0))
                        if not uo or uo in ("nan", "", "_"):
                            continue
                        if pd.isna(gc613(row, 9, float("nan"))):
                            continue
                        ug        = norm(gc613(row, 1))
                        funcao    = norm(gc613(row, 2))
                        subfuncao = norm(gc613(row, 3))
                        programa  = norm(gc613(row, 4))
                        projeto   = norm(gc613(row, 5))
                        regional  = norm(gc613(row, 6))
                        natureza  = norm(gc613(row, 7))
                        fonte     = norm(gc613(row, 8))
                        iduso     = norm(gc613(row, 9))
                        tipo_rec  = norm(gc613(row, 10))
                        emp = limpar_f(gc613(row, 21, 0))
                        liq = limpar_f(gc613(row, 22, 0))
                        pag = limpar_f(gc613(row, 24, 0))
                        if emp == 0 and liq == 0 and pag == 0:
                            continue
                        linhas.append((
                            m_final, 2026, uo, ug, funcao, subfuncao,
                            programa, projeto, regional, natureza, fonte,
                            iduso, tipo_rec, emp, liq, pag
                        ))
                    except Exception:
                        continue

                conn.execute(
                    "DELETE FROM execucao WHERE ano=2026 AND mes=?", (m_final,)
                )
                conn.executemany(
                    "INSERT INTO execucao VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                    linhas
                )
                conn.commit()
                ugs = sorted(set(r[3] for r in linhas))
                emp_t = sum(r[13] for r in linhas)
                liq_t = sum(r[14] for r in linhas)
                pag_t = sum(r[15] for r in linhas)
                st.success(
                    "Execucao " + MESES_NOMES[m_final - 1] + "/2026: "
                    + str(len(linhas)) + " linhas | "
                    + "Emp R$ {:,.0f} | Liq R$ {:,.0f} | Pago R$ {:,.0f}".format(
                        emp_t, liq_t, pag_t
                    )
                )
                st.info("UGs encontradas: " + str(ugs))

            # ----------------------------------------------------------------
            # SUB-ELEMENTO (FIP 701)
            # ----------------------------------------------------------------
            elif tipo_dado == "Sub-elemento (FIP 701)":
                df701 = pd.read_excel(arquivo, header=None)
                linhas = []
                cur_ug = ""
                cur_paoe = ""
                cur_nat_cod = ""
                cur_nat_desc = ""

                for i, row in df701.iterrows():
                    text = str(row.iloc[0]).strip().replace("\xa0", " ")
                    if i < 8 or not text or text == "nan":
                        continue
                    tu = sem_acento(text).upper()

                    if re.match(r"^UG\s+\d+", tu):
                        m = re.search(r"(\d+)", text)
                        if m:
                            cur_ug = m.group(1).strip()
                        continue

                    if "PROJ/ATIV" in tu and ":" in tu:
                        m = re.search(r"(\d{5,8})", text)
                        if m:
                            cur_paoe = m.group(1)
                        continue

                    if "NATUREZA" in tu and "DESPESA" in tu and ":" in tu:
                        m = re.search(r":\s*(\d+)\s*-\s*(.*)", text)
                        if m:
                            cur_nat_cod = m.group(1).strip()
                            raw = m.group(2).replace("\xa0", " ").strip()
                            cur_nat_desc = (
                                raw.split(" - ")[0].strip()
                                if " - " in raw else raw
                            )
                        continue

                    if (tu.startswith("TOTAL") or tu.startswith("CONSOLID")
                            or tu.startswith("DOTA")):
                        continue

                    if re.match(r"^\d+\.\d+", text) and cur_paoe and cur_nat_cod:
                        parts = text.split(" ", 1)
                        sub_cod  = parts[0].strip()
                        sub_desc = parts[1].strip() if len(parts) > 1 else ""
                        fonte_sub = sub_cod.rsplit(".", 1)[-1] if "." in sub_cod else ""
                        liq_cum  = (
                            limpar_f(row.iloc[1]) if pd.notna(row.iloc[1]) else 0.0
                        )
                        pag_cum  = (
                            limpar_f(row.iloc[2]) if pd.notna(row.iloc[2]) else 0.0
                        )
                        linhas.append({
                            "ug": cur_ug,
                            "paoe": cur_paoe,
                            "nat_cod": cur_nat_cod,
                            "nat_desc": cur_nat_desc,
                            "sub_cod": sub_cod,
                            "sub_desc": sub_desc,
                            "fonte": fonte_sub,
                            "liq_cum": liq_cum,
                            "pag_cum": pag_cum,
                        })

                if not linhas:
                    st.warning("Nenhum sub-elemento valido encontrado.")
                else:
                    chaves_701 = ["ug", "paoe", "nat_cod", "sub_cod"]
                    df_mes = (
                        pd.DataFrame(linhas)
                        .groupby(
                            chaves_701 + ["nat_desc", "sub_desc", "fonte"],
                            as_index=False
                        )
                        .agg(
                            liq_cum=("liq_cum", "sum"),
                            pag_cum=("pag_cum", "sum")
                        )
                    )
                    dados = [
                        (
                            m_final, 2026,
                            r.ug, r.paoe, r.nat_cod, r.nat_desc,
                            r.sub_cod, r.sub_desc, r.fonte,
                            float(r.liq_cum), float(r.pag_cum),
                        )
                        for r in df_mes.itertuples()
                    ]
                    conn.execute(
                        "DELETE FROM sub_elementos WHERE ano=2026 AND mes=?",
                        (m_final,)
                    )
                    conn.executemany(
                        "INSERT INTO sub_elementos "
                        "(mes, ano, ug, paoe, natureza_cod, natureza_desc, "
                        "subelemento_cod, subelemento_desc, fonte, liquidado, pago) "
                        "VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                        dados
                    )
                    conn.commit()
                    liq_t = sum(r[9] for r in dados)
                    pag_t = sum(r[10] for r in dados)
                    ugs_imp = sorted(set(r[2] for r in dados if r[2]))
                    st.success(
                        "Sub-elemento " + MESES_NOMES[m_final - 1] + "/2026: "
                        + str(len(dados)) + " registros | "
                        + "Liq R$ {:,.0f} | Pago R$ {:,.0f}".format(liq_t, pag_t)
                    )
                    if ugs_imp:
                        st.info("UGs importadas: " + str(ugs_imp))

            # ----------------------------------------------------------------
            # REPASSES RECEBIDOS (ANEXO V)
            # Estrutura: header na linha 9 (0-indexed), dados a partir da linha 10
            # Colunas: Data | Entidade Repassadora | Valor | Finalidade | Fundamento Legal
            # ----------------------------------------------------------------
            elif tipo_dado == "Repasses Recebidos (ANEXO V)":
                df_av = pd.read_excel(arquivo, skiprows=9, header=0)
                dados_av = []
                for _, row in df_av.iterrows():
                    try:
                        data_val = str(row.iloc[0]).strip()
                        # Para quando chega na linha de total ou vazia
                        if not data_val or data_val in ("nan", "") or "Total" in data_val:
                            continue
                        entidade = str(row.iloc[1]).strip().replace("\xa0", " ")
                        if not entidade or entidade in ("nan", "Entidade Repassadora"):
                            continue
                        valor = limpar_f(row.iloc[2])
                        finalidade = str(row.iloc[3]).strip().replace("\xa0", " ") if pd.notna(row.iloc[3]) else ""
                        fundamento = str(row.iloc[4]).strip() if pd.notna(row.iloc[4]) else ""
                        dados_av.append((
                            m_final, 2026, data_val, entidade, valor, finalidade, fundamento
                        ))
                    except Exception:
                        continue

                conn.execute(
                    "DELETE FROM anexo_v WHERE ano=2026 AND mes=?", (m_final,)
                )
                conn.executemany(
                    "INSERT INTO anexo_v VALUES (?,?,?,?,?,?,?)", dados_av
                )
                conn.commit()
                total_av = sum(r[4] for r in dados_av)
                st.success(
                    "ANEXO V " + MESES_NOMES[m_final - 1] + "/2026: "
                    + str(len(dados_av)) + " registros | "
                    + "Total R$ {:,.2f}".format(total_av)
                )

            # ----------------------------------------------------------------
            # RESTOS A PAGAR (FIP 226)
            # Estrutura: 7 linhas de cabeçalho (0-5 metadata, 6 = nomes colunas)
            # 15 colunas: credor, nome, empenho, data, dotacao,
            #             proc_ins, proc_pag, proc_can, proc_apa,
            #             np_ins, np_pag, np_can, np_apa, np_eliq, np_aliq
            # ----------------------------------------------------------------
            elif tipo_dado == "Restos a Pagar (FIP 226)":
                df226 = pd.read_excel(arquivo, skiprows=6, header=0)
                dados_rp = []
                for _, row in df226.iterrows():
                    try:
                        cod_cred = str(row.iloc[0]).strip()
                        if (not cod_cred or cod_cred in ("nan", "", "_")
                                or cod_cred.upper().startswith("TOTAL")
                                or "CREDOR" in cod_cred.upper()):
                            continue
                        nome_cred = str(row.iloc[1]).strip().replace("\xa0", " ")
                        num_emp   = str(row.iloc[2]).strip()
                        data_emp  = str(row.iloc[3]).strip()
                        dotacao   = str(row.iloc[4]).strip()
                        proc_ins  = limpar_f(row.iloc[5])
                        proc_pag  = limpar_f(row.iloc[6])
                        proc_can  = limpar_f(row.iloc[7])
                        proc_apa  = limpar_f(row.iloc[8])
                        np_ins    = limpar_f(row.iloc[9])
                        np_pag    = limpar_f(row.iloc[10])
                        np_can    = limpar_f(row.iloc[11])
                        np_apa    = limpar_f(row.iloc[12])
                        np_eliq   = limpar_f(row.iloc[13])
                        np_aliq   = limpar_f(row.iloc[14])
                        if all(v == 0.0 for v in [proc_ins, proc_pag, proc_can, proc_apa,
                                                   np_ins, np_pag, np_can, np_apa, np_eliq, np_aliq]):
                            continue
                        dados_rp.append((
                            m_final, 2026,
                            cod_cred, nome_cred, num_emp, data_emp, dotacao,
                            proc_ins, proc_pag, proc_can, proc_apa,
                            np_ins, np_pag, np_can, np_apa, np_eliq, np_aliq
                        ))
                    except Exception:
                        continue
                conn.execute(
                    "DELETE FROM restos_a_pagar WHERE ano=2026 AND mes=?", (m_final,)
                )
                conn.executemany(
                    "INSERT INTO restos_a_pagar VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                    dados_rp
                )
                conn.commit()
                tot_ins_rp = sum(r[7] + r[11] for r in dados_rp)
                tot_pag_rp = sum(r[8] + r[12] for r in dados_rp)
                st.success(
                    "Restos a Pagar " + MESES_NOMES[m_final - 1] + "/2026: "
                    + str(len(dados_rp)) + " empenhos | "
                    + "Inscrito R$ {:,.2f} | Pagos R$ {:,.2f}".format(tot_ins_rp, tot_pag_rp)
                )

            conn.close()

        except Exception as e:
            st.error("Erro: " + str(e))
            import traceback
            st.code(traceback.format_exc())

    st.divider()
    st.subheader("Backup Completo")
    conn_b = sqlite3.connect(DB_NAME)
    tbls = {
        "receitas":        pd.read_sql("SELECT * FROM receitas",        conn_b),
        "orcamento":       pd.read_sql("SELECT * FROM orcamento",       conn_b),
        "execucao":        pd.read_sql("SELECT * FROM execucao",        conn_b),
        "sub_elementos":   pd.read_sql("SELECT * FROM sub_elementos",   conn_b),
        "anexo_v":         pd.read_sql("SELECT * FROM anexo_v",         conn_b),
        "restos_a_pagar":  pd.read_sql("SELECT * FROM restos_a_pagar",  conn_b),
    }
    conn_b.close()
    for nome_tab, df_tab in tbls.items():
        if not df_tab.empty:
            st.download_button(
                "Baixar " + nome_tab + " (CSV)",
                data=df_tab.to_csv(index=False).encode("utf-8"),
                file_name="backup_" + nome_tab + ".csv",
                mime="text/csv",
                key="bkp_" + nome_tab
            )
    st.caption("Restaurar tabela (CSV do backup):")
    tabela_rest = st.selectbox(
        "Tabela a restaurar:",
        ["receitas", "orcamento", "execucao", "sub_elementos", "anexo_v", "restos_a_pagar"],
        key="tabela_rest"
    )
    file_restore = st.file_uploader("Arquivo CSV", type=["csv"], key="file_rest")
    if file_restore and st.button("Restaurar"):
        df_res = pd.read_csv(file_restore)
        # Garantir que só entra dados da UO 03101 (TJMT)
        UO_PROJETO = "3101"
        if "uo" in df_res.columns:
            qtd_antes = len(df_res)
            df_res = df_res[df_res["uo"].astype(str).str.lstrip("0") == UO_PROJETO.lstrip("0")]
            filtradas = qtd_antes - len(df_res)
            if filtradas > 0:
                st.warning(
                    f"{filtradas} linha(s) de outras UOs removidas — "
                    f"apenas UO {UO_PROJETO} (TJMT) foi mantida."
                )
        conn_r = sqlite3.connect(DB_NAME)
        df_res.to_sql(tabela_rest, conn_r, if_exists="replace", index=False)
        conn_r.commit()
        conn_r.close()
        st.success(f"Tabela '{tabela_rest}' restaurada! ({len(df_res)} linhas)")
        st.rerun()

    st.divider()
    st.subheader("Limpeza Geral")
    confirma = st.checkbox("Confirmo apagar TODOS os dados")
    if st.button("Limpar Tudo"):
        if confirma:
            limpar_todos_dados()
            st.rerun()
        else:
            st.warning("Marque a caixa de confirmacao.")


# ---------------------------------------------------------------------------
# CARGA PRINCIPAL
# ---------------------------------------------------------------------------
conn_main = sqlite3.connect(DB_NAME)
df_rec    = pd.read_sql("SELECT * FROM receitas",        conn_main)
df_orc    = pd.read_sql("SELECT * FROM orcamento",       conn_main)
df_exec   = pd.read_sql("SELECT * FROM execucao",        conn_main)
df_sub    = pd.read_sql("SELECT * FROM sub_elementos",   conn_main)
df_anexov = pd.read_sql("SELECT * FROM anexo_v",         conn_main)
df_rp     = pd.read_sql("SELECT * FROM restos_a_pagar",  conn_main)
conn_main.close()

tab1, tab2, tab3, tab4, tab5 = st.tabs(["Receitas", "Despesas", "Comparativo", "Relatorios LRF", "Restos a Pagar"])


# ---------------------------------------------------------------------------
# ABA 1: RECEITAS
# ---------------------------------------------------------------------------
with tab1:
    if df_rec.empty:
        st.info("Importe dados de Receita (FIP 729) para visualizar.")
    else:
        with st.expander("Classificar Categorias"):
            c1, c2, c3 = st.columns([2, 2, 1])
            sel_nat = c1.selectbox(
                "Natureza:", sorted(df_rec["natureza"].unique()), key="sel_nat_c"
            )
            sel_cat = c2.selectbox("Categoria:", CATEGORIAS_REC, key="sel_cat_c")
            if c3.button("Salvar"):
                cu = sqlite3.connect(DB_NAME)
                cu.execute(
                    "UPDATE receitas SET categoria=? WHERE natureza=?",
                    (sel_cat, sel_nat)
                )
                cu.commit()
                cu.close()
                st.rerun()

        st.divider()
        f1, f2, f3 = st.columns(3)
        ms_r = f1.multiselect(
            "Meses:", sorted(df_rec["mes"].unique()),
            default=list(df_rec["mes"].unique()),
            format_func=lambda x: MESES_NOMES[x - 1], key="ms_r"
        )
        cat_sel = f2.multiselect(
            "Categoria:", sorted(df_rec["categoria"].unique()),
            default=list(df_rec["categoria"].unique()), key="cat_r"
        )
        nat_sel = f3.multiselect(
            "Natureza:", sorted(df_rec["natureza"].unique()), key="nat_r"
        )

        df_rf = df_rec[df_rec["mes"].isin(ms_r) & df_rec["categoria"].isin(cat_sel)]
        if nat_sel:
            df_rf = df_rf[df_rf["natureza"].isin(nat_sel)]

        if not df_rf.empty and ms_r:
            v_real     = float(df_rf["realizado"].sum())
            v_real_tot = float(df_rec["realizado"].sum())
            v_orc      = float(
                df_rec[df_rec["mes"] == max(ms_r)]
                .groupby("codigo_full")["orcado"].max().sum()
            )
            pct_filtro = v_real / v_real_tot * 100 if v_real_tot > 0 else 0
            pct_ating  = v_real / v_orc      * 100 if v_orc      > 0 else 0

            k1, k2, k3, k4 = st.columns(4)
            k1.metric("Orçado (Atualizado)",  "R$ {:,.2f}".format(v_orc))
            k2.metric("Realizado (Filtro)",   "R$ {:,.2f}".format(v_real))
            k3.metric("Atingimento",          "{:.1f}%".format(pct_ating))
            k4.metric("% do Filtro s/ Total", "{:.1f}%".format(pct_filtro))

            # Gráfico principal: orçado + realizado + previsão por mês
            df_g = df_rf.groupby("mes").agg(
                {"realizado": "sum", "orcado": "sum"}).reset_index()
            df_g["previsao"] = [
                df_rf[df_rf["mes"] == m].groupby("codigo_full")["previsao"].max().sum()
                for m in df_g["mes"]
            ]
            fig = go.Figure()
            fig.add_trace(go.Bar(
                x=[MESES_NOMES[m - 1] for m in df_g["mes"]],
                y=df_g["orcado"], name="Orçado", marker_color=COR_AZUL,
                opacity=0.7,
                text=["{:.1f}M".format(v / 1e6) for v in df_g["orcado"]],
                textposition="inside"
            ))
            fig.add_trace(go.Bar(
                x=[MESES_NOMES[m - 1] for m in df_g["mes"]],
                y=df_g["realizado"], name="Realizado", marker_color=COR_VERDE,
                text=["{:.1f}M".format(v / 1e6) for v in df_g["realizado"]],
                textposition="inside"
            ))
            fig.add_trace(go.Scatter(
                x=[MESES_NOMES[m - 1] for m in df_g["mes"]],
                y=df_g["previsao"], name="Previsão Mensal",
                line=dict(color="#FF9800", width=2.5, dash="dot"),
                mode="lines+markers"
            ))
            fig.update_layout(
                title="Receita por Mês — Orçado vs Realizado",
                height=380, barmode="group",
                margin=dict(l=0, r=0, t=40, b=0),
                hovermode="x unified",
                plot_bgcolor="white",
                legend=dict(orientation="h", yanchor="bottom", y=1.02,
                            xanchor="right", x=1),
                yaxis=dict(tickformat=",.0f", gridcolor="#F0F0F0")
            )
            st.plotly_chart(fig, use_container_width=True)

            # Gráfico horizontal: realizado por categoria com %
            df_cat = (df_rf.groupby("categoria")
                      .agg({"realizado": "sum", "orcado": "sum"})
                      .reset_index()
                      .sort_values("realizado", ascending=True))
            tot_cat = float(df_cat["realizado"].sum())
            df_cat["pct"] = df_cat["realizado"].apply(
                lambda v: v / tot_cat * 100 if tot_cat > 0 else 0)
            df_cat["pct_ating"] = df_cat.apply(
                lambda r: r["realizado"] / r["orcado"] * 100 if r["orcado"] > 0 else 0,
                axis=1)
            fig_cat = go.Figure()
            fig_cat.add_trace(go.Bar(
                y=df_cat["categoria"],
                x=df_cat["orcado"], name="Orçado",
                orientation="h", marker_color=COR_AZUL, opacity=0.6
            ))
            fig_cat.add_trace(go.Bar(
                y=df_cat["categoria"],
                x=df_cat["realizado"], name="Realizado",
                orientation="h", marker_color=COR_VERDE,
                text=["{:.1f}%  ({:.1f}% s/total)".format(r["pct_ating"], r["pct"])
                      for _, r in df_cat.iterrows()],
                textposition="outside"
            ))
            fig_cat.update_layout(
                title="Realizado por Categoria — % Atingimento e % s/ Total",
                height=max(260, len(df_cat) * 60),
                barmode="overlay",
                margin=dict(l=0, r=0, t=40, b=0),
                plot_bgcolor="white",
                legend=dict(orientation="h", yanchor="bottom", y=1.02,
                            xanchor="right", x=1),
                xaxis=dict(tickformat=",.0f", gridcolor="#F0F0F0")
            )
            st.plotly_chart(fig_cat, use_container_width=True)

            st.dataframe(
                df_rf[["categoria", "codigo_full", "natureza",
                        "realizado", "orcado"]]
                .assign(pct_total=lambda d:
                    (d["realizado"] / v_real_tot * 100).round(2).astype(str) + "%")
                .rename(columns={"pct_total": "% s/ Total"})
                .style.format({"realizado": "{:,.2f}", "orcado": "{:,.2f}"}),
                use_container_width=True
            )

            # Exportar PDF
            st.divider()
            filtros_str_r = (
                ("Categorias: " + ", ".join(cat_sel) if cat_sel else "") +
                ("  |  Naturezas: " + ", ".join(str(n) for n in nat_sel) if nat_sel else "")
            ).strip(" | ") or "Todos"
            try:
                pdf_bytes = gerar_pdf_receitas(df_rf, df_rec, ms_r, filtros_str_r)
                st.download_button(
                    "📄 Exportar Relatório PDF — Receitas",
                    data=pdf_bytes,
                    file_name="relatorio_receitas.pdf",
                    mime="application/pdf",
                    key="pdf_rec"
                )
            except Exception as e:
                st.warning("PDF indisponível: " + str(e))

    # -----------------------------------------------------------------------
    # REPASSES RECEBIDOS (ANEXO V)
    # -----------------------------------------------------------------------
    st.divider()
    st.subheader("Repasses Recebidos - ANEXO V")

    if df_anexov.empty:
        st.info("Importe dados de Repasses Recebidos (ANEXO V) para visualizar.")
    else:
        av1, av2 = st.columns(2)
        meses_av = sorted(df_anexov["mes"].unique())
        ms_av = av1.multiselect(
            "Meses:", meses_av,
            default=meses_av,
            format_func=lambda x: MESES_NOMES[x - 1],
            key="ms_av"
        )
        entidades_disp = sorted(df_anexov["entidade_repassadora"].dropna().unique())
        entidade_sel = av2.multiselect(
            "Entidade Repassadora:", entidades_disp,
            default=entidades_disp,
            key="entidade_av"
        )

        df_avf = df_anexov[
            df_anexov["mes"].isin(ms_av) &
            df_anexov["entidade_repassadora"].isin(entidade_sel)
        ].copy()

        if not df_avf.empty:
            total_repasses = df_avf["valor"].sum()
            ka1, ka2 = st.columns(2)
            ka1.metric("Total Repassado", "R$ {:,.2f}".format(total_repasses))
            ka2.metric("Quantidade de Repasses", str(len(df_avf)))

            # Grafico por entidade repassadora
            df_por_entidade = (
                df_avf.groupby("entidade_repassadora")["valor"]
                .sum()
                .reset_index()
                .sort_values("valor", ascending=False)
            )
            fig_av = go.Figure(go.Bar(
                x=df_por_entidade["entidade_repassadora"],
                y=df_por_entidade["valor"],
                marker_color="#1565C0",
                text=df_por_entidade["valor"].apply(
                    lambda v: "R$ {:,.0f}".format(v)
                ),
                textposition="outside"
            ))
            fig_av.update_layout(
                height=350,
                margin=dict(l=0, r=0, t=30, b=0),
                xaxis_tickangle=-30,
                yaxis_title="Valor (R$)"
            )
            st.plotly_chart(fig_av, width='stretch')

            st.dataframe(
                df_avf[["mes", "data", "entidade_repassadora", "valor", "finalidade", "fundamento_legal"]]
                .rename(columns={
                    "mes": "Mes",
                    "data": "Data",
                    "entidade_repassadora": "Entidade Repassadora",
                    "valor": "Valor (R$)",
                    "finalidade": "Finalidade",
                    "fundamento_legal": "Fundamento Legal"
                })
                .style.format({"Valor (R$)": "{:,.2f}"}),
                width='stretch'
            )
        else:
            st.info("Nenhum dado para os filtros selecionados.")


# ---------------------------------------------------------------------------
# ABA 2: DESPESAS
# ---------------------------------------------------------------------------
with tab2:
    has_orc  = not df_orc.empty
    has_exec = not df_exec.empty

    if not has_orc and not has_exec:
        st.info(
            "Importe 'Orcamento (FIP 616)' e 'Execucao (FIP 613)' para visualizar."
        )
    else:
        meses_exec = sorted(df_exec["mes"].unique().tolist()) if has_exec else []
        ugs_disp   = sorted(df_exec["ug"].unique().tolist())  if has_exec else []

        f1, f2, f3 = st.columns(3)
        ms_d = f1.multiselect(
            "Meses:", meses_exec, default=meses_exec,
            format_func=lambda x: MESES_NOMES[x - 1], key="ms_d"
        )
        ug_sel = f2.multiselect(
            "UG (Unidade Gestora):", ugs_disp, default=ugs_disp, key="ug_d"
        )
        fs = f3.multiselect(
            "Funcao:",
            sorted(df_exec["funcao"].unique()) if has_exec else [],
            key="func_d"
        )

        f4, f5, f6 = st.columns(3)
        sf = f4.multiselect(
            "Subfuncao:",
            sorted(df_exec["subfuncao"].unique()) if has_exec else [],
            key="subf_d"
        )
        ps = f5.multiselect(
            "Programa:",
            sorted(df_exec["programa"].unique()) if has_exec else [],
            key="prog_d"
        )
        fts = f6.multiselect(
            "Fonte:",
            sorted(df_exec["fonte"].unique()) if has_exec else [],
            key="font_d"
        )
        nats_disp = sorted(df_exec["natureza"].dropna().unique().tolist()) if has_exec else []
        bd = st.multiselect("Natureza:", nats_disp, key="busca_d")

        df_ef = df_exec[df_exec["mes"].isin(ms_d)].copy() if has_exec else pd.DataFrame()
        if ug_sel and not df_ef.empty:
            df_ef = df_ef[df_ef["ug"].isin(ug_sel)]
        if fs and not df_ef.empty:
            df_ef = df_ef[df_ef["funcao"].isin(fs)]
        if sf and not df_ef.empty:
            df_ef = df_ef[df_ef["subfuncao"].isin(sf)]
        if ps and not df_ef.empty:
            df_ef = df_ef[df_ef["programa"].isin(ps)]
        if fts and not df_ef.empty:
            df_ef = df_ef[df_ef["fonte"].isin(fts)]
        if bd and not df_ef.empty:
            df_ef = df_ef[df_ef["natureza"].isin(bd)]

        m_max_orc = int(df_orc["mes"].max()) if has_orc else 0
        m_max_sel = max(ms_d) if ms_d else m_max_orc

        cred_total = (
            df_orc[df_orc["mes"] == m_max_orc]["cred_autorizado"].sum()
            if has_orc else 0
        )

        emp_total = float(df_ef["empenhado"].sum()) if not df_ef.empty else 0
        liq_total = float(df_ef["liquidado"].sum()) if not df_ef.empty else 0
        pag_total = float(df_ef["pago"].sum())      if not df_ef.empty else 0
        cred_total = float(cred_total)

        pct_emp = emp_total / cred_total * 100 if cred_total > 0 else 0
        pct_liq = liq_total / cred_total * 100 if cred_total > 0 else 0
        pct_pag = pag_total / cred_total * 100 if cred_total > 0 else 0

        k1, k2, k3, k4 = st.columns(4)
        k1.metric(
            "Cred. Autorizado (" + (MESES_NOMES[m_max_orc - 1] if m_max_orc else "—") + ")",
            "R$ {:,.2f}".format(cred_total)
        )
        k2.metric("Empenhado",  "R$ {:,.2f}".format(emp_total),
                  delta="{:.1f}% do cred.".format(pct_emp))
        k3.metric("Liquidado",  "R$ {:,.2f}".format(liq_total),
                  delta="{:.1f}% do cred.".format(pct_liq))
        k4.metric("Pago",       "R$ {:,.2f}".format(pag_total),
                  delta="{:.1f}% do cred.".format(pct_pag))

        if not df_ef.empty:
            df_g = (df_ef.groupby("mes")[["empenhado", "liquidado", "pago"]]
                    .sum().reset_index())
            fig = go.Figure()
            fig.add_trace(go.Bar(
                x=[MESES_NOMES[m - 1] for m in df_g["mes"]],
                y=df_g["empenhado"], name="Empenhado", marker_color=COR_AZUL,
                text=["{:.1f}%".format(v / cred_total * 100)
                      if cred_total > 0 else "" for v in df_g["empenhado"]],
                textposition="inside"
            ))
            fig.add_trace(go.Bar(
                x=[MESES_NOMES[m - 1] for m in df_g["mes"]],
                y=df_g["liquidado"], name="Liquidado", marker_color=COR_VERDE,
                text=["{:.1f}%".format(v / cred_total * 100)
                      if cred_total > 0 else "" for v in df_g["liquidado"]],
                textposition="inside"
            ))
            fig.add_trace(go.Bar(
                x=[MESES_NOMES[m - 1] for m in df_g["mes"]],
                y=df_g["pago"], name="Pago", marker_color=COR_LARAN,
                text=["{:.1f}%".format(v / cred_total * 100)
                      if cred_total > 0 else "" for v in df_g["pago"]],
                textposition="inside"
            ))
            # Linha de crédito autorizado
            meses_labels = [MESES_NOMES[m - 1] for m in df_g["mes"]]
            fig.add_trace(go.Scatter(
                x=meses_labels, y=[cred_total] * len(meses_labels),
                name="Cred. Autorizado", mode="lines",
                line=dict(color="#B71C1C", width=2, dash="dash")
            ))
            fig.update_layout(
                title="Despesas por Mês — % sobre Crédito Autorizado",
                height=380, barmode="group",
                margin=dict(l=0, r=0, t=40, b=0),
                hovermode="x unified", plot_bgcolor="white",
                legend=dict(orientation="h", yanchor="bottom", y=1.02,
                            xanchor="right", x=1),
                yaxis=dict(tickformat=",.0f", gridcolor="#F0F0F0")
            )
            st.plotly_chart(fig, use_container_width=True)

            # Horizontal por natureza (top 20 por liquidado)
            df_nat_agg = (df_ef.groupby("natureza")[["empenhado", "liquidado"]]
                          .sum().reset_index()
                          .sort_values("liquidado", ascending=True)
                          .tail(20))
            # Garantir que os rótulos do eixo Y sejam strings (evita escala numérica)
            df_nat_agg["nat_label"] = df_nat_agg["natureza"].astype(str).str.strip()
            tot_liq_nat = float(df_nat_agg["liquidado"].sum())
            fig_nat = go.Figure()
            fig_nat.add_trace(go.Bar(
                y=df_nat_agg["nat_label"],
                x=df_nat_agg["empenhado"], name="Empenhado",
                orientation="h", marker_color=COR_AZUL, opacity=0.65
            ))
            fig_nat.add_trace(go.Bar(
                y=df_nat_agg["nat_label"],
                x=df_nat_agg["liquidado"], name="Liquidado",
                orientation="h", marker_color=COR_VERDE,
                text=["{:.1f}%".format(v / tot_liq_nat * 100)
                      if tot_liq_nat > 0 else "" for v in df_nat_agg["liquidado"]],
                textposition="auto"
            ))
            fig_nat.update_layout(
                title="Liquidado por Natureza (Top 20) — % s/ Total Liquidado",
                height=max(300, len(df_nat_agg) * 42),
                barmode="group",
                margin=dict(l=130, r=60, t=45, b=0),
                plot_bgcolor="white",
                legend=dict(orientation="h", yanchor="bottom", y=1.02,
                            xanchor="right", x=1),
                xaxis=dict(tickformat=",.0f", gridcolor="#EEEEEE", rangemode="tozero"),
                yaxis=dict(type="category", automargin=True, tickfont=dict(size=11))
            )
            st.plotly_chart(fig_nat, use_container_width=True)

            ug_filtrada = set(ug_sel) != set(ugs_disp)
            col_chave = (["ug"] if ug_filtrada else []) + [
                "funcao", "subfuncao", "programa", "projeto", "fonte", "natureza"
            ]
            df_agg = df_ef.groupby(col_chave, as_index=False)[
                ["empenhado", "liquidado", "pago"]
            ].sum()
            st.dataframe(
                df_agg[col_chave + ["empenhado", "liquidado", "pago"]]
                .style.format({"empenhado": "{:,.2f}",
                               "liquidado": "{:,.2f}", "pago": "{:,.2f}"}),
                use_container_width=True
            )

            st.divider()
            filtros_str_d = " | ".join(filter(None, [
                ("UGs: " + ", ".join(ug_sel)) if set(ug_sel) != set(ugs_disp) else "",
                ("Meses: " + ", ".join(MESES_NOMES[m-1] for m in ms_d)) if ms_d else "",
                ("Natureza: " + ", ".join(str(n) for n in bd)) if bd else "",
            ])) or "Todos"
            try:
                pdf_bytes_d = gerar_pdf_despesas(df_ef, cred_total, ms_d, filtros_str_d)
                st.download_button(
                    "📄 Exportar Relatório PDF — Despesas",
                    data=pdf_bytes_d,
                    file_name="relatorio_despesas.pdf",
                    mime="application/pdf",
                    key="pdf_desp"
                )
            except Exception as e:
                st.warning("PDF indisponível: " + str(e))

        if not df_sub.empty:
            st.divider()
            with st.expander("Sub-elementos por PAOE (FIP 701)"):
                meses_sub = sorted(df_sub["mes"].unique())
                fontes_sub = (
                    sorted(df_sub["fonte"].dropna().unique())
                    if "fonte" in df_sub.columns else []
                )

                fs1, fs2, fs3 = st.columns(3)
                ms_s = fs1.multiselect(
                    "Meses:", meses_sub,
                    default=meses_sub,
                    format_func=lambda x: MESES_NOMES[x - 1], key="ms_s"
                )
                paoe_s = fs2.multiselect(
                    "PAOE:", sorted(df_sub["paoe"].unique()), key="paoe_s"
                )
                nat_s = fs3.multiselect(
                    "Natureza:", sorted(df_sub["natureza_cod"].unique()), key="nat_s"
                )

                fs4, fs5, fs6 = st.columns(3)
                fonte_s = fs4.multiselect(
                    "Fonte:", fontes_sub, key="fonte_s"
                )
                sub_busca = fs5.text_input(
                    "Buscar sub-elemento (cod. ou desc.):",
                    key="sub_busca",
                    placeholder="Ex: 90.14.14.001 ou Diarias"
                )
                if sub_busca:
                    mask_sub = (
                        df_sub["subelemento_cod"].str.contains(sub_busca, case=False, na=False)
                        | df_sub["subelemento_desc"].str.contains(sub_busca, case=False, na=False)
                    )
                    subs_disp = sorted(df_sub.loc[mask_sub, "subelemento_desc"].dropna().unique().tolist())
                else:
                    subs_disp = sorted(df_sub["subelemento_desc"].dropna().unique().tolist())
                sub_sel = fs5.multiselect(
                    "Sub-elemento:", subs_disp, key="sub_sel"
                )
                ugs_sub = (
                    sorted(df_sub["ug"].dropna().unique().tolist())
                    if "ug" in df_sub.columns and not df_sub.empty else []
                )
                ug_sel_s = fs6.multiselect(
                    "UG:", ugs_sub, key="ug_sub"
                )

                df_sf = df_sub[df_sub["mes"].isin(ms_s)].copy()
                if ug_sel_s and "ug" in df_sf.columns:
                    df_sf = df_sf[df_sf["ug"].isin(ug_sel_s)]
                if paoe_s:
                    df_sf = df_sf[df_sf["paoe"].isin(paoe_s)]
                if nat_s:
                    df_sf = df_sf[df_sf["natureza_cod"].isin(nat_s)]
                if fonte_s and "fonte" in df_sf.columns:
                    df_sf = df_sf[df_sf["fonte"].isin(fonte_s)]
                if sub_sel:
                    df_sf = df_sf[df_sf["subelemento_desc"].isin(sub_sel)]

                if not df_sf.empty:
                    has_fonte = "fonte" in df_sf.columns
                    has_ug    = "ug"    in df_sf.columns
                    col_s = []
                    if has_ug:    col_s += ["ug"]
                    col_s += ["paoe", "natureza_cod", "natureza_desc"]
                    if has_fonte: col_s += ["fonte"]
                    col_s += ["subelemento_cod", "subelemento_desc"]

                    df_sv = df_sf.groupby(col_s, as_index=False)[
                        ["liquidado", "pago"]].sum()

                    liq_sel_701  = float(df_sv["liquidado"].sum())
                    pago_sel_701 = float(df_sv["pago"].sum())
                    liq_tot_701  = float(df_sub["liquidado"].sum())
                    pago_tot_701 = float(df_sub["pago"].sum())
                    pct_liq_701  = liq_sel_701 / liq_tot_701 * 100 if liq_tot_701 > 0 else 0
                    pct_pag_701  = pago_sel_701 / pago_tot_701 * 100 if pago_tot_701 > 0 else 0

                    ks1, ks2, ks3, ks4 = st.columns(4)
                    ks1.metric("Liquidado (Filtro)",
                               "R$ {:,.2f}".format(liq_sel_701))
                    ks2.metric("% s/ Total Liquidado",
                               "{:.1f}%".format(pct_liq_701))
                    ks3.metric("Pago (Filtro)",
                               "R$ {:,.2f}".format(pago_sel_701))
                    ks4.metric("% s/ Total Pago",
                               "{:.1f}%".format(pct_pag_701))

                    # Gráfico horizontal por sub-elemento (top 20)
                    df_sub_plot = (df_sv.groupby("subelemento_desc")["liquidado"]
                                   .sum().reset_index()
                                   .sort_values("liquidado", ascending=True).tail(20))
                    tot_sp = float(df_sub_plot["liquidado"].sum())
                    fig_701 = go.Figure(go.Bar(
                        y=df_sub_plot["subelemento_desc"],
                        x=df_sub_plot["liquidado"],
                        orientation="h",
                        marker_color=COR_AZUL,
                        text=["{:.1f}%".format(v / tot_sp * 100)
                              if tot_sp > 0 else "" for v in df_sub_plot["liquidado"]],
                        textposition="outside"
                    ))
                    fig_701.update_layout(
                        title="Liquidado por Sub-elemento (Top 20) — % s/ Total",
                        height=max(350, len(df_sub_plot) * 32),
                        margin=dict(l=0, r=80, t=40, b=0),
                        plot_bgcolor="white",
                        xaxis=dict(tickformat=",.0f", gridcolor="#F0F0F0")
                    )
                    st.plotly_chart(fig_701, use_container_width=True)

                    # Gráfico por natureza
                    df_nat_701 = (df_sv.groupby(["natureza_cod", "natureza_desc"])["liquidado"]
                                  .sum().reset_index()
                                  .sort_values("liquidado", ascending=True))
                    tot_nat_701 = float(df_nat_701["liquidado"].sum())
                    fig_nat_701 = go.Figure(go.Bar(
                        y=df_nat_701["natureza_desc"],
                        x=df_nat_701["liquidado"],
                        orientation="h",
                        marker_color=COR_VERDE,
                        text=["{:.1f}%".format(v / tot_nat_701 * 100)
                              if tot_nat_701 > 0 else "" for v in df_nat_701["liquidado"]],
                        textposition="outside"
                    ))
                    fig_nat_701.update_layout(
                        title="Liquidado por Natureza de Despesa — % s/ Total",
                        height=max(260, len(df_nat_701) * 40),
                        margin=dict(l=0, r=80, t=40, b=0),
                        plot_bgcolor="white",
                        xaxis=dict(tickformat=",.0f", gridcolor="#F0F0F0")
                    )
                    st.plotly_chart(fig_nat_701, use_container_width=True)

                    st.dataframe(
                        df_sv[col_s + ["liquidado", "pago"]]
                        .style.format({"liquidado": "{:,.2f}", "pago": "{:,.2f}"}),
                        use_container_width=True
                    )

                    st.divider()
                    filtros_701 = " | ".join(filter(None, [
                        ("UG: " + ", ".join(ug_sel_s)) if ug_sel_s else "",
                        ("PAOE: " + ", ".join(paoe_s)) if paoe_s else "",
                        ("Natureza: " + ", ".join(nat_s)) if nat_s else "",
                    ])) or "Todos"
                    try:
                        pdf_701 = gerar_pdf_701(df_sv, df_sub, ms_s, filtros_701)
                        st.download_button(
                            "📄 Exportar Relatório PDF — Sub-elementos (FIP 701)",
                            data=pdf_701,
                            file_name="relatorio_701.pdf",
                            mime="application/pdf",
                            key="pdf_701"
                        )
                    except Exception as e:
                        st.warning("PDF indisponível: " + str(e))
                else:
                    st.info("Nenhum dado para os filtros selecionados.")


# ---------------------------------------------------------------------------
# ABA 3: COMPARATIVO
# ---------------------------------------------------------------------------
with tab3:
    st.subheader("Confronto Geral - Receita x Despesa")
    if df_rec.empty and df_exec.empty:
        st.info("Importe dados para visualizar.")
    else:
        todos = sorted(set(
            (df_rec["mes"].tolist() if not df_rec.empty else [])
            + (df_exec["mes"].tolist() if not df_exec.empty else [])
        ))
        ms_c = st.multiselect(
            "Meses:", todos, default=todos,
            format_func=lambda x: MESES_NOMES[x - 1], key="ms_c"
        )
        tr = (
            df_rec[df_rec["mes"].isin(ms_c)]["realizado"].sum()
            if not df_rec.empty else 0
        )
        te = (
            df_exec[df_exec["mes"].isin(ms_c)]["empenhado"].sum()
            if not df_exec.empty else 0
        )
        tl = (
            df_exec[df_exec["mes"].isin(ms_c)]["liquidado"].sum()
            if not df_exec.empty else 0
        )
        tp = (
            df_exec[df_exec["mes"].isin(ms_c)]["pago"].sum()
            if not df_exec.empty else 0
        )

        kc1, kc2, kc3, kc4 = st.columns(4)
        kc1.metric("Receita Arrecadada", "R$ {:,.2f}".format(tr))
        kc2.metric("Desp. Empenhada",    "R$ {:,.2f}".format(te))
        kc3.metric("Desp. Liquidada",    "R$ {:,.2f}".format(tl))
        kc4.metric("Desp. Paga",         "R$ {:,.2f}".format(tp))

        st.divider()
        m1, m2 = st.columns(2)
        m1.info(
            "Superavit Financeiro (Receita - Pago): R$ {:,.2f}".format(tr - tp)
        )
        m2.warning(
            "Superavit Orcamentario (Receita - Empenhado): R$ {:,.2f}".format(tr - te)
        )

        fig_c = go.Figure()
        fig_c.add_trace(
            go.Bar(name="Receita", x=["Confronto"], y=[tr], marker_color="green")
        )
        fig_c.add_trace(
            go.Bar(name="Empenhado", x=["Confronto"], y=[te], marker_color="orange")
        )
        fig_c.add_trace(
            go.Bar(name="Liquidado", x=["Confronto"], y=[tl], marker_color="#72A0C1")
        )
        fig_c.add_trace(
            go.Bar(name="Pago", x=["Confronto"], y=[tp], marker_color="red")
        )
        fig_c.update_layout(
            height=400, barmode="group", margin=dict(l=0, r=0, t=30, b=0)
        )
        st.plotly_chart(fig_c, width='stretch')

        v_orc_comp = float(
            df_orc[df_orc["mes"] == int(df_orc["mes"].max())]["cred_autorizado"].sum()
            if not df_orc.empty else 0
        )
        try:
            pdf_comp = gerar_pdf_comparativo(
                float(tr), float(te), float(tl), float(tp),
                v_orc_comp, list(ms_c), df_rec, df_exec
            )
            st.download_button(
                "📄 Exportar Relatório PDF — Comparativo",
                data=pdf_comp,
                file_name="relatorio_comparativo.pdf",
                mime="application/pdf",
                key="pdf_comp"
            )
        except Exception as e:
            st.warning("PDF indisponível: " + str(e))


# ---------------------------------------------------------------------------
# ABA 4: RELATORIOS LRF
# ---------------------------------------------------------------------------
with tab4:
    st.subheader("Relatorios LRF - Anexos Bimestrais")
    if df_rec.empty and df_exec.empty and df_orc.empty:
        st.info("Importe dados de Receita, Orcamento e Execucao para gerar os anexos.")
    else:
        bim = st.selectbox("Bimestre de referencia:", list(BIMESTRES.keys()), key="bim_lrf")
        meses_bim = BIMESTRES[bim]
        meses_ate_agora = list(range(1, max(meses_bim) + 1))

        c1, c2, c3 = st.columns(3)

        c1.write("**Anexo I — Receitas**")
        c1.caption("Previsao x Realizado por categoria/natureza")
        c1.download_button(
            "Baixar Anexo I (.xlsx)",
            data=gerar_excel_anexo1(df_rec, meses_bim, meses_ate_agora),
            file_name="AnexoI_LRF.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="lrf1"
        )

        c2.write("**Anexo IA — Despesas por Natureza**")
        c2.caption("Dotacao x Empenhado x Liquidado x Pago por natureza")
        c2.download_button(
            "Baixar Anexo IA (.xlsx)",
            data=gerar_excel_anexo1a(df_orc, df_exec, df_rec, meses_bim, meses_ate_agora),
            file_name="AnexoIA_LRF.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="lrf1a"
        )

        c3.write("**Anexo II — Despesas por Funcao**")
        c3.caption("Dotacao x Empenhado x Liquidado por funcao/subfuncao")
        c3.download_button(
            "Baixar Anexo II (.xlsx)",
            data=gerar_excel_anexo2(df_orc, df_exec, meses_bim, meses_ate_agora),
            file_name="AnexoII_LRF.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="lrf2"
        )

        if not df_rp.empty:
            st.divider()
            c4, _, _ = st.columns(3)
            c4.write("**Anexo RP — Restos a Pagar**")
            c4.caption("Inscrito x Pagos x Cancelados x A Liquidar x Liquidados (FIP 226)")
            c4.download_button(
                "Baixar Anexo RP (.xlsx)",
                data=gerar_excel_rp(df_rp, meses_bim, meses_ate_agora),
                file_name="AnexoRP_LRF.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="lrf_rp"
            )


# ---------------------------------------------------------------------------
# ABA 5: RESTOS A PAGAR
# ---------------------------------------------------------------------------
with tab5:
    st.subheader("Restos a Pagar — FIP 226")
    if df_rp.empty:
        st.info("Importe arquivos FIP 226 (Restos a Pagar) pela barra lateral para visualizar.")
    else:
        meses_rp_disp = sorted(df_rp["mes"].unique())

        def _ext_fonte_rp(dot):
            try:
                p = str(dot).split(".")
                return (p[-3] + "." + p[-2]).strip() if len(p) >= 3 else "N/D"
            except Exception:
                return "N/D"

        fc1, fc2, fc3, fc4 = st.columns([2, 2, 2, 2])
        ms_rp = fc1.multiselect(
            "Mês de referência:",
            meses_rp_disp,
            default=[max(meses_rp_disp)],
            format_func=lambda x: MESES_NOMES[x - 1],
            key="ms_rp"
        )
        tipo_proc = fc2.radio(
            "Tipo de RP:",
            ["Todos", "Processados", "Não Processados"],
            horizontal=True,
            key="tipo_proc_rp"
        )
        tipo_rp = fc3.radio(
            "Visualizar por:",
            ["Resumo Geral", "Por Credor", "Por Dotação"],
            horizontal=True,
            key="tipo_rp"
        )

        if not ms_rp:
            st.warning("Selecione ao menos um mês.")
        else:
            mes_ref_rp = max(ms_rp)
            df_rp_ref = df_rp[df_rp["mes"] == mes_ref_rp].copy()
            df_rp_ref["_fonte"] = df_rp_ref["dotacao"].apply(_ext_fonte_rp)
            fontes_disp_rp = sorted(df_rp_ref["_fonte"].unique())
            fontes_sel = fc4.multiselect(
                "Fonte:", fontes_disp_rp, default=fontes_disp_rp, key="fonte_rp"
            )
            if fontes_sel:
                df_rp_ref = df_rp_ref[df_rp_ref["_fonte"].isin(fontes_sel)]

            # ---- Cálculos de totais ----------------------------------------
            proc_ins  = float(df_rp_ref["proc_inscrito"].sum())
            proc_pag  = float(df_rp_ref["proc_pagos"].sum())
            proc_can  = float(df_rp_ref["proc_cancelados"].sum())
            proc_apa  = float(df_rp_ref["proc_a_pagar"].sum())
            np_ins    = float(df_rp_ref["np_inscrito"].sum())
            np_pag    = float(df_rp_ref["np_pagos"].sum())
            np_can    = float(df_rp_ref["np_cancelados"].sum())
            np_apa    = float(df_rp_ref["np_a_pagar"].sum())
            np_eliq   = float(df_rp_ref["np_em_liquidacao"].sum())
            np_aliq   = float(df_rp_ref["np_a_liquidar"].sum())

            # Liquidados NP = inscritos - a_liquidar - cancelados (fórmula do usuário)
            # = np_pag + np_apa + np_eliq (equivalente pela equação de balanço do FIP)
            np_liquidados = max(0.0, np_ins - np_aliq - np_can)
            # Processados: todos são liquidados por definição (inscrito = pag + can + a_pagar)
            proc_liquidados = proc_ins  # todos já foram liquidados ao ser inscrito

            if tipo_proc == "Processados":
                tot_inscrito   = proc_ins
                tot_pagos      = proc_pag
                tot_cancelados = proc_can
                tot_a_liquidar = proc_apa     # proc: "a liquidar" = a pagar
                tot_liquidados = proc_liquidados
            elif tipo_proc == "Não Processados":
                tot_inscrito   = np_ins
                tot_pagos      = np_pag
                tot_cancelados = np_can
                tot_a_liquidar = np_aliq
                tot_liquidados = np_liquidados
            else:
                tot_inscrito   = proc_ins + np_ins
                tot_pagos      = proc_pag + np_pag
                tot_cancelados = proc_can + np_can
                tot_a_liquidar = np_aliq      # só NP tem "a liquidar"; proc tem "a pagar"
                tot_liquidados = np_liquidados # liquidados = conceito NP (proc já são todos liquidados)

            pct_pago = tot_pagos / tot_inscrito * 100 if tot_inscrito > 0 else 0

            # ---- KPIs -------------------------------------------------------
            k1, k2, k3, k4, k5 = st.columns(5)
            k1.metric("Inscrito",        "R$ {:,.2f}".format(tot_inscrito))
            k2.metric("NP Liquidados",   "R$ {:,.2f}".format(tot_liquidados))
            k3.metric("Pagos (acum.)",   "R$ {:,.2f}".format(tot_pagos))
            k4.metric("Cancelados",      "R$ {:,.2f}".format(tot_cancelados))
            k5.metric("NP A Liquidar",   "R$ {:,.2f}".format(tot_a_liquidar))

            st.caption(
                "Referência: **{}** | Filtro: **{}** | Pago = **{:.1f}%** do Inscrito | "
                "NP Liquidados = np_inscrito − np_a_liquidar − np_cancelados".format(
                    MESES_NOMES[mes_ref_rp - 1], tipo_proc, pct_pago
                )
            )
            if tipo_proc == "Todos":
                st.info(
                    "Proc. A Pagar (liquidados aguardando pagamento): **R$ {:,.2f}**  |  "
                    "NP A Pagar (liquidados em trânsito): **R$ {:,.2f}**  |  "
                    "NP Em Liquidação: **R$ {:,.2f}**".format(proc_apa, np_apa, np_eliq)
                )
            st.divider()

            # ---- Gráfico: evolução mensal -----------------------------------
            if len(meses_rp_disp) > 1:
                st.markdown("#### Evolução Mensal")
                evo = []
                for m in meses_rp_disp:
                    dm = df_rp[df_rp["mes"] == m]
                    if tipo_proc == "Processados":
                        ins_m  = float(dm["proc_inscrito"].sum())
                        pag_m  = float(dm["proc_pagos"].sum())
                        can_m  = float(dm["proc_cancelados"].sum())
                        aliq_m = float(dm["proc_a_pagar"].sum())
                        liq_m  = ins_m
                    elif tipo_proc == "Não Processados":
                        ins_m  = float(dm["np_inscrito"].sum())
                        pag_m  = float(dm["np_pagos"].sum())
                        can_m  = float(dm["np_cancelados"].sum())
                        aliq_m = float(dm["np_a_liquidar"].sum())
                        liq_m  = max(0.0, ins_m - aliq_m - can_m)
                    else:
                        ins_m  = float((dm["proc_inscrito"] + dm["np_inscrito"]).sum())
                        pag_m  = float((dm["proc_pagos"]    + dm["np_pagos"]).sum())
                        can_m  = float((dm["proc_cancelados"] + dm["np_cancelados"]).sum())
                        np_ins_m = float(dm["np_inscrito"].sum())
                        np_aliq_m = float(dm["np_a_liquidar"].sum())
                        np_can_m  = float(dm["np_cancelados"].sum())
                        aliq_m = float(dm["np_a_liquidar"].sum())
                        liq_m  = max(0.0, np_ins_m - np_aliq_m - np_can_m)
                    evo.append({"mes": m, "inscrito": ins_m, "pagos": pag_m,
                                "cancelados": can_m, "a_liquidar": aliq_m,
                                "liquidados": liq_m})
                df_evo = pd.DataFrame(evo)
                labels_evo = [MESES_NOMES[m - 1] for m in df_evo["mes"]]

                fig_evo = go.Figure()
                fig_evo.add_trace(go.Bar(
                    name="Inscrito", x=labels_evo, y=df_evo["inscrito"],
                    marker_color=COR_AZUL, opacity=0.7))
                fig_evo.add_trace(go.Bar(
                    name="Liquidados", x=labels_evo, y=df_evo["liquidados"],
                    marker_color=COR_VERDE, opacity=0.85))
                fig_evo.add_trace(go.Bar(
                    name="Pagos", x=labels_evo, y=df_evo["pagos"],
                    marker_color="#43A047", opacity=0.9))
                fig_evo.add_trace(go.Bar(
                    name="Cancelados", x=labels_evo, y=df_evo["cancelados"],
                    marker_color=COR_LARAN, opacity=0.85))
                fig_evo.add_trace(go.Bar(
                    name="A Liquidar", x=labels_evo, y=df_evo["a_liquidar"],
                    marker_color="#78909C", opacity=0.85))
                fig_evo.update_layout(
                    barmode="group", height=380,
                    margin=dict(l=0, r=0, t=30, b=0),
                    legend=dict(orientation="h", y=-0.15),
                    yaxis=dict(tickformat=",.0f")
                )
                st.plotly_chart(fig_evo, use_container_width=True)
                st.divider()

            # ---- Gráfico: Processados vs Não Processados -------------------
            st.markdown("#### Processados vs Não Processados")
            col_g1, col_g2 = st.columns(2)
            fig_tipos = go.Figure()
            fig_tipos.add_trace(go.Bar(
                name="Processados",
                x=["Inscrito", "Pagos", "Cancelados", "A Pagar"],
                y=[proc_ins, proc_pag, proc_can, proc_apa],
                marker_color=COR_AZUL))
            fig_tipos.add_trace(go.Bar(
                name="Não Processados",
                x=["Inscrito", "Pagos", "Cancelados", "A Liquidar"],
                y=[np_ins, np_pag, np_can, np_aliq],
                marker_color=COR_VERDE))
            fig_tipos.update_layout(
                barmode="group", height=350,
                margin=dict(l=0, r=0, t=30, b=0),
                yaxis=dict(tickformat=",.0f")
            )
            col_g1.plotly_chart(fig_tipos, use_container_width=True)

            # Pizza: composição do inscrito
            fig_pizza = go.Figure(go.Pie(
                labels=["Processados", "Não Processados"],
                values=[proc_ins, np_ins],
                hole=0.4,
                marker_colors=[COR_AZUL, COR_VERDE]
            ))
            fig_pizza.update_layout(
                height=350, margin=dict(l=0, r=0, t=30, b=0),
                title_text="Composição do Inscrito"
            )
            col_g2.plotly_chart(fig_pizza, use_container_width=True)

            st.divider()

            # ---- Tabela detalhada ------------------------------------------
            def _grp_rp(df_src, grp_col):
                """Agrega colunas de RP por coluna de agrupamento, separando A Pagar / A Liquidar."""
                return (df_src
                    .assign(
                        inscrito   = df_src["proc_inscrito"] + df_src["np_inscrito"],
                        pagos      = df_src["proc_pagos"]    + df_src["np_pagos"],
                        cancelados = df_src["proc_cancelados"] + df_src["np_cancelados"],
                        a_pagar    = df_src["proc_a_pagar"]  + df_src["np_a_pagar"],
                        em_liq     = df_src["np_em_liquidacao"],
                        a_liquidar = df_src["np_a_liquidar"],
                        np_ins_g   = df_src["np_inscrito"],
                        np_can_g   = df_src["np_cancelados"],
                    )
                    .groupby(grp_col, as_index=False)
                    .agg({
                        "inscrito": "sum", "pagos": "sum", "cancelados": "sum",
                        "a_pagar": "sum", "em_liq": "sum", "a_liquidar": "sum",
                        "np_ins_g": "sum", "np_can_g": "sum",
                    })
                    .assign(np_liquidados=lambda d: (d["np_ins_g"] - d["a_liquidar"] - d["np_can_g"]).clip(lower=0))
                    .drop(columns=["np_ins_g", "np_can_g"])
                    .sort_values("inscrito", ascending=False)
                )

            def _fmt_grp(df_t, label_col, label_rename):
                pct_tot = df_t["inscrito"].sum()
                df_t["pct"] = df_t["inscrito"].apply(
                    lambda v: "{:.1f}%".format(v / pct_tot * 100) if pct_tot > 0 else "0.0%")
                for c in ["inscrito", "pagos", "cancelados", "a_pagar", "em_liq", "a_liquidar", "np_liquidados"]:
                    df_t[c] = df_t[c].apply(lambda v: "R$ {:,.2f}".format(v))
                df_t.rename(columns={
                    label_col: label_rename,
                    "inscrito": "Inscrito", "pagos": "Pagos", "cancelados": "Cancelados",
                    "a_pagar": "A Pagar (liq.)", "em_liq": "Em Liquidação",
                    "a_liquidar": "NP A Liquidar", "np_liquidados": "NP Liquidados",
                    "pct": "% Inscrito",
                }, inplace=True)
                return df_t

            if tipo_rp == "Por Credor":
                st.markdown("#### Por Credor")
                df_cred_tab = _grp_rp(df_rp_ref, "nome_credor")
                st.dataframe(_fmt_grp(df_cred_tab, "nome_credor", "Credor"),
                             use_container_width=True, hide_index=True)

            elif tipo_rp == "Por Dotação":
                st.markdown("#### Por Dotação Orçamentária")
                df_dot_tab = _grp_rp(df_rp_ref, "dotacao")
                st.dataframe(_fmt_grp(df_dot_tab, "dotacao", "Dotação"),
                             use_container_width=True, hide_index=True)

            else:
                st.markdown("#### Resumo Geral")
                fv = lambda v: "R$ {:,.2f}".format(v)
                resumo = pd.DataFrame([
                    {
                        "Tipo": "Processados",
                        "Inscrito":      fv(proc_ins),
                        "Pagos":         fv(proc_pag),
                        "Cancelados":    fv(proc_can),
                        "A Pagar":       fv(proc_apa),   # liquidados aguardando pagamento
                        "Em Liquidação": "—",
                        "A Liquidar":    "—",
                        "NP Liquidados": "—",            # todos já são liquidados por definição
                    },
                    {
                        "Tipo": "Não Processados",
                        "Inscrito":      fv(np_ins),
                        "Pagos":         fv(np_pag),
                        "Cancelados":    fv(np_can),
                        "A Pagar":       fv(np_apa),     # liquidados aguardando pagamento
                        "Em Liquidação": fv(np_eliq),
                        "A Liquidar":    fv(np_aliq),
                        "NP Liquidados": fv(np_liquidados),
                    },
                    {
                        "Tipo": "TOTAL",
                        "Inscrito":      fv(tot_inscrito),
                        "Pagos":         fv(tot_pagos),
                        "Cancelados":    fv(tot_cancelados),
                        "A Pagar":       fv(proc_apa + np_apa),
                        "Em Liquidação": fv(np_eliq),
                        "A Liquidar":    fv(np_aliq),
                        "NP Liquidados": fv(np_liquidados),
                    },
                ])
                st.dataframe(resumo, use_container_width=True, hide_index=True)
                st.caption(
                    "**A Pagar** = liquidados aguardando pagamento (Proc. + NP).  "
                    "**NP Liquidados** = NP inscrito − NP A Liquidar − Cancelados."
                )

            st.divider()

            # ---- Tabela de empenhos ----------------------------------------
            with st.expander("Ver todos os empenhos"):
                df_det = df_rp_ref[["num_empenho", "nome_credor", "dotacao",
                                    "proc_inscrito", "np_inscrito",
                                    "proc_pagos", "np_pagos",
                                    "proc_cancelados", "np_cancelados",
                                    "proc_a_pagar", "np_a_pagar",
                                    "np_em_liquidacao", "np_a_liquidar"]].copy()
                df_det["Inscrito"]   = df_det["proc_inscrito"] + df_det["np_inscrito"]
                df_det["Pagos"]      = df_det["proc_pagos"]    + df_det["np_pagos"]
                df_det["Cancelados"] = df_det["proc_cancelados"] + df_det["np_cancelados"]
                df_det["A Pagar"]    = df_det["proc_a_pagar"]  + df_det["np_a_pagar"]
                df_det["Em Liq."]    = df_det["np_em_liquidacao"]
                df_det["A Liquidar"] = df_det["np_a_liquidar"]
                df_det["NP Liquidados"] = (
                    df_det["np_inscrito"] - df_det["np_a_liquidar"] - df_det["np_cancelados"]
                ).clip(lower=0)
                df_det = df_det[["num_empenho", "nome_credor", "dotacao",
                                 "Inscrito", "Pagos", "Cancelados",
                                 "A Pagar", "Em Liq.", "A Liquidar", "NP Liquidados"]]
                df_det.rename(columns={
                    "num_empenho": "Empenho", "nome_credor": "Credor", "dotacao": "Dotação"
                }, inplace=True)
                df_det = df_det.sort_values("Inscrito", ascending=False)
                for col in ["Inscrito", "Pagos", "Cancelados", "A Pagar",
                            "Em Liq.", "A Liquidar", "NP Liquidados"]:
                    df_det[col] = df_det[col].apply(lambda v: "R$ {:,.2f}".format(v))
                st.dataframe(df_det, use_container_width=True, hide_index=True)

            # ---- Exportar PDF ----------------------------------------------
            try:
                pdf_rp = gerar_pdf_rp(df_rp_ref, mes_ref_rp, ms_rp)
                st.download_button(
                    "📄 Exportar Relatório PDF — Restos a Pagar",
                    data=pdf_rp,
                    file_name="relatorio_restos_a_pagar.pdf",
                    mime="application/pdf",
                    key="pdf_rp"
                )
            except Exception as e:
                st.warning("PDF indisponível: " + str(e))
