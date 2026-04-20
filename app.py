import streamlit as st
import pandas as pd
import sqlite3
import plotly.graph_objects as go
import unicodedata
import re
import io

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
    v_sup = {**v_zero, "emp_no_bim": max(receita_bim - v_sub["emp_no_bim"], 0), "emp_ate": max(receita_ate - v_sub["emp_ate"], 0), "liq_no_bim": max(receita_bim - v_sub["liq_no_bim"], 0), "liq_ate": max(receita_ate - v_sub["liq_ate"], 0), "pago_ate": max(receita_ate - v_sub["pago_ate"], 0)}

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
# BANCO DE DADOS
# ---------------------------------------------------------------------------
def inicializar_banco():
    conn = sqlite3.connect(DB_NAME)

    conn.execute(
        "CREATE TABLE IF NOT EXISTS receitas ("
        "mes INTEGER, ano INTEGER, codigo_full TEXT, natureza TEXT, "
        "orcado REAL, realizado REAL, previsao REAL, "
        "categoria TEXT DEFAULT 'Nao Classificada')"
    )
    try:
        conn.execute(
            "ALTER TABLE receitas ADD COLUMN categoria TEXT DEFAULT 'Nao Classificada'"
        )
    except Exception:
        pass

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

    cols_sub = [r[1] for r in conn.execute("PRAGMA table_info(sub_elementos)").fetchall()]
    schema_correto = (
        cols_sub == ["mes", "ano", "paoe", "natureza_cod", "natureza_desc",
                     "subelemento_cod", "subelemento_desc", "fonte", "liquidado", "pago"]
    )
    if not schema_correto:
        conn.execute("DROP TABLE IF EXISTS sub_elementos")
    conn.execute(
        "CREATE TABLE IF NOT EXISTS sub_elementos ("
        "mes INTEGER, ano INTEGER, paoe TEXT, natureza_cod TEXT, natureza_desc TEXT, "
        "subelemento_cod TEXT, subelemento_desc TEXT, fonte TEXT, "
        "liquidado REAL, pago REAL)"
    )

    # Tabela para repasses recebidos (ANEXO V)
    conn.execute(
        "CREATE TABLE IF NOT EXISTS anexo_v ("
        "mes INTEGER, ano INTEGER, data TEXT, "
        "entidade_repassadora TEXT, valor REAL, "
        "finalidade TEXT, fundamento_legal TEXT)"
    )

    conn.execute(
        "UPDATE receitas SET categoria='Receita Corrente' "
        "WHERE categoria='Repasses Correntes'"
    )
    conn.commit()
    conn.close()


def limpar_todos_dados():
    conn = sqlite3.connect(DB_NAME)
    conn.execute("DELETE FROM receitas")
    conn.execute("DELETE FROM orcamento")
    conn.execute("DELETE FROM execucao")
    conn.execute("DELETE FROM sub_elementos")
    conn.execute("DELETE FROM anexo_v")
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
            "Repasses Recebidos (ANEXO V)"
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
                        if not re.match(r"^\d", cod) or cod[-1] == "0":
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
                    "INSERT INTO receitas VALUES (?,?,?,?,?,?,?,?)", dados
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
                        if not uo or uo in ("nan", ""):
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

                conn.execute(
                    "DELETE FROM orcamento WHERE ano=2026 AND mes=?", (m_final,)
                )
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
                cur_paoe = ""
                cur_nat_cod = ""
                cur_nat_desc = ""

                for i, row in df701.iterrows():
                    text = str(row.iloc[0]).strip().replace("\xa0", " ")
                    if i < 8 or not text or text == "nan":
                        continue
                    tu = sem_acento(text).upper()

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
                    chaves_701 = ["paoe", "nat_cod", "sub_cod"]
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
                            r.paoe, r.nat_cod, r.nat_desc,
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
                        "(mes, ano, paoe, natureza_cod, natureza_desc, "
                        "subelemento_cod, subelemento_desc, fonte, liquidado, pago) "
                        "VALUES (?,?,?,?,?,?,?,?,?,?)",
                        dados
                    )
                    conn.commit()
                    liq_t = sum(r[8] for r in dados)
                    pag_t = sum(r[9] for r in dados)
                    st.success(
                        "Sub-elemento " + MESES_NOMES[m_final - 1] + "/2026: "
                        + str(len(dados)) + " registros | "
                        + "Liq R$ {:,.0f} | Pago R$ {:,.0f}".format(liq_t, pag_t)
                    )

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

            conn.close()

        except Exception as e:
            st.error("Erro: " + str(e))
            import traceback
            st.code(traceback.format_exc())

    st.divider()
    st.subheader("Backup Completo")
    conn_b = sqlite3.connect(DB_NAME)
    tbls = {
        "receitas":      pd.read_sql("SELECT * FROM receitas",      conn_b),
        "orcamento":     pd.read_sql("SELECT * FROM orcamento",     conn_b),
        "execucao":      pd.read_sql("SELECT * FROM execucao",      conn_b),
        "sub_elementos": pd.read_sql("SELECT * FROM sub_elementos", conn_b),
        "anexo_v":       pd.read_sql("SELECT * FROM anexo_v",       conn_b),
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
        ["receitas", "orcamento", "execucao", "sub_elementos", "anexo_v"],
        key="tabela_rest"
    )
    file_restore = st.file_uploader("Arquivo CSV", type=["csv"], key="file_rest")
    if file_restore and st.button("Restaurar"):
        df_res = pd.read_csv(file_restore)
        conn_r = sqlite3.connect(DB_NAME)
        df_res.to_sql(tabela_rest, conn_r, if_exists="replace", index=False)
        conn_r.commit()
        conn_r.close()
        st.success("Tabela '" + tabela_rest + "' restaurada!")
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
df_rec    = pd.read_sql("SELECT * FROM receitas",      conn_main)
df_orc    = pd.read_sql("SELECT * FROM orcamento",     conn_main)
df_exec   = pd.read_sql("SELECT * FROM execucao",      conn_main)
df_sub    = pd.read_sql("SELECT * FROM sub_elementos", conn_main)
df_anexov = pd.read_sql("SELECT * FROM anexo_v",       conn_main)
conn_main.close()

tab1, tab2, tab3, tab4 = st.tabs(["Receitas", "Despesas", "Comparativo", "Relatorios LRF"])


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
            v_real = df_rf["realizado"].sum()
            v_orc  = (
                df_rec[df_rec["mes"] == max(ms_r)]
                .groupby("codigo_full")["orcado"].max().sum()
            )
            k1, k2, k3 = st.columns(3)
            k1.metric("Orcado Atual",  "R$ {:,.2f}".format(v_orc))
            k2.metric("Realizado",     "R$ {:,.2f}".format(v_real))
            k3.metric("Atingimento",
                      "{:.1f}%".format(v_real / v_orc * 100 if v_orc != 0 else 0))

            df_g = df_rf.groupby("mes").agg({"realizado": "sum"}).reset_index()
            df_g["previsao"] = [
                df_rf[df_rf["mes"] == m].groupby("codigo_full")["previsao"].max().sum()
                for m in df_g["mes"]
            ]
            fig = go.Figure()
            fig.add_trace(go.Bar(
                x=[MESES_NOMES[m - 1] for m in df_g["mes"]],
                y=df_g["realizado"], name="Realizado", marker_color="#2E7D32"
            ))
            fig.add_trace(go.Scatter(
                x=[MESES_NOMES[m - 1] for m in df_g["mes"]],
                y=df_g["previsao"], name="Previsao",
                line=dict(color="#FF9800", width=3, dash="dot")
            ))
            fig.update_layout(
                height=350, margin=dict(l=0, r=0, t=30, b=0),
                hovermode="x unified",
                legend=dict(orientation="h", yanchor="bottom", y=1.02,
                            xanchor="right", x=1)
            )
            st.plotly_chart(fig, width='stretch')
            st.dataframe(
                df_rf[["categoria", "codigo_full", "natureza", "realizado", "orcado"]]
                .style.format({"realizado": "{:,.2f}", "orcado": "{:,.2f}"}),
                width='stretch'
            )

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

        emp_total = df_ef["empenhado"].sum() if not df_ef.empty else 0
        liq_total = df_ef["liquidado"].sum() if not df_ef.empty else 0
        pag_total = df_ef["pago"].sum()      if not df_ef.empty else 0

        k1, k2, k3, k4 = st.columns(4)
        k1.metric(
            "Cred. Autorizado (mes " + MESES_NOMES[m_max_orc - 1] + ")",
            "R$ {:,.2f}".format(cred_total)
        )
        k2.metric("Empenhado",  "R$ {:,.2f}".format(emp_total))
        k3.metric("Liquidado",  "R$ {:,.2f}".format(liq_total))
        k4.metric("Pago",       "R$ {:,.2f}".format(pag_total))

        if not df_ef.empty:
            df_g = (
                df_ef.groupby("mes")[["empenhado", "liquidado", "pago"]]
                .sum().reset_index()
            )
            fig = go.Figure()
            fig.add_trace(go.Bar(
                x=[MESES_NOMES[m - 1] for m in df_g["mes"]],
                y=df_g["empenhado"], name="Empenhado", marker_color="#1565C0"
            ))
            fig.add_trace(go.Bar(
                x=[MESES_NOMES[m - 1] for m in df_g["mes"]],
                y=df_g["liquidado"], name="Liquidado", marker_color="#2E7D32"
            ))
            fig.add_trace(go.Bar(
                x=[MESES_NOMES[m - 1] for m in df_g["mes"]],
                y=df_g["pago"], name="Pago", marker_color="#E65100"
            ))
            fig.update_layout(
                height=320, barmode="group",
                margin=dict(l=0, r=0, t=30, b=0),
                hovermode="x unified",
                legend=dict(orientation="h", yanchor="bottom", y=1.02,
                            xanchor="right", x=1)
            )
            st.plotly_chart(fig, width='stretch')

            ug_filtrada = set(ug_sel) != set(ugs_disp)
            col_chave = (["ug"] if ug_filtrada else []) + [
                "funcao", "subfuncao", "programa", "projeto", "fonte", "natureza"
            ]
            df_agg = df_ef.groupby(col_chave, as_index=False)[
                ["empenhado", "liquidado", "pago"]
            ].sum()

            st.dataframe(
                df_agg[col_chave + ["empenhado", "liquidado", "pago"]]
                .style.format({
                    "empenhado": "{:,.2f}",
                    "liquidado": "{:,.2f}",
                    "pago": "{:,.2f}"
                }),
                width='stretch'
            )

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
                    sorted(df_exec["ug"].dropna().unique().tolist())
                    if not df_exec.empty else []
                )
                ug_sel_s = fs6.multiselect(
                    "UG (via Natureza):", ugs_sub, key="ug_sub"
                )

                df_sf = df_sub[df_sub["mes"].isin(ms_s)].copy()
                if ug_sel_s and not df_exec.empty:
                    nats_ug = df_exec[df_exec["ug"].isin(ug_sel_s)]["natureza"].dropna().unique()
                    nats_cod_ug = set(
                        re.match(r"^(\d+)", str(n).strip()).group(1)
                        for n in nats_ug
                        if re.match(r"^(\d+)", str(n).strip())
                    )
                    df_sf = df_sf[
                        df_sf["natureza_cod"].apply(
                            lambda x: re.match(r"^(\d+)", str(x).strip()).group(1)
                            if re.match(r"^(\d+)", str(x).strip()) else x
                        ).isin(nats_cod_ug)
                    ]
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
                    col_s = ["paoe", "natureza_cod", "natureza_desc"]
                    if has_fonte:
                        col_s += ["fonte"]
                    col_s += ["subelemento_cod", "subelemento_desc"]

                    df_sv = df_sf.groupby(col_s, as_index=False)[
                        ["liquidado", "pago"]
                    ].sum()

                    ks1, ks2 = st.columns(2)
                    ks1.metric(
                        "Liquidado", "R$ {:,.2f}".format(df_sv["liquidado"].sum())
                    )
                    ks2.metric(
                        "Pago", "R$ {:,.2f}".format(df_sv["pago"].sum())
                    )
                    st.dataframe(
                        df_sv[col_s + ["liquidado", "pago"]]
                        .style.format({
                            "liquidado": "{:,.2f}",
                            "pago": "{:,.2f}"
                        }),
                        width='stretch'
                    )
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
