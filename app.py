import os
import re

import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Equipamentos - BSS Gongo Soco", layout="wide", page_icon="🚜")

DEFAULT_PATH = r"C:\CLAUDE\Relação de equipamentos Filial 010148  Obra BSS GONGO SOCO.xlsm"
SHEET = "RELAÇÃO EMS"

COMB_PATH = r"C:\CLAUDE\combustivel.xlsx"
COMB_SHEET = "Planilha1"
COMB_COLS = ["Frota", "Descricao", "Dt. Abastec.", "Contador 1", "Quantidade", "Placa"]

PERIODOS = {"Diário": "D", "Semanal": "W", "Mensal": "M"}

# Tipos cujo "Contador 1" é odômetro (KM rodado) em vez de horímetro (horas trabalhadas).
KM_TYPES = {"CAMINHONETE", "ONIBUS", "ONIBUS MPOLO", "VAN", "VAN SPRINTER M"}

TEXT_COLS = ["CÓDIGO", "CÓDIGO CLIENTE", "MOBILIZAÇÃO", "TIPO", "MARCA", "MODELO",
             "PLACA", "SÉRIE", "CAPAC.", "SITUAÇÃO", "EMPRESA"]

DISPLAY_COLS = ["CÓDIGO", "CÓDIGO CLIENTE", "MOBILIZAÇÃO", "TIPO", "MARCA", "MODELO",
                "PLACA", "SÉRIE", "ANO", "CAPAC.", "SITUAÇÃO", "EMPRESA",
                "DATA CHEGADA NA OBRA", "DATA LIBERAÇÃO CLIENTE"]

DOC_COLS = ["CIV / CIPP", "CRONOTACOGRAFO", "LIQUIDO PENETRANTE",
            "HIGIENIZAÇÃO TANQUE - AGUA POTÁVEL", "LAUDO VASOS DE PRESSÃO",
            "FUMAÇA PRETA", "SELO DE LIBERAÇÃO", "CHECK LIST VALE"]

ID_COLS = ["CÓDIGO", "CÓDIGO CLIENTE", "TIPO", "MARCA", "MODELO", "PLACA", "EMPRESA", "MOBILIZAÇÃO"]


def build_vencimentos(df, doc_cols, dias_alerta):
    hoje = pd.Timestamp.now().normalize()
    partes = []
    for col in doc_cols:
        datas = df[col + "__DATA"]
        validas = datas.notna()
        if not validas.any():
            continue
        parte = df.loc[validas, ID_COLS].copy()
        parte["DOCUMENTO"] = col
        parte["VENCIMENTO"] = datas[validas]
        partes.append(parte)

    if not partes:
        return pd.DataFrame(columns=ID_COLS + ["DOCUMENTO", "VENCIMENTO", "DIAS_RESTANTES", "STATUS"])

    venc = pd.concat(partes, ignore_index=True)
    venc["DIAS_RESTANTES"] = (venc["VENCIMENTO"] - hoje).dt.days

    def status(d):
        if d < 0:
            return "🔴 VENCIDO"
        if d <= dias_alerta:
            return "🟡 VENCE EM BREVE"
        return "🟢 OK"

    venc["STATUS"] = venc["DIAS_RESTANTES"].apply(status)
    return venc.sort_values("DIAS_RESTANTES")


DATE_DISPLAY_COLS = ["DATA CHEGADA NA OBRA", "DATA LIBERAÇÃO CLIENTE"]


@st.cache_data(ttl=300)
def load_data(source):
    df = pd.read_excel(source, sheet_name=SHEET, header=7, engine="openpyxl")
    df = df.dropna(subset=["CÓDIGO"]).copy()
    for col in DOC_COLS:
        df[col + "__DATA"] = pd.to_datetime(df[col], errors="coerce")
    for col in TEXT_COLS:
        df[col] = df[col].astype(str).str.strip().replace({"nan": "-", "None": "-"})
    for col in DATE_DISPLAY_COLS:
        df[col] = pd.to_datetime(df[col], errors="coerce").dt.strftime("%d/%m/%Y").fillna("-")
    return df


def normaliza_frota(frota):
    m = re.match(r"^Z148(\d+)$", frota)
    if m:
        return f"CBM-{m.group(1)}"
    return frota


@st.cache_data(ttl=600)
def load_combustivel(path, ems_df):
    comb = pd.read_excel(path, sheet_name=COMB_SHEET, header=3, engine="openpyxl", usecols=COMB_COLS)
    comb = comb.dropna(subset=["Frota"]).copy()
    comb = comb[comb["Frota"].astype(str) != "Frota"].copy()
    comb["Frota"] = comb["Frota"].astype(str).str.strip()
    comb["Contador 1"] = pd.to_numeric(comb["Contador 1"], errors="coerce")
    comb["Dt. Abastec."] = pd.to_datetime(comb["Dt. Abastec."], errors="coerce")
    comb = comb.dropna(subset=["Contador 1", "Dt. Abastec."])
    comb["EQUIPAMENTO"] = comb["Frota"].map(normaliza_frota)
    comb["UNIDADE"] = comb["Descricao"].apply(lambda d: "km" if d in KM_TYPES else "h")

    ems_info = ems_df[["CÓDIGO", "TIPO", "MARCA", "EMPRESA"]].drop_duplicates(subset="CÓDIGO")
    comb = comb.merge(ems_info, left_on="EQUIPAMENTO", right_on="CÓDIGO", how="left")
    comb["TIPO"] = comb["TIPO"].fillna(comb["Descricao"])
    comb["MARCA"] = comb["MARCA"].fillna("-")
    comb["EMPRESA"] = comb["EMPRESA"].fillna("CBM (frota própria)")
    return comb


GRUPO_COLS = ["EQUIPAMENTO", "TIPO", "MARCA", "EMPRESA"]


def build_leitura_diaria(comb):
    """Reduz a um horímetro por dia por equipamento (o maior valor lido no dia).
    A leitura anterior de cada dia é sempre a leitura (máxima) do dia anterior com
    registro — é essa cadeia dia a dia que alimenta o cálculo de horas trabalhadas."""
    d = comb.copy()
    d["DATA"] = d["Dt. Abastec."].dt.normalize()
    diaria = (
        d.groupby(GRUPO_COLS + ["DATA"], dropna=False)["Contador 1"]
        .max().reset_index().rename(columns={"Contador 1": "LEITURA"})
    )
    diaria = diaria.sort_values(GRUPO_COLS + ["DATA"])
    diaria["LEITURA_ANTERIOR"] = diaria.groupby("EQUIPAMENTO")["LEITURA"].shift(1)
    diaria["DELTA"] = diaria["LEITURA"] - diaria["LEITURA_ANTERIOR"]
    diaria.loc[diaria["DELTA"] < 0, "DELTA"] = None
    return diaria


def calc_delta_por_periodo(comb, freq, coluna_saida):
    """Agrupa a série diária de leituras (build_leitura_diaria) por período. Usa TODO
    o histórico disponível do equipamento (não apenas o intervalo de datas exibido),
    para que o período inicial de um filtro de data ainda tenha uma "leitura anterior"
    válida. Filtre o resultado por data DEPOIS de chamar esta função (coluna PERIODO_INICIO)."""
    diaria = build_leitura_diaria(comb)
    diaria["PERIODO"] = diaria["DATA"].dt.to_period(freq)
    diaria = diaria.sort_values(GRUPO_COLS + ["DATA"])

    chaves = GRUPO_COLS + ["PERIODO"]
    agg = diaria.groupby(chaves, dropna=False).agg(
        LEITURA_INICIAL=("LEITURA_ANTERIOR", "first"),
        LEITURA_FINAL=("LEITURA", "last"),
        **{coluna_saida: ("DELTA", "sum")},
    ).reset_index()
    tem_dado = diaria.groupby(chaves, dropna=False)["DELTA"].apply(lambda s: s.notna().any())
    agg = agg.merge(tem_dado.rename("_TEM_DADO"), on=chaves)
    agg.loc[~agg["_TEM_DADO"], coluna_saida] = None
    agg = agg.drop(columns="_TEM_DADO")

    agg["PERIODO_INICIO"] = agg["PERIODO"].dt.start_time
    agg["PERIODO"] = agg["PERIODO"].astype(str)
    return agg.sort_values(["EQUIPAMENTO", "PERIODO_INICIO"])


st.title("🚜 Relação de Equipamentos Móveis de Superfície")
st.caption("Filial 010148 · Obra BSS · Gongo Soco / Barão de Cocais")

with st.sidebar:
    st.header("Fonte de dados")
    uploaded = st.file_uploader("Substituir planilha (opcional)", type=["xlsm", "xlsx"])
    if st.button("🔄 Recarregar dados da planilha"):
        st.cache_data.clear()

source = uploaded if uploaded is not None else DEFAULT_PATH

if uploaded is None and not os.path.exists(DEFAULT_PATH):
    st.info("📤 Envie a planilha de equipamentos (.xlsm/.xlsx) na barra lateral para começar.")
    st.stop()

try:
    df = load_data(source)
except Exception as e:
    st.error(f"Erro ao ler a planilha: {e}")
    st.stop()

with st.sidebar:
    st.header("Filtros")
    f_status = st.multiselect("Mobilização", sorted(df["MOBILIZAÇÃO"].unique()))
    f_tipo = st.multiselect("Tipo de equipamento", sorted(df["TIPO"].unique()))
    f_marca = st.multiselect("Marca", sorted(df["MARCA"].unique()))
    f_empresa = st.multiselect("Empresa / Locadora", sorted(df["EMPRESA"].unique()))
    f_situacao = st.multiselect("Situação", sorted(df["SITUAÇÃO"].unique()))
    busca = st.text_input("Buscar (código, placa, série, código cliente)")
    st.divider()
    dias_alerta = st.slider("Alertar vencimento em até (dias)", 0, 90, 30)

dff = df.copy()
if f_status:
    dff = dff[dff["MOBILIZAÇÃO"].isin(f_status)]
if f_tipo:
    dff = dff[dff["TIPO"].isin(f_tipo)]
if f_marca:
    dff = dff[dff["MARCA"].isin(f_marca)]
if f_empresa:
    dff = dff[dff["EMPRESA"].isin(f_empresa)]
if f_situacao:
    dff = dff[dff["SITUAÇÃO"].isin(f_situacao)]
if busca:
    campos = ["CÓDIGO", "CÓDIGO CLIENTE", "PLACA", "SÉRIE"]
    mask = dff[campos].apply(lambda c: c.str.contains(busca, case=False, na=False)).any(axis=1)
    dff = dff[mask]

venc_all = build_vencimentos(df, DOC_COLS, dias_alerta)
n_vencidos = int((venc_all["STATUS"] == "🔴 VENCIDO").sum())
n_vence_breve = int((venc_all["STATUS"] == "🟡 VENCE EM BREVE").sum())

if n_vencidos or n_vence_breve:
    st.warning(
        f"⚠️ {n_vencidos} documento(s) **vencido(s)** e {n_vence_breve} **vencendo em até {dias_alerta} dias**. "
        "Veja a aba **🔔 Vencimento de Documentos**."
    )

tab_geral, tab_venc, tab_horas = st.tabs(
    ["📊 Visão Geral", "🔔 Vencimento de Documentos", "⛽ Horas Trabalhadas"]
)

with tab_geral:
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total de Equipamentos", len(dff))
    c2.metric("Mobilizados", int((dff["MOBILIZAÇÃO"] == "MOBILIZADO").sum()))
    c3.metric("Desmobilizados", int((dff["MOBILIZAÇÃO"] == "DESMOBILIZADO").sum()))
    c4.metric("Ag. Mobilização", int((dff["MOBILIZAÇÃO"] == "AG. MOBILIZAÇÃO").sum()))

    st.divider()

    col1, col2 = st.columns(2)
    with col1:
        tipo_counts = dff["TIPO"].value_counts().reset_index()
        tipo_counts.columns = ["TIPO", "QTD"]
        fig = px.bar(tipo_counts, x="QTD", y="TIPO", orientation="h", title="Equipamentos por Tipo")
        fig.update_layout(yaxis={"categoryorder": "total ascending"}, height=550)
        st.plotly_chart(fig, width='stretch')

    with col2:
        mob_counts = dff["MOBILIZAÇÃO"].value_counts().reset_index()
        mob_counts.columns = ["STATUS", "QTD"]
        fig2 = px.pie(mob_counts, names="STATUS", values="QTD", title="Status de Mobilização", hole=0.45)
        st.plotly_chart(fig2, width='stretch')

        empresa_counts = dff["EMPRESA"].value_counts().reset_index().head(15)
        empresa_counts.columns = ["EMPRESA", "QTD"]
        fig4 = px.bar(empresa_counts, x="QTD", y="EMPRESA", orientation="h", title="Top 15 Empresas / Locadoras")
        fig4.update_layout(yaxis={"categoryorder": "total ascending"})
        st.plotly_chart(fig4, width='stretch')

    st.divider()
    st.subheader(f"Lista de Equipamentos ({len(dff)})")
    st.dataframe(dff[DISPLAY_COLS], width='stretch', hide_index=True)

    csv = dff[DISPLAY_COLS].to_csv(index=False).encode("utf-8-sig")
    st.download_button("⬇️ Baixar CSV filtrado", csv, "equipamentos_filtrado.csv", "text/csv")

with tab_venc:
    st.subheader("Vencimento de Documentos / Certificações")
    st.caption(
        "Considera as colunas CIV/CIPP, CRONOTACOGRAFO, LIQUIDO PENETRANTE, HIGIENIZAÇÃO TANQUE, "
        "LAUDO VASOS DE PRESSÃO, FUMAÇA PRETA, SELO DE LIBERAÇÃO e CHECK LIST VALE. "
        "Células que hoje têm apenas texto de status (ex.: \"OK\", \"-\") são ignoradas até serem preenchidas com uma data real."
    )

    codigos_filtrados = set(dff["CÓDIGO"])
    venc = venc_all[venc_all["CÓDIGO"].isin(codigos_filtrados)].copy()

    vcol1, vcol2, vcol3 = st.columns(3)
    vcol1.metric("🔴 Vencidos", int((venc["STATUS"] == "🔴 VENCIDO").sum()))
    vcol2.metric("🟡 Vencendo em breve", int((venc["STATUS"] == "🟡 VENCE EM BREVE").sum()))
    vcol3.metric("🟢 Em dia", int((venc["STATUS"] == "🟢 OK").sum()))

    doc_filter = st.multiselect("Filtrar por documento", DOC_COLS)
    status_filter = st.multiselect(
        "Filtrar por status",
        ["🔴 VENCIDO", "🟡 VENCE EM BREVE", "🟢 OK"],
        default=["🔴 VENCIDO", "🟡 VENCE EM BREVE"],
    )

    venc_f = venc.copy()
    if doc_filter:
        venc_f = venc_f[venc_f["DOCUMENTO"].isin(doc_filter)]
    if status_filter:
        venc_f = venc_f[venc_f["STATUS"].isin(status_filter)]

    if venc_f.empty:
        st.info("Nenhum documento com data cadastrada corresponde aos filtros selecionados.")
    else:
        cols_venc = ["STATUS", "DOCUMENTO", "CÓDIGO", "CÓDIGO CLIENTE", "TIPO", "MARCA",
                     "PLACA", "EMPRESA", "VENCIMENTO", "DIAS_RESTANTES"]
        st.dataframe(venc_f[cols_venc], width='stretch', hide_index=True)

        resumo = venc[venc["STATUS"] != "🟢 OK"].groupby(["DOCUMENTO", "STATUS"]).size().reset_index(name="QTD")
        if not resumo.empty:
            fig5 = px.bar(
                resumo, x="DOCUMENTO", y="QTD", color="STATUS", barmode="stack",
                title="Documentos vencidos / a vencer por tipo",
                color_discrete_map={"🔴 VENCIDO": "#e74c3c", "🟡 VENCE EM BREVE": "#f1c40f"},
            )
            st.plotly_chart(fig5, width='stretch')

        csv_venc = venc_f[cols_venc].to_csv(index=False).encode("utf-8-sig")
        st.download_button("⬇️ Baixar alertas em CSV", csv_venc, "alertas_vencimento.csv", "text/csv")

with tab_horas:
    st.subheader("Horas Trabalhadas por Período (via Combustível)")
    st.caption(
        "Calcula horas trabalhadas pela diferença entre leituras do horímetro acumulado (\"Contador 1\") "
        "registradas a cada abastecimento. Equipamentos com Frota no padrão \"Z148XXX\" são exibidos como "
        "\"CBM-XXX\" (frota própria da CBM). Veículos rodoviários (caminhonete, ônibus, van) usam "
        "odômetro (KM) em vez de horímetro e são calculados à parte, na seção de quilometragem abaixo."
    )

    comb_uploaded = st.file_uploader("Substituir planilha de combustível (opcional)", type=["xlsx"], key="comb_upload")
    comb_source = comb_uploaded if comb_uploaded is not None else COMB_PATH

    comb = None
    if comb_uploaded is None and not os.path.exists(COMB_PATH):
        st.info("📤 Envie a planilha de combustível (.xlsx) acima para calcular as horas trabalhadas.")
    else:
        try:
            comb = load_combustivel(comb_source, df)
        except Exception as e:
            st.error(f"Erro ao ler a planilha de combustível: {e}")

    if comb is not None:
        hcol1, hcol2, hcol3 = st.columns([1, 1, 2])
        with hcol1:
            periodo_label = st.selectbox("Agrupar por período", list(PERIODOS.keys()), index=2)
        with hcol2:
            data_min, data_max = comb["Dt. Abastec."].min().date(), comb["Dt. Abastec."].max().date()
            intervalo = st.date_input("Intervalo de datas", value=(data_min, data_max),
                                       min_value=data_min, max_value=data_max)
        with hcol3:
            equip_opcoes = sorted(comb.loc[comb["UNIDADE"] == "h", "EQUIPAMENTO"].unique())
            equip_sel = st.multiselect("Filtrar equipamento(s) (vazio = todos)", equip_opcoes)

        ini, fim = None, None
        if isinstance(intervalo, tuple) and len(intervalo) == 2:
            ini, fim = pd.Timestamp(intervalo[0]), pd.Timestamp(intervalo[1])

        # Calcula sobre o histórico completo do(s) equipamento(s) selecionado(s) — sem
        # cortar por data ainda — para que o primeiro período do filtro tenha uma leitura
        # anterior válida. O filtro de data é aplicado só depois, na exibição.
        comb_equip = comb[comb["UNIDADE"] == "h"]
        if equip_sel:
            comb_equip = comb_equip[comb_equip["EQUIPAMENTO"].isin(equip_sel)]

        horas_completo = calc_delta_por_periodo(comb_equip, PERIODOS[periodo_label], "HORAS_TRABALHADAS")
        horas = horas_completo
        if ini is not None:
            horas = horas[(horas["PERIODO_INICIO"] >= ini.to_period(PERIODOS[periodo_label]).start_time) &
                           (horas["PERIODO_INICIO"] <= fim)]
        horas_validas = horas.dropna(subset=["HORAS_TRABALHADAS"])

        comb_f = comb_equip
        if ini is not None:
            comb_f = comb_f[(comb_f["Dt. Abastec."] >= ini) & (comb_f["Dt. Abastec."] <= fim)]

        kcol1, kcol2, kcol3 = st.columns(3)
        kcol1.metric("Total de Horas Trabalhadas", f"{horas_validas['HORAS_TRABALHADAS'].sum():,.0f} h")
        kcol2.metric("Equipamentos no filtro", comb_f["EQUIPAMENTO"].nunique())
        kcol3.metric("Registros de abastecimento", len(comb_f))

        st.divider()

        ranking = (
            horas_validas.groupby(["EQUIPAMENTO", "TIPO"])["HORAS_TRABALHADAS"]
            .sum().reset_index().sort_values("HORAS_TRABALHADAS", ascending=False).head(20)
        )
        if not ranking.empty:
            fig6 = px.bar(
                ranking, x="HORAS_TRABALHADAS", y="EQUIPAMENTO", orientation="h", color="TIPO",
                title="Top 20 Equipamentos por Horas Trabalhadas (no período filtrado)",
            )
            fig6.update_layout(yaxis={"categoryorder": "total ascending"}, height=600)
            st.plotly_chart(fig6, width='stretch')

        st.subheader(f"Detalhamento por Equipamento e Período ({len(horas)})")
        cols_horas = ["EQUIPAMENTO", "TIPO", "MARCA", "EMPRESA", "PERIODO",
                      "LEITURA_INICIAL", "LEITURA_FINAL", "HORAS_TRABALHADAS"]
        st.dataframe(
            horas[cols_horas].sort_values(["EQUIPAMENTO", "PERIODO"]),
            width='stretch', hide_index=True,
        )
        st.caption("HORAS_TRABALHADAS em branco = primeiro período do equipamento no filtro (sem leitura anterior para comparar) ou leitura de horímetro inconsistente (contador retrocedeu).")

        csv_horas = horas[cols_horas].to_csv(index=False).encode("utf-8-sig")
        st.download_button("⬇️ Baixar horas trabalhadas em CSV", csv_horas, "horas_trabalhadas.csv", "text/csv")

        with st.expander("🚗 Quilometragem rodada (caminhonete, ônibus, van — medidos por odômetro)"):
            comb_km = comb[comb["UNIDADE"] == "km"]
            km = calc_delta_por_periodo(comb_km, PERIODOS[periodo_label], "KM_RODADOS")
            if ini is not None:
                km = km[(km["PERIODO_INICIO"] >= ini.to_period(PERIODOS[periodo_label]).start_time) &
                        (km["PERIODO_INICIO"] <= fim)]
            km_validos = km.dropna(subset=["KM_RODADOS"])
            st.metric("Total de KM Rodados", f"{km_validos['KM_RODADOS'].sum():,.0f} km")
            cols_km = ["EQUIPAMENTO", "TIPO", "MARCA", "EMPRESA", "PERIODO",
                       "LEITURA_INICIAL", "LEITURA_FINAL", "KM_RODADOS"]
            st.dataframe(km[cols_km].sort_values(["EQUIPAMENTO", "PERIODO"]), width='stretch', hide_index=True)
