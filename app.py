"""
RNC Dashboard - Relatório de Não Conformidades (Reclamações de Clientes)
Aplicativo BI para análise visual de ocorrências RNC.
"""

import os
import io
import glob
import datetime
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# ─────────────────────────────────────────────
# CONFIGURAÇÃO DA PÁGINA
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="RNC Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# ESTILOS CUSTOMIZADOS
# ─────────────────────────────────────────────
st.markdown("""
<style>
    .main { background-color: #f0f2f6; }
    .block-container { padding-top: 1rem; padding-bottom: 1rem; }

    .kpi-card {
        background: white;
        border-radius: 10px;
        padding: 18px 20px;
        text-align: center;
        box-shadow: 0 2px 8px rgba(0,0,0,0.10);
        border-left: 5px solid #1f77b4;
    }
    .kpi-card.red { border-left-color: #d62728; }
    .kpi-card.green { border-left-color: #2ca02c; }
    .kpi-card.orange { border-left-color: #ff7f0e; }
    .kpi-card.purple { border-left-color: #9467bd; }

    .kpi-value { font-size: 2.2rem; font-weight: 700; color: #1f2937; margin: 0; }
    .kpi-label { font-size: 0.85rem; color: #6b7280; margin-top: 4px; }
    .kpi-delta { font-size: 0.80rem; margin-top: 4px; }
    .kpi-delta.up { color: #d62728; }
    .kpi-delta.down { color: #2ca02c; }

    .section-title {
        font-size: 1.15rem;
        font-weight: 700;
        color: #1f2937;
        margin-bottom: 8px;
        padding-bottom: 4px;
        border-bottom: 2px solid #e5e7eb;
    }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] {
        background: white;
        border-radius: 8px 8px 0 0;
        padding: 8px 20px;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# CONSTANTES
# ─────────────────────────────────────────────
DATA_DIR = os.path.join(os.path.dirname(__file__), "data")
os.makedirs(DATA_DIR, exist_ok=True)

STATUS_COLORS = {
    "Concluída": "#2ca02c",
    "Aberta": "#1f77b4",
    "Reprovada": "#d62728",
    "Associada": "#9467bd",
    "Cancelada": "#7f7f7f",
}

SITUACAO_COLORS = {
    "Fechada no prazo": "#2ca02c",
    "No prazo": "#1f77b4",
    "Fechada atrasada": "#ff7f0e",
    "Atrasada": "#d62728",
}

MESES_PT = {
    1: "Jan", 2: "Fev", 3: "Mar", 4: "Abr", 5: "Mai", 6: "Jun",
    7: "Jul", 8: "Ago", 9: "Set", 10: "Out", 11: "Nov", 12: "Dez",
}

# ─────────────────────────────────────────────
# FUNÇÕES DE DADOS
# ─────────────────────────────────────────────

def get_latest_file() -> str | None:
    files = glob.glob(os.path.join(DATA_DIR, "*.xlsx"))
    if not files:
        return None
    return max(files, key=os.path.getmtime)


def load_data(filepath: str) -> pd.DataFrame:
    df = pd.read_excel(filepath)
    df["Data de emissão"] = pd.to_datetime(df["Data de emissão"], errors="coerce")
    df["Ano"] = df["Data de emissão"].dt.year
    df["Mês_num"] = df["Data de emissão"].dt.month
    df["Mês_abrev"] = df["Mês_num"].map(MESES_PT)
    df["Ano-Mês"] = df["Data de emissão"].dt.to_period("M").astype(str)
    df["Atrasada"] = df["Situação"].isin(["Atrasada", "Fechada atrasada"])
    df["Embalagem_clean"] = df["Embalagem"].fillna("Não informado")
    df["Motivo Reclamação"] = df["Motivo Reclamação"].fillna("Não informado")
    df["Cliente"] = df["Cliente"].fillna("Não informado")
    df["Responsável"] = df["Responsável"].fillna("Não informado")
    df["Situação"] = df["Situação"].fillna("Não informado")
    return df


def save_uploaded(uploaded_file) -> str:
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"Consultas_RNC_{timestamp}.xlsx"
    filepath = os.path.join(DATA_DIR, filename)
    with open(filepath, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return filepath


def apply_filters(df: pd.DataFrame, filters: dict) -> pd.DataFrame:
    if filters.get("date_start"):
        df = df[df["Data de emissão"] >= pd.Timestamp(filters["date_start"])]
    if filters.get("date_end"):
        df = df[df["Data de emissão"] <= pd.Timestamp(filters["date_end"])]
    if filters.get("status") and "Todos" not in filters["status"]:
        df = df[df["Status"].isin(filters["status"])]
    if filters.get("situacao") and "Todas" not in filters["situacao"]:
        df = df[df["Situação"].isin(filters["situacao"])]
    if filters.get("cliente") and "Todos" not in filters["cliente"]:
        df = df[df["Cliente"].isin(filters["cliente"])]
    if filters.get("responsavel") and "Todos" not in filters["responsavel"]:
        df = df[df["Responsável"].isin(filters["responsavel"])]
    if filters.get("motivo") and "Todos" not in filters["motivo"]:
        df = df[df["Motivo Reclamação"].isin(filters["motivo"])]
    return df


# ─────────────────────────────────────────────
# COMPONENTES VISUAIS – KPI CARDS
# ─────────────────────────────────────────────

def kpi_card(label: str, value, color: str = "blue", delta: str = None, delta_up: bool = True):
    color_class = {"blue": "", "red": "red", "green": "green", "orange": "orange", "purple": "purple"}.get(color, "")
    delta_html = ""
    if delta:
        arrow = "▲" if delta_up else "▼"
        cls = "up" if delta_up else "down"
        delta_html = f'<div class="kpi-delta {cls}">{arrow} {delta}</div>'
    return f"""
    <div class="kpi-card {color_class}">
        <div class="kpi-value">{value}</div>
        <div class="kpi-label">{label}</div>
        {delta_html}
    </div>
    """


# ─────────────────────────────────────────────
# EXPORTAÇÃO EXCEL
# ─────────────────────────────────────────────

def build_excel_report(df: pd.DataFrame, filters: dict, tab_name: str) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        wb = writer.book

        # Formatos
        fmt_title = wb.add_format({"bold": True, "font_size": 14, "font_color": "#1f2937", "bg_color": "#e0e7ff", "align": "center", "valign": "vcenter", "border": 1})
        fmt_header = wb.add_format({"bold": True, "bg_color": "#1f4e79", "font_color": "white", "align": "center", "valign": "vcenter", "border": 1, "text_wrap": True})
        fmt_cell = wb.add_format({"align": "left", "valign": "vcenter", "border": 1, "text_wrap": True})
        fmt_cell_center = wb.add_format({"align": "center", "valign": "vcenter", "border": 1})
        fmt_red = wb.add_format({"bg_color": "#fde8e8", "align": "left", "border": 1})
        fmt_green = wb.add_format({"bg_color": "#d1fae5", "align": "left", "border": 1})
        fmt_filter = wb.add_format({"bold": True, "font_color": "#4b5563", "bg_color": "#f3f4f6", "border": 1})
        fmt_filter_val = wb.add_format({"font_color": "#1f2937", "bg_color": "#f9fafb", "border": 1, "text_wrap": True})

        # ── Aba: Visão Geral ──
        ws = wb.add_worksheet("Visão Geral")
        ws.set_column("A:A", 28)
        ws.set_column("B:B", 18)
        ws.set_column("C:H", 14)

        ws.merge_range("A1:H1", "RNC DASHBOARD – RELATÓRIO DE NÃO CONFORMIDADES", fmt_title)
        ws.set_row(0, 28)
        ws.write("A2", f"Gerado em: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M')}", fmt_filter)
        ws.write("B2", f"Aba de análise: {tab_name}", fmt_filter)

        # Filtros aplicados
        ws.merge_range("A3:H3", "FILTROS APLICADOS", fmt_header)
        filtros = [
            ("Período", f"{filters.get('date_start', 'Início')} a {filters.get('date_end', 'Fim')}"),
            ("Status", ", ".join(filters.get("status", ["Todos"]))),
            ("Situação", ", ".join(filters.get("situacao", ["Todas"]))),
            ("Cliente", ", ".join(filters.get("cliente", ["Todos"]))),
            ("Responsável", ", ".join(filters.get("responsavel", ["Todos"]))),
            ("Motivo", ", ".join(filters.get("motivo", ["Todos"]))),
        ]
        for i, (k, v) in enumerate(filtros):
            ws.write(3 + i, 0, k, fmt_filter)
            ws.merge_range(3 + i, 1, 3 + i, 7, v, fmt_filter_val)

        # KPIs
        row = 10
        ws.merge_range(row, 0, row, 7, "INDICADORES PRINCIPAIS", fmt_header)
        row += 1
        kpis = [
            ("Total de Ocorrências", len(df)),
            ("Ocorrências em Atraso", int(df["Atrasada"].sum())),
            ("% em Atraso", f"{100*df['Atrasada'].mean():.1f}%"),
            ("Concluídas", int((df["Status"] == "Concluída").sum())),
            ("Abertas", int((df["Status"] == "Aberta").sum())),
            ("Reprovadas", int((df["Status"] == "Reprovada").sum())),
            ("Clientes Afetados", df["Cliente"].nunique()),
            ("Motivos Distintos", df["Motivo Reclamação"].nunique()),
        ]
        ws.write(row, 0, "Indicador", fmt_header)
        ws.write(row, 1, "Valor", fmt_header)
        row += 1
        for k, v in kpis:
            ws.write(row, 0, k, fmt_filter)
            ws.write(row, 1, str(v), fmt_cell_center)
            row += 1

        # ── Aba: Dados Detalhados ──
        cols_export = ["Código", "Título", "Status", "Situação", "Data de emissão",
                       "Responsável", "Cliente", "Embalagem", "Motivo Reclamação",
                       "Quantidade não conforme", "Turno/Horário", "Atrasada"]
        df_exp = df[cols_export].copy()
        df_exp["Data de emissão"] = df_exp["Data de emissão"].dt.strftime("%d/%m/%Y")
        df_exp["Atrasada"] = df_exp["Atrasada"].map({True: "SIM", False: "NÃO"})
        df_exp.to_excel(writer, sheet_name="Dados Detalhados", index=False)
        ws2 = writer.sheets["Dados Detalhados"]
        ws2.set_column("A:A", 12)
        ws2.set_column("B:B", 45)
        ws2.set_column("C:D", 16)
        ws2.set_column("E:E", 14)
        ws2.set_column("F:G", 22)
        ws2.set_column("H:I", 22)
        ws2.set_column("J:L", 18)
        for col_num, col_name in enumerate(df_exp.columns):
            ws2.write(0, col_num, col_name, fmt_header)
        for row_num in range(len(df_exp)):
            atrasada = df_exp.iloc[row_num]["Atrasada"] == "SIM"
            row_fmt = fmt_red if atrasada else fmt_cell
            for col_num, val in enumerate(df_exp.iloc[row_num]):
                ws2.write(row_num + 1, col_num, str(val) if pd.notna(val) else "", row_fmt)

        # ── Aba: Por Mês ──
        mensal = (
            df.groupby("Ano-Mês").size().reset_index(name="Ocorrências")
        )
        mensal_atr = df[df["Atrasada"]].groupby("Ano-Mês").size().reset_index(name="Em Atraso")
        mensal = mensal.merge(mensal_atr, on="Ano-Mês", how="left").fillna(0)
        mensal["Em Atraso"] = mensal["Em Atraso"].astype(int)
        mensal.to_excel(writer, sheet_name="Evolução Mensal", index=False)
        ws3 = writer.sheets["Evolução Mensal"]
        ws3.set_column("A:A", 14)
        ws3.set_column("B:C", 18)
        for col_num, col_name in enumerate(mensal.columns):
            ws3.write(0, col_num, col_name, fmt_header)

        # ── Aba: Motivos ──
        motivos = df["Motivo Reclamação"].value_counts().reset_index()
        motivos.columns = ["Motivo", "Quantidade"]
        motivos["% do Total"] = (100 * motivos["Quantidade"] / len(df)).round(1)
        motivos["% Acumulado"] = motivos["% do Total"].cumsum().round(1)
        motivos.to_excel(writer, sheet_name="Pareto Motivos", index=False)
        ws4 = writer.sheets["Pareto Motivos"]
        ws4.set_column("A:A", 45)
        ws4.set_column("B:D", 16)
        for col_num, col_name in enumerate(motivos.columns):
            ws4.write(0, col_num, col_name, fmt_header)

        # ── Aba: Por Cliente ──
        clientes = df.groupby("Cliente").agg(
            Ocorrências=("Código", "count"),
            Em_Atraso=("Atrasada", "sum"),
        ).reset_index().sort_values("Ocorrências", ascending=False)
        clientes["% Atraso"] = (100 * clientes["Em_Atraso"] / clientes["Ocorrências"]).round(1)
        clientes.to_excel(writer, sheet_name="Por Cliente", index=False)
        ws5 = writer.sheets["Por Cliente"]
        ws5.set_column("A:A", 28)
        ws5.set_column("B:D", 16)
        for col_num, col_name in enumerate(clientes.columns):
            ws5.write(0, col_num, col_name, fmt_header)

        # ── Aba: Por Responsável ──
        resp = df.groupby("Responsável").agg(
            Ocorrências=("Código", "count"),
            Em_Atraso=("Atrasada", "sum"),
        ).reset_index().sort_values("Ocorrências", ascending=False)
        resp["% Atraso"] = (100 * resp["Em_Atraso"] / resp["Ocorrências"]).round(1)
        resp.to_excel(writer, sheet_name="Por Responsável", index=False)
        ws6 = writer.sheets["Por Responsável"]
        ws6.set_column("A:A", 32)
        ws6.set_column("B:D", 16)
        for col_num, col_name in enumerate(resp.columns):
            ws6.write(0, col_num, col_name, fmt_header)

    return output.getvalue()


# ─────────────────────────────────────────────
# SIDEBAR – UPLOAD E FILTROS
# ─────────────────────────────────────────────

def render_sidebar(df_full: pd.DataFrame) -> dict:
    with st.sidebar:
        st.image("https://img.icons8.com/color/96/combo-chart.png", width=64)
        st.title("RNC Dashboard")
        st.caption("Reclamações de Clientes · Brasilata RS")
        st.divider()

        st.subheader("📂 Importar Planilha")
        uploaded = st.file_uploader("Nova planilha Excel (.xlsx)", type=["xlsx"], label_visibility="collapsed")
        if uploaded:
            filepath = save_uploaded(uploaded)
            st.success(f"Arquivo salvo! Recarregando...")
            st.session_state["data_path"] = filepath
            st.rerun()

        latest = get_latest_file()
        if latest:
            st.caption(f"Arquivo ativo: `{os.path.basename(latest)}`")
        st.divider()

        st.subheader("🔍 Filtros")

        min_date = df_full["Data de emissão"].min().date()
        max_date = df_full["Data de emissão"].max().date()

        col1, col2 = st.columns(2)
        with col1:
            date_start = st.date_input("De", value=min_date, min_value=min_date, max_value=max_date)
        with col2:
            date_end = st.date_input("Até", value=max_date, min_value=min_date, max_value=max_date)

        status_opts = ["Todos"] + sorted(df_full["Status"].dropna().unique().tolist())
        status_sel = st.multiselect("Status", status_opts, default=["Todos"])

        sit_opts = ["Todas"] + sorted(df_full["Situação"].dropna().unique().tolist())
        sit_sel = st.multiselect("Situação", sit_opts, default=["Todas"])

        cliente_opts = ["Todos"] + sorted(df_full["Cliente"].dropna().unique().tolist())
        cliente_sel = st.multiselect("Cliente", cliente_opts, default=["Todos"])

        resp_opts = ["Todos"] + sorted(df_full["Responsável"].dropna().unique().tolist())
        resp_sel = st.multiselect("Responsável", resp_opts, default=["Todos"])

        motivo_opts = ["Todos"] + sorted(df_full["Motivo Reclamação"].dropna().unique().tolist())
        motivo_sel = st.multiselect("Motivo", motivo_opts, default=["Todos"])

        st.divider()
        if st.button("🔄 Limpar Filtros", use_container_width=True):
            st.rerun()

    return {
        "date_start": date_start,
        "date_end": date_end,
        "status": status_sel,
        "situacao": sit_sel,
        "cliente": cliente_sel,
        "responsavel": resp_sel,
        "motivo": motivo_sel,
    }


# ─────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────

def tab_visao_geral(df: pd.DataFrame, df_full: pd.DataFrame):
    st.markdown('<div class="section-title">📋 Visão Geral das Ocorrências</div>', unsafe_allow_html=True)

    total = len(df)
    atrasadas = int(df["Atrasada"].sum())
    pct_atraso = 100 * atrasadas / total if total > 0 else 0
    concluidas = int((df["Status"] == "Concluída").sum())
    abertas = int((df["Status"] == "Aberta").sum())

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        st.markdown(kpi_card("Total de Ocorrências", total, "blue"), unsafe_allow_html=True)
    with c2:
        st.markdown(kpi_card("Em Atraso", atrasadas, "red"), unsafe_allow_html=True)
    with c3:
        st.markdown(kpi_card("% em Atraso", f"{pct_atraso:.1f}%", "orange"), unsafe_allow_html=True)
    with c4:
        st.markdown(kpi_card("Concluídas", concluidas, "green"), unsafe_allow_html=True)
    with c5:
        st.markdown(kpi_card("Abertas", abertas, "purple"), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    col1, col2 = st.columns([2, 1])

    with col1:
        # Evolução mensal
        mensal = df.groupby("Ano-Mês").size().reset_index(name="Ocorrências")
        mensal_atr = df[df["Atrasada"]].groupby("Ano-Mês").size().reset_index(name="Em Atraso")
        mensal = mensal.merge(mensal_atr, on="Ano-Mês", how="left").fillna(0)
        mensal["Em Atraso"] = mensal["Em Atraso"].astype(int)

        fig = make_subplots(specs=[[{"secondary_y": False}]])
        fig.add_trace(go.Bar(x=mensal["Ano-Mês"], y=mensal["Ocorrências"], name="Total", marker_color="#1f77b4", opacity=0.85))
        fig.add_trace(go.Bar(x=mensal["Ano-Mês"], y=mensal["Em Atraso"], name="Em Atraso", marker_color="#d62728", opacity=0.85))
        fig.update_layout(
            title="Evolução Mensal de Ocorrências",
            barmode="overlay",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            plot_bgcolor="white",
            paper_bgcolor="white",
            height=340,
            margin=dict(t=50, b=30, l=30, r=10),
            xaxis=dict(tickangle=-45),
        )
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        # Distribuição por Status
        status_cnt = df["Status"].value_counts().reset_index()
        status_cnt.columns = ["Status", "Quantidade"]
        colors = [STATUS_COLORS.get(s, "#aaa") for s in status_cnt["Status"]]
        fig2 = px.pie(
            status_cnt, values="Quantidade", names="Status",
            title="Distribuição por Status",
            color="Status",
            color_discrete_map=STATUS_COLORS,
            hole=0.42,
        )
        fig2.update_layout(
            height=340,
            margin=dict(t=50, b=10, l=10, r=10),
            paper_bgcolor="white",
            legend=dict(orientation="h", yanchor="bottom", y=-0.25),
        )
        fig2.update_traces(textposition="outside", textinfo="percent+label")
        st.plotly_chart(fig2, use_container_width=True)

    col3, col4 = st.columns([1, 1])
    with col3:
        # Distribuição por Situação
        sit_cnt = df["Situação"].value_counts().reset_index()
        sit_cnt.columns = ["Situação", "Quantidade"]
        fig3 = px.bar(
            sit_cnt.sort_values("Quantidade"),
            x="Quantidade", y="Situação",
            orientation="h",
            title="Distribuição por Situação",
            color="Situação",
            color_discrete_map=SITUACAO_COLORS,
        )
        fig3.update_layout(
            height=300,
            showlegend=False,
            plot_bgcolor="white",
            paper_bgcolor="white",
            margin=dict(t=40, b=20, l=5, r=10),
        )
        st.plotly_chart(fig3, use_container_width=True)

    with col4:
        # Top 10 clientes
        top_cli = df["Cliente"].value_counts().head(10).reset_index()
        top_cli.columns = ["Cliente", "Ocorrências"]
        fig4 = px.bar(
            top_cli.sort_values("Ocorrências"),
            x="Ocorrências", y="Cliente",
            orientation="h",
            title="Top 10 Clientes com mais Ocorrências",
            color="Ocorrências",
            color_continuous_scale="Blues",
        )
        fig4.update_layout(
            height=300,
            showlegend=False,
            plot_bgcolor="white",
            paper_bgcolor="white",
            margin=dict(t=40, b=20, l=5, r=10),
            coloraxis_showscale=False,
        )
        st.plotly_chart(fig4, use_container_width=True)


def tab_evolucao(df: pd.DataFrame):
    st.markdown('<div class="section-title">📈 Evolução das Ocorrências</div>', unsafe_allow_html=True)

    granularidade = st.radio("Granularidade", ["Mensal", "Trimestral", "Anual"], horizontal=True, key="gran")

    if granularidade == "Mensal":
        grp_col = "Ano-Mês"
        df_grp = df.copy()
    elif granularidade == "Trimestral":
        df_grp = df.copy()
        df_grp["Trim"] = df_grp["Data de emissão"].dt.to_period("Q").astype(str)
        grp_col = "Trim"
    else:
        df_grp = df.copy()
        grp_col = "Ano"

    mensal = df_grp.groupby(grp_col).size().reset_index(name="Total")
    mensal_atr = df_grp[df_grp["Atrasada"]].groupby(grp_col).size().reset_index(name="Em Atraso")
    mensal_ok = df_grp[~df_grp["Atrasada"]].groupby(grp_col).size().reset_index(name="No Prazo")
    mensal = mensal.merge(mensal_atr, on=grp_col, how="left").merge(mensal_ok, on=grp_col, how="left").fillna(0)
    mensal["Em Atraso"] = mensal["Em Atraso"].astype(int)
    mensal["No Prazo"] = mensal["No Prazo"].astype(int)
    mensal["% Atraso"] = (100 * mensal["Em Atraso"] / mensal["Total"]).round(1)

    fig = make_subplots(rows=2, cols=1, shared_xaxes=True, vertical_spacing=0.12,
                        subplot_titles=["Quantidade de Ocorrências", "% em Atraso"])

    fig.add_trace(go.Bar(x=mensal[grp_col].astype(str), y=mensal["No Prazo"], name="No Prazo",
                         marker_color="#2ca02c", opacity=0.85), row=1, col=1)
    fig.add_trace(go.Bar(x=mensal[grp_col].astype(str), y=mensal["Em Atraso"], name="Em Atraso",
                         marker_color="#d62728", opacity=0.85), row=1, col=1)
    fig.add_trace(go.Scatter(x=mensal[grp_col].astype(str), y=mensal["Total"], name="Total",
                             mode="lines+markers", line=dict(color="#1f77b4", width=2.5),
                             marker=dict(size=7)), row=1, col=1)
    fig.add_trace(go.Scatter(x=mensal[grp_col].astype(str), y=mensal["% Atraso"], name="% Atraso",
                             mode="lines+markers+text", line=dict(color="#ff7f0e", width=2.5, dash="dot"),
                             text=mensal["% Atraso"].apply(lambda x: f"{x:.1f}%"),
                             textposition="top center", marker=dict(size=7)), row=2, col=1)

    fig.update_layout(
        barmode="stack",
        height=500,
        plot_bgcolor="white",
        paper_bgcolor="white",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(t=60, b=30, l=30, r=10),
    )
    fig.update_xaxes(tickangle=-45)
    st.plotly_chart(fig, use_container_width=True)

    # Heatmap: Ano x Mês
    st.markdown('<div class="section-title">Mapa de Calor – Ocorrências por Ano/Mês</div>', unsafe_allow_html=True)
    pivot = df.groupby(["Ano", "Mês_num"]).size().unstack(fill_value=0)
    pivot.columns = [MESES_PT[c] for c in pivot.columns]

    fig_heat = go.Figure(go.Heatmap(
        z=pivot.values,
        x=list(pivot.columns),
        y=pivot.index.astype(str).tolist(),
        colorscale="Blues",
        text=pivot.values,
        texttemplate="%{text}",
        showscale=True,
        colorbar=dict(title="Qtd"),
    ))
    fig_heat.update_layout(
        title="Ocorrências por Mês e Ano",
        height=280,
        plot_bgcolor="white",
        paper_bgcolor="white",
        margin=dict(t=40, b=20, l=40, r=10),
    )
    st.plotly_chart(fig_heat, use_container_width=True)


def tab_motivos(df: pd.DataFrame):
    st.markdown('<div class="section-title">🔎 Análise de Motivos de Ocorrências</div>', unsafe_allow_html=True)

    col1, col2 = st.columns([3, 2])

    with col1:
        # Pareto
        motivos = df["Motivo Reclamação"].value_counts().reset_index()
        motivos.columns = ["Motivo", "Quantidade"]
        motivos["% do Total"] = (100 * motivos["Quantidade"] / len(df)).round(1)
        motivos["% Acumulado"] = motivos["% do Total"].cumsum().round(1)

        top_n = st.slider("Mostrar top N motivos", 5, min(30, len(motivos)), 15, key="top_n_motivos")
        motivos_top = motivos.head(top_n)

        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_trace(go.Bar(x=motivos_top["Motivo"], y=motivos_top["Quantidade"],
                             name="Quantidade", marker_color="#1f77b4", opacity=0.85))
        fig.add_trace(go.Scatter(x=motivos_top["Motivo"], y=motivos_top["% Acumulado"],
                                 name="% Acumulado", mode="lines+markers",
                                 line=dict(color="#d62728", width=2.5),
                                 marker=dict(size=7)), secondary_y=True)
        fig.add_hline(y=80, line_dash="dash", line_color="#ff7f0e",
                      annotation_text="80%", secondary_y=True)
        fig.update_layout(
            title=f"Diagrama de Pareto – Top {top_n} Motivos",
            height=420,
            plot_bgcolor="white",
            paper_bgcolor="white",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            margin=dict(t=50, b=80, l=30, r=10),
            xaxis=dict(tickangle=-40),
        )
        fig.update_yaxes(title_text="Quantidade", secondary_y=False)
        fig.update_yaxes(title_text="% Acumulado", secondary_y=True, range=[0, 105])
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        # Pie motivos (top 10)
        motivos_pie = motivos.head(10).copy()
        outros = pd.DataFrame([{"Motivo": "Outros", "Quantidade": motivos.iloc[10:]["Quantidade"].sum(),
                                 "% do Total": motivos.iloc[10:]["% do Total"].sum(), "% Acumulado": 100}])
        if len(motivos) > 10:
            motivos_pie = pd.concat([motivos_pie, outros], ignore_index=True)

        fig2 = px.pie(motivos_pie, values="Quantidade", names="Motivo",
                      title="Top 10 Motivos + Outros",
                      hole=0.38, color_discrete_sequence=px.colors.qualitative.Plotly)
        fig2.update_layout(
            height=420,
            paper_bgcolor="white",
            margin=dict(t=50, b=10, l=10, r=10),
            legend=dict(orientation="v", x=1.02),
        )
        fig2.update_traces(textposition="inside", textinfo="percent")
        st.plotly_chart(fig2, use_container_width=True)

    # Motivos por Status
    st.markdown('<div class="section-title">Motivos × Status</div>', unsafe_allow_html=True)
    top_motivos = motivos.head(12)["Motivo"].tolist()
    df_cross = df[df["Motivo Reclamação"].isin(top_motivos)]
    cross = df_cross.groupby(["Motivo Reclamação", "Status"]).size().reset_index(name="Qtd")
    fig3 = px.bar(cross, x="Motivo Reclamação", y="Qtd", color="Status",
                  title="Top 12 Motivos por Status",
                  barmode="stack",
                  color_discrete_map=STATUS_COLORS)
    fig3.update_layout(
        height=380,
        plot_bgcolor="white",
        paper_bgcolor="white",
        margin=dict(t=40, b=80, l=30, r=10),
        xaxis=dict(tickangle=-35),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    st.plotly_chart(fig3, use_container_width=True)


def tab_atrasos(df: pd.DataFrame):
    st.markdown('<div class="section-title">⚠️ Ocorrências em Atraso</div>', unsafe_allow_html=True)

    df_atr = df[df["Atrasada"]].copy()
    df_ok = df[~df["Atrasada"]].copy()
    total = len(df)
    n_atr = len(df_atr)
    pct = 100 * n_atr / total if total > 0 else 0

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(kpi_card("Total em Atraso", n_atr, "red"), unsafe_allow_html=True)
    with c2:
        st.markdown(kpi_card("% em Atraso", f"{pct:.1f}%", "orange"), unsafe_allow_html=True)
    with c3:
        st.markdown(kpi_card("No Prazo", len(df_ok), "green"), unsafe_allow_html=True)
    with c4:
        st.markdown(kpi_card("Clientes com Atraso", df_atr["Cliente"].nunique(), "purple"), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        # Atrasos por mês
        atr_mes = df_atr.groupby("Ano-Mês").size().reset_index(name="Em Atraso")
        tot_mes = df.groupby("Ano-Mês").size().reset_index(name="Total")
        atr_mes = atr_mes.merge(tot_mes, on="Ano-Mês")
        atr_mes["% Atraso"] = (100 * atr_mes["Em Atraso"] / atr_mes["Total"]).round(1)

        fig = px.bar(atr_mes, x="Ano-Mês", y="% Atraso",
                     title="% de Atraso por Mês",
                     color="% Atraso",
                     color_continuous_scale=["#2ca02c", "#ff7f0e", "#d62728"],
                     text=atr_mes["% Atraso"].apply(lambda x: f"{x:.0f}%"))
        fig.update_traces(textposition="outside")
        fig.update_layout(
            height=340, plot_bgcolor="white", paper_bgcolor="white",
            margin=dict(t=40, b=50, l=30, r=10),
            xaxis=dict(tickangle=-45), coloraxis_showscale=False,
        )
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        # Atrasos por responsável
        resp_atr = df_atr["Responsável"].value_counts().head(10).reset_index()
        resp_atr.columns = ["Responsável", "Em Atraso"]
        fig2 = px.bar(resp_atr.sort_values("Em Atraso"),
                      x="Em Atraso", y="Responsável", orientation="h",
                      title="Atrasos por Responsável (Top 10)",
                      color="Em Atraso",
                      color_continuous_scale="Reds")
        fig2.update_layout(
            height=340, plot_bgcolor="white", paper_bgcolor="white",
            margin=dict(t=40, b=20, l=5, r=10), coloraxis_showscale=False,
        )
        st.plotly_chart(fig2, use_container_width=True)

    # Atrasos por motivo
    motivo_atr = df_atr["Motivo Reclamação"].value_counts().head(15).reset_index()
    motivo_atr.columns = ["Motivo", "Em Atraso"]
    fig3 = px.bar(motivo_atr, x="Motivo", y="Em Atraso",
                  title="Motivos com Maior Incidência de Atraso",
                  color="Em Atraso",
                  color_continuous_scale="Reds",
                  text="Em Atraso")
    fig3.update_traces(textposition="outside")
    fig3.update_layout(
        height=340, plot_bgcolor="white", paper_bgcolor="white",
        margin=dict(t=40, b=80, l=30, r=10),
        xaxis=dict(tickangle=-35), coloraxis_showscale=False,
    )
    st.plotly_chart(fig3, use_container_width=True)

    # Tabela detalhada
    st.markdown('<div class="section-title">Detalhamento – Ocorrências em Atraso</div>', unsafe_allow_html=True)
    cols_show = ["Código", "Título", "Status", "Situação", "Data de emissão",
                 "Responsável", "Cliente", "Motivo Reclamação"]
    df_show = df_atr[cols_show].copy()
    df_show["Data de emissão"] = df_show["Data de emissão"].dt.strftime("%d/%m/%Y")
    st.dataframe(
        df_show.reset_index(drop=True),
        use_container_width=True,
        height=400,
        column_config={
            "Título": st.column_config.TextColumn(width="large"),
            "Situação": st.column_config.TextColumn(width="medium"),
        }
    )


def tab_comparativo(df_full: pd.DataFrame):
    st.markdown('<div class="section-title">⚖️ Comparativo entre Períodos</div>', unsafe_allow_html=True)

    anos = sorted(df_full["Ano"].dropna().unique().tolist(), reverse=True)
    meses_opts = list(MESES_PT.items())

    col1, col2, col3 = st.columns([1, 1, 1])
    with col1:
        st.markdown("**Período A**")
        ano_a = st.selectbox("Ano A", anos, index=0, key="ano_a")
        mes_a = st.selectbox("Mês A", [f"{n} – {m}" for n, m in meses_opts], index=4, key="mes_a")
        mes_a_num = int(mes_a.split(" – ")[0])
    with col2:
        st.markdown("**Período B**")
        ano_b = st.selectbox("Ano B", anos, index=min(1, len(anos) - 1), key="ano_b")
        mes_b = st.selectbox("Mês B", [f"{n} – {m}" for n, m in meses_opts], index=4, key="mes_b")
        mes_b_num = int(mes_b.split(" – ")[0])
    with col3:
        st.markdown("**Tipo de Comparativo**")
        comp_type = st.radio("", ["Mês específico", "Ano completo"], key="comp_type")

    st.divider()

    if comp_type == "Mês específico":
        df_a = df_full[(df_full["Ano"] == ano_a) & (df_full["Mês_num"] == mes_a_num)]
        df_b = df_full[(df_full["Ano"] == ano_b) & (df_full["Mês_num"] == mes_b_num)]
        label_a = f"{MESES_PT[mes_a_num]}/{ano_a}"
        label_b = f"{MESES_PT[mes_b_num]}/{ano_b}"
    else:
        df_a = df_full[df_full["Ano"] == ano_a]
        df_b = df_full[df_full["Ano"] == ano_b]
        label_a = str(ano_a)
        label_b = str(ano_b)

    # KPIs comparativos
    def comp_kpi(label, val_a, val_b, fmt="{}", inverse=False):
        delta = val_b - val_a if isinstance(val_a, (int, float)) else 0
        pct = (100 * delta / val_a) if val_a else 0
        delta_str = f"{delta:+.0f} ({pct:+.1f}%)" if isinstance(delta, float) and delta % 1 != 0 else f"{delta:+.0f} ({pct:+.1f}%)"
        # For things like atraso: increase is bad (red), decrease is good (green)
        delta_up = delta > 0
        if inverse:
            delta_up = not delta_up
        return label, fmt.format(val_a), fmt.format(val_b), delta_str, delta_up

    kpis_data = [
        comp_kpi("Total Ocorrências", len(df_a), len(df_b)),
        comp_kpi("Em Atraso", int(df_a["Atrasada"].sum()), int(df_b["Atrasada"].sum()), inverse=True),
        comp_kpi("Concluídas", int((df_a["Status"] == "Concluída").sum()), int((df_b["Status"] == "Concluída").sum())),
        comp_kpi("Abertas", int((df_a["Status"] == "Aberta").sum()), int((df_b["Status"] == "Aberta").sum()), inverse=True),
    ]

    st.subheader(f"Comparativo: {label_a} vs {label_b}")
    cols = st.columns(4)
    for i, (label, va, vb, delta_str, delta_up) in enumerate(kpis_data):
        with cols[i]:
            st.metric(label=f"{label} – {label_b}", value=vb, delta=f"{delta_str} vs {label_a}",
                      delta_color="inverse" if "Atraso" in label or "Abertas" in label else "normal")

    st.markdown("<br>", unsafe_allow_html=True)
    col_a, col_b = st.columns(2)

    def make_status_chart(df_p, label):
        sc = df_p["Status"].value_counts().reset_index()
        sc.columns = ["Status", "Quantidade"]
        fig = px.pie(sc, values="Quantidade", names="Status", title=f"Status – {label}",
                     color="Status", color_discrete_map=STATUS_COLORS, hole=0.38)
        fig.update_layout(height=300, paper_bgcolor="white", margin=dict(t=40, b=5, l=5, r=5),
                          legend=dict(orientation="h", y=-0.2))
        fig.update_traces(textposition="outside", textinfo="percent+label")
        return fig

    with col_a:
        st.plotly_chart(make_status_chart(df_a, label_a), use_container_width=True)
    with col_b:
        st.plotly_chart(make_status_chart(df_b, label_b), use_container_width=True)

    # Motivos comparativos
    st.markdown('<div class="section-title">Comparativo de Motivos</div>', unsafe_allow_html=True)
    mot_a = df_a["Motivo Reclamação"].value_counts().head(10).reset_index()
    mot_a.columns = ["Motivo", label_a]
    mot_b = df_b["Motivo Reclamação"].value_counts().head(10).reset_index()
    mot_b.columns = ["Motivo", label_b]
    mot_comp = mot_a.merge(mot_b, on="Motivo", how="outer").fillna(0).sort_values(label_a, ascending=False)

    fig_comp = go.Figure()
    fig_comp.add_trace(go.Bar(name=label_a, x=mot_comp["Motivo"], y=mot_comp[label_a], marker_color="#1f77b4", opacity=0.85))
    fig_comp.add_trace(go.Bar(name=label_b, x=mot_comp["Motivo"], y=mot_comp[label_b], marker_color="#ff7f0e", opacity=0.85))
    fig_comp.update_layout(
        barmode="group",
        height=360,
        plot_bgcolor="white",
        paper_bgcolor="white",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(t=40, b=60, l=30, r=10),
        xaxis=dict(tickangle=-35),
    )
    st.plotly_chart(fig_comp, use_container_width=True)

    # Comparativo clientes
    st.markdown('<div class="section-title">Comparativo de Clientes</div>', unsafe_allow_html=True)
    cli_a = df_a["Cliente"].value_counts().head(10).reset_index()
    cli_a.columns = ["Cliente", label_a]
    cli_b = df_b["Cliente"].value_counts().head(10).reset_index()
    cli_b.columns = ["Cliente", label_b]
    cli_comp = cli_a.merge(cli_b, on="Cliente", how="outer").fillna(0).sort_values(label_a, ascending=False)

    fig_cli = go.Figure()
    fig_cli.add_trace(go.Bar(name=label_a, x=cli_comp["Cliente"], y=cli_comp[label_a], marker_color="#2ca02c", opacity=0.85))
    fig_cli.add_trace(go.Bar(name=label_b, x=cli_comp["Cliente"], y=cli_comp[label_b], marker_color="#9467bd", opacity=0.85))
    fig_cli.update_layout(
        barmode="group",
        height=340,
        plot_bgcolor="white",
        paper_bgcolor="white",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(t=40, b=50, l=30, r=10),
        xaxis=dict(tickangle=-35),
    )
    st.plotly_chart(fig_cli, use_container_width=True)


def tab_responsaveis(df: pd.DataFrame):
    st.markdown('<div class="section-title">👥 Análise por Responsável</div>', unsafe_allow_html=True)

    resp_df = df.groupby("Responsável").agg(
        Total=("Código", "count"),
        Em_Atraso=("Atrasada", "sum"),
        Concluidas=("Status", lambda x: (x == "Concluída").sum()),
        Abertas=("Status", lambda x: (x == "Aberta").sum()),
    ).reset_index()
    resp_df["% Atraso"] = (100 * resp_df["Em_Atraso"] / resp_df["Total"]).round(1)
    resp_df = resp_df.sort_values("Total", ascending=False)

    col1, col2 = st.columns([2, 1])
    with col1:
        fig = go.Figure()
        fig.add_trace(go.Bar(name="Concluídas", x=resp_df["Responsável"], y=resp_df["Concluidas"],
                             marker_color="#2ca02c", opacity=0.85))
        fig.add_trace(go.Bar(name="Em Atraso", x=resp_df["Responsável"], y=resp_df["Em_Atraso"],
                             marker_color="#d62728", opacity=0.85))
        fig.add_trace(go.Bar(name="Abertas", x=resp_df["Responsável"], y=resp_df["Abertas"],
                             marker_color="#1f77b4", opacity=0.85))
        fig.update_layout(
            barmode="stack",
            title="Ocorrências por Responsável",
            height=380,
            plot_bgcolor="white",
            paper_bgcolor="white",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            margin=dict(t=50, b=60, l=30, r=10),
            xaxis=dict(tickangle=-40),
        )
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        fig2 = px.bar(resp_df.sort_values("% Atraso", ascending=False).head(10),
                      x="% Atraso", y="Responsável", orientation="h",
                      title="% Atraso por Responsável",
                      color="% Atraso",
                      color_continuous_scale="Reds",
                      text=resp_df.sort_values("% Atraso", ascending=False).head(10)["% Atraso"].apply(lambda x: f"{x:.0f}%"))
        fig2.update_traces(textposition="outside")
        fig2.update_layout(
            height=380, plot_bgcolor="white", paper_bgcolor="white",
            margin=dict(t=40, b=20, l=5, r=10), coloraxis_showscale=False,
        )
        st.plotly_chart(fig2, use_container_width=True)

    # Tabela resumo
    st.markdown('<div class="section-title">Tabela Resumo por Responsável</div>', unsafe_allow_html=True)
    resp_show = resp_df.rename(columns={"Em_Atraso": "Em Atraso", "Concluidas": "Concluídas"})
    st.dataframe(resp_show.reset_index(drop=True), use_container_width=True, height=350)


# ─────────────────────────────────────────────
# DOWNLOAD BUTTON
# ─────────────────────────────────────────────

def render_download(df: pd.DataFrame, filters: dict, tab_name: str):
    with st.spinner("Preparando relatório Excel..."):
        excel_bytes = build_excel_report(df, filters, tab_name)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"RNC_Relatorio_{tab_name.replace(' ', '_')}_{timestamp}.xlsx"
    st.download_button(
        label="⬇️ Baixar Relatório Excel",
        data=excel_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

def main():
    # Carrega dados
    if "data_path" not in st.session_state:
        latest = get_latest_file()
        st.session_state["data_path"] = latest

    if not st.session_state["data_path"]:
        st.warning("Nenhuma planilha encontrada. Faça o upload de um arquivo Excel na barra lateral.")
        st.stop()

    @st.cache_data(ttl=300)
    def cached_load(path):
        return load_data(path)

    df_full = cached_load(st.session_state["data_path"])

    # Sidebar com filtros
    filters = render_sidebar(df_full)

    # Aplica filtros
    df = apply_filters(df_full.copy(), filters)

    if len(df) == 0:
        st.warning("Nenhuma ocorrência encontrada com os filtros aplicados.")
        st.stop()

    # Cabeçalho
    st.markdown(
        f"<h2 style='color:#1f2937;margin-bottom:0'>📊 RNC Dashboard – Reclamações de Clientes</h2>"
        f"<p style='color:#6b7280;margin-top:4px'>Exibindo <b>{len(df)}</b> de <b>{len(df_full)}</b> ocorrências | "
        f"Arquivo: <code>{os.path.basename(st.session_state['data_path'])}</code></p>",
        unsafe_allow_html=True,
    )

    # Tabs principais
    tab_labels = [
        "📋 Visão Geral",
        "📈 Evolução",
        "🔎 Motivos",
        "⚠️ Atrasos",
        "⚖️ Comparativo",
        "👥 Responsáveis",
    ]
    tabs = st.tabs(tab_labels)

    with tabs[0]:
        tab_visao_geral(df, df_full)
        st.divider()
        render_download(df, filters, "Visão_Geral")

    with tabs[1]:
        tab_evolucao(df)
        st.divider()
        render_download(df, filters, "Evolução")

    with tabs[2]:
        tab_motivos(df)
        st.divider()
        render_download(df, filters, "Motivos")

    with tabs[3]:
        tab_atrasos(df)
        st.divider()
        render_download(df, filters, "Atrasos")

    with tabs[4]:
        tab_comparativo(df_full)

    with tabs[5]:
        tab_responsaveis(df)
        st.divider()
        render_download(df, filters, "Responsáveis")


if __name__ == "__main__":
    main()
