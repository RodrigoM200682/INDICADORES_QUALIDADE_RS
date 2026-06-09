"""
SMQ_RS — Sistema de Monitoramento de Qualidade RS
100% Python / Streamlit / Plotly.
Planilha salva em data/planilha.xlsx — persiste entre reinícios.
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import json
from datetime import datetime, date
from pathlib import Path
from io import BytesIO

# ── Caminhos ──────────────────────────────────────────────────────────────────
BASE_DIR      = Path(__file__).parent
DATA_DIR      = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)
PLANILHA_FILE = DATA_DIR / "planilha.xlsx"
META_FILE     = DATA_DIR / "meta.json"

STATUS_INVALIDOS = {"Reprovada", "Cancelada"}
CORES_STATUS = {
    "Concluída": "#22c55e", "Aberta": "#f59e0b",
    "Associada": "#a78bfa", "Reprovada": "#ef4444", "Cancelada": "#6b7280",
}
CORES_TURNO = ["#3b82f6", "#22c55e", "#f59e0b", "#a78bfa", "#6b7280"]

# ── Página ────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="SMQ_RS", page_icon="🔵",
    layout="wide", initial_sidebar_state="expanded",
)
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@300;400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');
html,body,[class*="css"]{font-family:'IBM Plex Sans',sans-serif;}
.stApp{background:#0f1117;color:#e8ecf4;}
.block-container{padding:1.5rem 2rem!important;}
[data-testid="stSidebar"]{background:#181c27;border-right:1px solid rgba(255,255,255,0.08);}
[data-testid="stSidebar"] *{color:#e8ecf4!important;}
div[data-testid="metric-container"]{background:#181c27;border:1px solid rgba(255,255,255,0.08);border-radius:8px;padding:12px 16px;}
div[data-testid="metric-container"] label{color:#8892a8!important;font-size:11px!important;text-transform:uppercase;letter-spacing:.06em;}
div[data-testid="metric-container"] div[data-testid="stMetricValue"]{font-family:'IBM Plex Mono',monospace;font-size:24px!important;}
.stTabs [data-baseweb="tab-list"]{background:#181c27;border-radius:8px;padding:4px;gap:4px;}
.stTabs [data-baseweb="tab"]{background:transparent;color:#8892a8;border-radius:6px;font-size:12px;font-weight:500;letter-spacing:.04em;text-transform:uppercase;}
.stTabs [aria-selected="true"]{background:#0f1117;color:#60a5fa!important;}
.stButton>button[kind="primary"]{background:#3b82f6;border:none;font-weight:500;}
.stButton>button[kind="primary"]:hover{background:#2563eb;}
</style>
""", unsafe_allow_html=True)

LAYOUT = dict(
    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family="IBM Plex Sans", color="#8892a8", size=11),
    margin=dict(l=10, r=10, t=40, b=10),
    legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(color="#8892a8")),
    title_font=dict(color="#e8ecf4", size=13),
)


# ══════════════════════════════════════════════════════════════════════════════
# PERSISTÊNCIA
# ══════════════════════════════════════════════════════════════════════════════

def salvar_planilha(file_bytes: bytes, ts: str) -> None:
    PLANILHA_FILE.write_bytes(file_bytes)
    META_FILE.write_text(json.dumps({"salvo_em": ts}), encoding="utf-8")


def carregar_planilha() -> tuple[bytes | None, str]:
    if PLANILHA_FILE.exists():
        try:
            meta = json.loads(META_FILE.read_text()) if META_FILE.exists() else {}
            return PLANILHA_FILE.read_bytes(), meta.get("salvo_em", "—")
        except Exception:
            pass
    return None, ""


# ══════════════════════════════════════════════════════════════════════════════
# LEITURA DA PLANILHA
# ══════════════════════════════════════════════════════════════════════════════

@st.cache_data(show_spinner=False)
def ler_xlsx(file_bytes: bytes) -> pd.DataFrame:
    df = pd.read_excel(BytesIO(file_bytes), engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    def col(palavras):
        for p in palavras:
            for c in df.columns:
                if p.lower() in c.lower():
                    return c
        return None

    mapa = {
        "codigo":            col(["código","codigo"]),
        "titulo":            col(["título","titulo"]),
        "status":            col(["status"]),
        "situacao":          col(["situação","situacao"]),
        "data":              col(["emissão","emissao","data"]),
        "responsavel":       col(["responsável","responsavel"]),
        "cliente":           col(["cliente"]),
        "responsavel_causa": col(["análise de causa","analise de causa"]),
        "motivo":            col(["motivo"]),
        "qtd":               col(["quantidade"]),
        "turno_raw":         col(["turno"]),
    }

    out = pd.DataFrame()
    for novo, original in mapa.items():
        out[novo] = df[original] if (original and original in df.columns) else ""

    def norm_turno(v):
        s = str(v)
        if "1" in s and "2" in s and "3" in s: return "Múltiplos Turnos"
        if "1" in s: return "1° Turno"
        if "2" in s: return "2° Turno"
        if "3" in s: return "3° Turno"
        return "Não Informado"

    out["turno"]   = out["turno_raw"].apply(norm_turno)
    out["data_dt"] = pd.to_datetime(out["data"], errors="coerce")
    out["ano"]     = out["data_dt"].dt.year.astype("Int64")
    out["mes"]     = out["data_dt"].dt.month.astype("Int64")
    out["mes_ano"] = out["data_dt"].dt.to_period("M").astype(str)
    out = out[(out["codigo"].astype(str).str.strip() != "") &
              (out["codigo"].astype(str) != "nan")].reset_index(drop=True)
    return out


def validos(df: pd.DataFrame) -> pd.DataFrame:
    return df[~df["status"].isin(STATUS_INVALIDOS)]


def aplicar_filtros(df: pd.DataFrame, f: dict) -> pd.DataFrame:
    d = df.copy()
    if f.get("de") and f.get("ate"):
        d = d[(d["data_dt"] >= pd.Timestamp(f["de"])) &
              (d["data_dt"] <= pd.Timestamp(f["ate"]))]
    for campo, coluna in [("status","status"),("turno","turno"),
                           ("cliente","cliente"),("motivo","motivo"),
                           ("responsavel","responsavel_causa")]:
        if f.get(campo):
            d = d[d[coluna].isin(f[campo])]
    return d


# ══════════════════════════════════════════════════════════════════════════════
# INICIALIZAÇÃO — carrega do disco automaticamente
# ══════════════════════════════════════════════════════════════════════════════

if "df" not in st.session_state:
    xlsx_bytes, salvo_em = carregar_planilha()
    if xlsx_bytes:
        try:
            st.session_state["df"]       = ler_xlsx(xlsx_bytes)
            st.session_state["salvo_em"] = salvo_em
        except Exception:
            st.session_state["df"]       = None
            st.session_state["salvo_em"] = ""
    else:
        st.session_state["df"]       = None
        st.session_state["salvo_em"] = ""


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("## 🔵 SMQ_RS")
    st.caption("Sistema de Monitoramento de Qualidade")
    st.divider()

    if st.session_state["df"] is not None:
        n = len(st.session_state["df"])
        st.success(f"✅ {n:,} registros carregados")
        if st.session_state["salvo_em"]:
            st.caption(f"🕐 {st.session_state['salvo_em']}")
    else:
        st.info("Nenhuma planilha carregada")

    st.divider()
    st.markdown("### 📂 Planilha")
    st.caption("O arquivo é salvo internamente e carregado automaticamente a cada reinício.")

    arquivo = st.file_uploader("Selecione o .xlsx", type=["xlsx","xls"])

    if arquivo:
        if st.button("💾 Salvar e Carregar", type="primary", use_container_width=True):
            file_bytes = arquivo.read()
            try:
                with st.spinner("Processando e salvando..."):
                    df_novo = ler_xlsx(file_bytes)
                    ts = datetime.now().strftime("%d/%m/%Y às %H:%M")
                    salvar_planilha(file_bytes, ts)
                    st.session_state["df"]       = df_novo
                    st.session_state["salvo_em"] = ts
                    st.cache_data.clear()
                st.success(f"✅ {len(df_novo):,} registros salvos!\n\n🕐 {ts}")
                st.rerun()
            except Exception as e:
                st.error(f"❌ {e}")

    if PLANILHA_FILE.exists():
        st.divider()
        if st.button("🗑️ Remover planilha salva", type="secondary"):
            PLANILHA_FILE.unlink(missing_ok=True)
            META_FILE.unlink(missing_ok=True)
            st.session_state["df"]       = None
            st.session_state["salvo_em"] = ""
            st.cache_data.clear()
            st.rerun()

    st.divider()
    st.caption("SMQ_RS v5.0 · Streamlit + Plotly")


# ══════════════════════════════════════════════════════════════════════════════
# SEM DADOS
# ══════════════════════════════════════════════════════════════════════════════

if st.session_state["df"] is None:
    st.markdown("## 🔵 SMQ_RS")
    st.info("👈 Faça upload da planilha na sidebar para começar.")
    st.stop()

df_all = st.session_state["df"]


# ══════════════════════════════════════════════════════════════════════════════
# FILTROS GLOBAIS
# ══════════════════════════════════════════════════════════════════════════════

with st.expander("🔍 Filtros", expanded=True):
    c1,c2,c3,c4,c5,c6 = st.columns([1.2,1.2,1.5,1.5,1.5,1.5])
    dt_min = df_all["data_dt"].min().date() if df_all["data_dt"].notna().any() else date(2025,1,1)
    dt_max = df_all["data_dt"].max().date() if df_all["data_dt"].notna().any() else date.today()

    with c1: de  = st.date_input("De",  dt_min, min_value=dt_min, max_value=dt_max)
    with c2: ate = st.date_input("Até", dt_max, min_value=dt_min, max_value=dt_max)
    with c3: status_sel  = st.multiselect("Status",   sorted(df_all["status"].dropna().unique()), placeholder="Todos")
    with c4: turno_sel   = st.multiselect("Turno",    sorted(df_all["turno"].dropna().unique()),  placeholder="Todos")
    with c5: cliente_sel = st.multiselect("Cliente",  sorted(df_all["cliente"].dropna().unique()),placeholder="Todos")
    with c6: motivo_sel  = st.multiselect("Motivo",   sorted(df_all["motivo"].dropna().unique()), placeholder="Todos")
    c7, _ = st.columns([2,4])
    with c7: resp_sel = st.multiselect("Responsável análise", sorted(df_all["responsavel_causa"].dropna().unique()), placeholder="Todos")

filtros = dict(de=de, ate=ate,
               status=status_sel or None, turno=turno_sel or None,
               cliente=cliente_sel or None, motivo=motivo_sel or None,
               responsavel=resp_sel or None)

df  = aplicar_filtros(df_all, filtros)
df_v = validos(df)


# ══════════════════════════════════════════════════════════════════════════════
# ABAS
# ══════════════════════════════════════════════════════════════════════════════

tab1,tab2,tab3,tab4,tab5 = st.tabs([
    "📊 Visão Geral","📅 Por Período","🔀 Comparação",
    "👤 Por Responsável","📋 Registros",
])


# ─── VISÃO GERAL ──────────────────────────────────────────────────────────────
with tab1:
    st.caption("Base completa · filtros não aplicados nesta aba")
    d0  = df_all
    d0v = validos(d0)
    hoje = pd.Timestamp(date.today())

    k1,k2,k3,k4,k5,k6 = st.columns(6)
    k1.metric("Total na Base",  len(d0))
    k2.metric("Válidas",        len(d0v))
    k3.metric("Abertas",        len(d0[d0.status=="Aberta"]))
    k4.metric("Concluídas",     len(d0[d0.status=="Concluída"]))
    k5.metric("Reprovadas",     len(d0[d0.status=="Reprovada"]))
    k6.metric("Canceladas",     len(d0[d0.status=="Cancelada"]))

    st.divider()

    # RNCs em atraso
    em_atraso = d0[(d0.status.isin(["Aberta","Associada"])) & (d0.situacao=="Atrasada")].copy()
    em_atraso["Dias em Aberto"] = (hoje - em_atraso["data_dt"]).dt.days.astype("Int64")

    if len(em_atraso):
        st.markdown(f"### 🔴 RNCs em Atraso Oficial — {len(em_atraso)} ocorrência(s)")
        cols_show = [c for c in ["codigo","data","responsavel_causa","motivo","cliente","status","Dias em Aberto"] if c in em_atraso.columns or c=="Dias em Aberto"]
        st.dataframe(
            em_atraso.assign(**{"Dias em Aberto": em_atraso["Dias em Aberto"]})
                     [[c for c in ["codigo","data","responsavel_causa","motivo","cliente","status","Dias em Aberto"] if c in em_atraso.columns or c=="Dias em Aberto"]]
                     .sort_values("Dias em Aberto", ascending=False)
                     .rename(columns={"codigo":"Código","data":"Data","responsavel_causa":"Responsável","motivo":"Motivo","cliente":"Cliente","status":"Status"}),
            use_container_width=True, hide_index=True,
        )
    else:
        st.success("✅ Nenhuma RNC em atraso oficial na base.")

    st.divider()

    p1,p2,p3 = st.columns(3)
    with p1:
        cnt = d0["status"].value_counts().reset_index(); cnt.columns=["Status","Qtd"]
        fig = px.pie(cnt, names="Status", values="Qtd", title="Situação Geral",
                     color="Status", color_discrete_map=CORES_STATUS, hole=0.4)
        fig.update_layout(**LAYOUT); fig.update_traces(textfont_color="#e8ecf4")
        st.plotly_chart(fig, use_container_width=True)
    with p2:
        fig2 = px.pie(values=[len(d0v), len(d0)-len(d0v)],
                      names=["Válidas","Reprov.+Cancel."], title="Válidas vs Inválidas",
                      color_discrete_sequence=["#22c55e","#ef4444"], hole=0.4)
        fig2.update_layout(**LAYOUT); fig2.update_traces(textfont_color="#e8ecf4")
        st.plotly_chart(fig2, use_container_width=True)
    with p3:
        cnt_t = d0v["turno"].value_counts().reset_index(); cnt_t.columns=["Turno","Qtd"]
        fig3 = px.pie(cnt_t, names="Turno", values="Qtd", title="Distribuição por Turno (válidas)",
                      color_discrete_sequence=CORES_TURNO, hole=0.4)
        fig3.update_layout(**LAYOUT); fig3.update_traces(textfont_color="#e8ecf4")
        st.plotly_chart(fig3, use_container_width=True)

    by_m = d0.groupby("mes_ano").size().reset_index(name="Qtd")
    by_m = by_m[by_m.mes_ano!="NaT"].sort_values("mes_ano")
    fig_m = px.bar(by_m, x="mes_ano", y="Qtd", title="Evolução Mensal — Base Completa",
                   labels={"mes_ano":"Mês","Qtd":"Ocorrências"}, color_discrete_sequence=["#3b82f6"])
    fig_m.update_layout(**{**LAYOUT, "xaxis":dict(tickangle=-45, gridcolor="rgba(255,255,255,0.05)"),
                            "yaxis":dict(gridcolor="rgba(255,255,255,0.05)")})
    st.plotly_chart(fig_m, use_container_width=True)

    g1,g2 = st.columns(2)
    with g1:
        top_m = d0v["motivo"].value_counts().head(10).reset_index(); top_m.columns=["Motivo","Qtd"]
        fig_tm = px.bar(top_m, x="Qtd", y="Motivo", orientation="h", title="Top 10 Motivos (válidas)",
                        color_discrete_sequence=["#3b82f6"])
        fig_tm.update_layout(**{**LAYOUT,"yaxis":dict(autorange="reversed")})
        st.plotly_chart(fig_tm, use_container_width=True)
    with g2:
        top_c = d0v["cliente"].value_counts().head(10).reset_index(); top_c.columns=["Cliente","Qtd"]
        fig_tc = px.bar(top_c, x="Qtd", y="Cliente", orientation="h", title="Top 10 Clientes (válidas)",
                        color_discrete_sequence=["#14b8a6"])
        fig_tc.update_layout(**{**LAYOUT,"yaxis":dict(autorange="reversed")})
        st.plotly_chart(fig_tc, use_container_width=True)


# ─── POR PERÍODO ──────────────────────────────────────────────────────────────
with tab2:
    if df.empty:
        st.warning("Nenhum dado para os filtros selecionados.")
    else:
        k1,k2,k3,k4,k5,k6 = st.columns(6)
        k1.metric("Total RNCs", len(df)); k2.metric("Válidas", len(df_v))
        k3.metric("Abertas", len(df[df.status=="Aberta"]))
        k4.metric("Concluídas", len(df[df.status=="Concluída"]))
        k5.metric("Reprovadas", len(df[df.status=="Reprovada"]))
        k6.metric("Responsáveis", df_v["responsavel_causa"].nunique())

        g1,g2 = st.columns(2)
        with g1:
            by_m = df_v.groupby("mes_ano").size().reset_index(name="Qtd")
            by_m = by_m[by_m.mes_ano!="NaT"].sort_values("mes_ano")
            fig = px.bar(by_m, x="mes_ano", y="Qtd", title="Reclamações por Mês (válidas)",
                         labels={"mes_ano":"Mês"}, color_discrete_sequence=["#3b82f6"])
            fig.update_layout(**{**LAYOUT,"xaxis":dict(tickangle=-45,gridcolor="rgba(255,255,255,0.05)"),
                                  "yaxis":dict(gridcolor="rgba(255,255,255,0.05)")})
            st.plotly_chart(fig, use_container_width=True)
        with g2:
            cnt = df["status"].value_counts().reset_index(); cnt.columns=["Status","Qtd"]
            fig2 = px.pie(cnt, names="Status", values="Qtd", title="Status das RNCs",
                          color="Status", color_discrete_map=CORES_STATUS, hole=0.4)
            fig2.update_layout(**LAYOUT); fig2.update_traces(textfont_color="#e8ecf4")
            st.plotly_chart(fig2, use_container_width=True)

        g3,g4 = st.columns(2)
        with g3:
            top_m = df_v["motivo"].value_counts().head(10).reset_index(); top_m.columns=["Motivo","Qtd"]
            fig3 = px.bar(top_m, x="Qtd", y="Motivo", orientation="h", title="Top 10 Motivos",
                          color_discrete_sequence=["#3b82f6"])
            fig3.update_layout(**{**LAYOUT,"yaxis":dict(autorange="reversed")})
            st.plotly_chart(fig3, use_container_width=True)
        with g4:
            cnt_t = df_v["turno"].value_counts().reset_index(); cnt_t.columns=["Turno","Qtd"]
            fig4 = px.pie(cnt_t, names="Turno", values="Qtd", title="Por Turno (válidas)",
                          color_discrete_sequence=CORES_TURNO, hole=0.4)
            fig4.update_layout(**LAYOUT); fig4.update_traces(textfont_color="#e8ecf4")
            st.plotly_chart(fig4, use_container_width=True)

        top_cli = df_v["cliente"].value_counts().head(10).reset_index(); top_cli.columns=["Cliente","Qtd"]
        fig5 = px.bar(top_cli, x="Qtd", y="Cliente", orientation="h", title="Top 10 Clientes",
                      color_discrete_sequence=["#14b8a6"])
        fig5.update_layout(**{**LAYOUT,"yaxis":dict(autorange="reversed")})
        st.plotly_chart(fig5, use_container_width=True)


# ─── COMPARAÇÃO ───────────────────────────────────────────────────────────────
with tab3:
    st.markdown("#### Comparar dois intervalos (apenas ocorrências válidas)")
    cc1,cc2,cc3,cc4 = st.columns(4)
    with cc1: a_de  = st.date_input("A — De",  date(2025,1,1), key="ca_de")
    with cc2: a_ate = st.date_input("A — Até", date(2025,3,31), key="ca_ate")
    with cc3: b_de  = st.date_input("B — De",  date(2026,1,1), key="cb_de")
    with cc4: b_ate = st.date_input("B — Até", date(2026,3,31), key="cb_ate")

    fa = {**filtros, "de":a_de, "ate":a_ate, "status":None}
    fb = {**filtros, "de":b_de, "ate":b_ate, "status":None}
    dA = validos(aplicar_filtros(df_all, fa))
    dB = validos(aplicar_filtros(df_all, fb))
    lA = f"{a_de.strftime('%b/%y')}–{a_ate.strftime('%b/%y')}"
    lB = f"{b_de.strftime('%b/%y')}–{b_ate.strftime('%b/%y')}"

    k1,k2,k3,k4 = st.columns(4)
    k1.metric(f"Total A", len(dA), help=lA)
    k2.metric(f"Total B", len(dB), delta=len(dB)-len(dA), help=lB)
    k3.metric("Motivos A", dA["motivo"].nunique())
    k4.metric("Motivos B", dB["motivo"].nunique(), delta=dB["motivo"].nunique()-dA["motivo"].nunique())

    # Mensal lado a lado
    MESES = {1:"Jan",2:"Fev",3:"Mar",4:"Abr",5:"Mai",6:"Jun",
             7:"Jul",8:"Ago",9:"Set",10:"Out",11:"Nov",12:"Dez"}
    mA = dA.groupby("mes").size(); mB = dB.groupby("mes").size()
    meses = sorted(set(list(mA.index)+list(mB.index)))
    fig_c = go.Figure()
    fig_c.add_bar(x=[MESES[m] for m in meses], y=[mA.get(m,0) for m in meses],
                  name=lA, marker_color="#3b82f6", offsetgroup=0)
    fig_c.add_bar(x=[MESES[m] for m in meses], y=[mB.get(m,0) for m in meses],
                  name=lB, marker_color="#a78bfa", offsetgroup=1)
    fig_c.update_layout(**{**LAYOUT,"title":"Evolução Mensal Lado a Lado","barmode":"group",
                            "xaxis":dict(gridcolor="rgba(255,255,255,0.05)"),
                            "yaxis":dict(gridcolor="rgba(255,255,255,0.05)")})
    st.plotly_chart(fig_c, use_container_width=True)

    # Top motivos A vs B
    motA = dA["motivo"].value_counts()
    motB = dB["motivo"].value_counts()
    todos = sorted(set(list(motA.index)+list(motB.index)))
    df_mot = pd.DataFrame({"Motivo":todos, lA:[int(motA.get(m,0)) for m in todos],
                            lB:[int(motB.get(m,0)) for m in todos]})
    df_mot = df_mot.sort_values(lB,ascending=False).head(10)
    fig_mot = go.Figure()
    fig_mot.add_bar(y=df_mot["Motivo"], x=df_mot[lA], name=lA, orientation="h", marker_color="#3b82f6")
    fig_mot.add_bar(y=df_mot["Motivo"], x=df_mot[lB], name=lB, orientation="h", marker_color="#a78bfa")
    fig_mot.update_layout(**{**LAYOUT,"title":"Top Motivos A vs B","barmode":"group",
                              "yaxis":dict(autorange="reversed"),
                              "xaxis":dict(gridcolor="rgba(255,255,255,0.05)")})
    st.plotly_chart(fig_mot, use_container_width=True)

    # Tabela delta
    df_delta = df_mot.copy()
    df_delta["Variação"] = df_delta[lB] - df_delta[lA]
    st.dataframe(df_delta, use_container_width=True, hide_index=True)


# ─── POR RESPONSÁVEL ──────────────────────────────────────────────────────────
with tab4:
    if df_v.empty:
        st.warning("Nenhum dado válido.")
    else:
        by_resp = df_v.groupby("responsavel_causa").size().reset_index(name="RNCs").sort_values("RNCs",ascending=False)
        fig_r = px.bar(by_resp, x="RNCs", y="responsavel_causa", orientation="h",
                       title="RNCs por Responsável de Análise (válidas)",
                       color_discrete_sequence=["#3b82f6"])
        fig_r.update_layout(**{**LAYOUT,"yaxis":dict(autorange="reversed"),
                                "xaxis":dict(gridcolor="rgba(255,255,255,0.05)"),
                                "height":max(300,len(by_resp)*28)})
        st.plotly_chart(fig_r, use_container_width=True)

        st.markdown("#### Top 10 Responsáveis — Motivos Principais")
        top10 = by_resp.head(10)["responsavel_causa"].tolist()
        cols = st.columns(2)
        for i, resp in enumerate(top10):
            with cols[i%2]:
                df_r  = df_v[df_v["responsavel_causa"]==resp]
                top_m = df_r["motivo"].value_counts().head(3)
                with st.expander(f"**{resp}** ({len(df_r)} RNCs)"):
                    for mot, cnt in top_m.items():
                        st.progress(int(cnt/len(df_r)*100)/100, text=f"{mot[:45]} — {cnt}")


# ─── REGISTROS ────────────────────────────────────────────────────────────────
with tab5:
    st.caption(f"{len(df):,} registros com os filtros aplicados")
    colunas = [c for c in ["codigo","data","status","situacao","motivo",
                            "responsavel_causa","turno","cliente","qtd"] if c in df.columns]
    st.dataframe(
        df[colunas].rename(columns={
            "codigo":"Código","data":"Data","status":"Status","situacao":"Situação",
            "motivo":"Motivo","responsavel_causa":"Resp. Análise",
            "turno":"Turno","cliente":"Cliente","qtd":"Qtd",
        }),
        use_container_width=True, hide_index=True, height=600,
    )
    csv = df[colunas].to_csv(index=False).encode("utf-8")
    st.download_button("⬇️ Baixar tabela filtrada (.csv)", data=csv,
                       file_name=f"smq_rs_{datetime.now().strftime('%Y%m%d')}.csv",
                       mime="text/csv")
