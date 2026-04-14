"""
═══════════════════════════════════════════════════════════════
Dashboard: Faturamento x Despesa Realizada
Empresa: First / Integracsc
Arquivo de dados: Base_2.xlsx
Auto-reload: monitora o arquivo e recarrega ao detectar mudança

REGRAS DE FATURAMENTO:
  - Coluna D (valor) considerada SOMENTE quando coluna C = "Considerar"
  - Ano fixo: 2026
  - Corte automático: apenas meses com dados completos de despesa
═══════════════════════════════════════════════════════════════
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from pathlib import Path
import time
import os
import warnings

warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────────────────────
# CAMINHO DO ARQUIVO
# ─────────────────────────────────────────────────────────────
CAMINHO_EXCEL = Path(r"C:\Users\lara.matos\Desktop\Viagens\First\Base_2.xlsx")
if not CAMINHO_EXCEL.exists():
    CAMINHO_EXCEL = Path(__file__).parent / "Base_2.xlsx"

ANO_DASHBOARD = 2026

MESES_NOMES = {1:"Jan",2:"Fev",3:"Mar",4:"Abr",5:"Mai",6:"Jun",
               7:"Jul",8:"Ago",9:"Set",10:"Out",11:"Nov",12:"Dez"}
MESES_INV   = {v: k for k, v in MESES_NOMES.items()}

# ─────────────────────────────────────────────────────────────
# CONFIGURAÇÃO DA PÁGINA
# ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Faturamento x Despesa Realizada",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────────────────────────
# CSS GLOBAL
# ─────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .stApp { background-color: #f0f2f6; }
    #MainMenu, footer, header { visibility: hidden; }

    .header-bar {
        background: linear-gradient(90deg, #1a2744 0%, #2c3e7a 100%);
        padding: 16px 28px; border-radius: 12px; margin-bottom: 20px;
        display: flex; align-items: center; gap: 14px;
    }
    .header-title {
        color: white; font-size: 1.6rem; font-weight: 800;
        letter-spacing: 1px; margin: 0;
    }
    .header-logo {
        background: #e95d0c; border-radius: 50%;
        width: 48px; height: 48px;
        display: flex; align-items: center; justify-content: center;
        font-size: 1.4rem; flex-shrink: 0;
    }
    .header-badge {
        margin-left: auto; background: rgba(255,255,255,0.15);
        color: #a0b4d8; font-size: 0.7rem; padding: 4px 12px;
        border-radius: 20px; border: 1px solid rgba(255,255,255,0.2);
        white-space: nowrap;
    }
    .kpi-card {
        background: white; border-radius: 10px;
        padding: 16px 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.07);
        border-top: 4px solid #1a2744; height: 100%;
    }
    .kpi-card.laranja { border-top-color: #e95d0c; }
    .kpi-card.verde   { border-top-color: #27ae60; }
    .kpi-card.azul    { border-top-color: #2980b9; }
    .kpi-card.vermelho{ border-top-color: #c0392b; }
    .kpi-label { font-size: 0.7rem; color: #888; font-weight: 700;
                 text-transform: uppercase; letter-spacing: 0.6px; }
    .kpi-value { font-size: 1.45rem; font-weight: 800; color: #1a2744; margin-top: 6px; }
    .kpi-sub   { font-size: 0.72rem; color: #aaa; margin-top: 3px; }
    .sec-title {
        font-size: 0.85rem; font-weight: 800; color: #1a2744;
        text-transform: uppercase; letter-spacing: 0.8px;
        border-left: 4px solid #e95d0c; padding-left: 10px;
        margin: 20px 0 14px 0;
    }
    .info-corte {
        background: #e8f4fd; border: 1px solid #aed6f1;
        border-radius: 8px; padding: 8px 16px;
        font-size: 0.8rem; color: #1a5276; margin-bottom: 12px;
    }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────
# UTILITÁRIOS
# ─────────────────────────────────────────────────────────────
def fmt_brl(valor):
    if pd.isna(valor) or valor is None:
        return "R$ 0,00"
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_pct(valor):
    if pd.isna(valor):
        return "0,0%"
    return f"{valor:.1f}%".replace(".", ",")


def get_file_mtime():
    try:
        return os.path.getmtime(CAMINHO_EXCEL)
    except Exception:
        return 0


# ─────────────────────────────────────────────────────────────
# AUTO-RELOAD
# ─────────────────────────────────────────────────────────────
if "last_mtime" not in st.session_state:
    st.session_state.last_mtime = get_file_mtime()

current_mtime = get_file_mtime()
if current_mtime != st.session_state.last_mtime:
    st.session_state.last_mtime = current_mtime
    st.cache_data.clear()
    st.rerun()


# ─────────────────────────────────────────────────────────────
# CARREGAMENTO E TRATAMENTO DOS DADOS
# ─────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="⏳ Carregando dados...")
def load_data(mtime: float):
    """
    Carrega e trata todos os dados.
    mtime: força invalidação do cache quando o arquivo muda.

    REGRA DE FATURAMENTO:
      1. Coluna D somente quando coluna C = 'Considerar'
      2. Apenas ano 2026
      3. Corte automático pelo último mês COMPLETO de despesas
         (mês incompleto = menos de 60% da média dos meses anteriores)
    """
    if not CAMINHO_EXCEL.exists():
        return None, None, None, None, None, False, f"Arquivo não encontrado: {CAMINHO_EXCEL}"

    try:
        xl = pd.ExcelFile(CAMINHO_EXCEL)
    except Exception as e:
        return None, None, None, None, None, False, str(e)

    # ── 1. Despesas ───────────────────────────────────────────
    try:
        df_desp = pd.read_excel(xl, sheet_name="VW_RelatorioViagensDesoesasCola")
        df_desp.columns = [c.strip() for c in df_desp.columns]
        df_desp.rename(columns={
            "DT_DESPESA":          "data",
            "NR_VIAGEM":           "nr_viagem",
            "CD_ENTIDADE_FUNC":    "cod_colaborador",
            "NM_RAZAOSOC":         "colaborador",
            "DS_CENTROCUSTO":      "centro_custo",
            "DS_TIPO_DESP_VIAGEM": "tipo_despesa",
            "VR_DESPESA":          "valor_despesa",
            "FG_SITUACAO":         "situacao",
        }, inplace=True)
        df_desp["data"]          = pd.to_datetime(df_desp["data"], errors="coerce")
        df_desp["valor_despesa"] = pd.to_numeric(df_desp.get("valor_despesa", 0), errors="coerce").fillna(0)
        df_desp["ano"]           = df_desp["data"].dt.year
        df_desp["mes"]           = df_desp["data"].dt.month
        df_desp["mes_ano"]       = df_desp["data"].dt.to_period("M").astype(str)
        # Somente 2026
        df_desp = df_desp[df_desp["ano"] == ANO_DASHBOARD].copy()
    except Exception as e:
        return None, None, None, None, None, False, f"Erro aba VW_Relatorio: {e}"

    # ── 2. Colaborador → canal ────────────────────────────────
    try:
        df_colab = pd.read_excel(xl, sheet_name="dimColaborador")
        df_colab.columns = [c.strip() for c in df_colab.columns]
        df_colab.rename(columns={
            "Cód. Colaborador": "cod_colaborador",
            "CANAL":            "canal",
            "SUB-REGIÃO":       "sub_regiao_colab",
        }, inplace=True)
        df_colab = df_colab[["cod_colaborador", "canal", "sub_regiao_colab"]].drop_duplicates()
        df_desp  = df_desp.merge(df_colab, on="cod_colaborador", how="left")
    except Exception:
        df_desp["canal"]          = "N/I"
        df_desp["sub_regiao_colab"] = "N/I"

    # ── 3. Corte automático de meses ──────────────────────────
    #  Último mês com < 60% da média dos anteriores = incompleto
    try:
        contagem = df_desp.groupby("mes")["valor_despesa"].count()
        meses    = sorted(contagem.index.tolist())
        if len(meses) >= 2:
            media_ant        = contagem.iloc[:-1].mean()
            ultimo_mes       = meses[-1]
            incompleto       = contagem[ultimo_mes] < media_ant * 0.60
            mes_corte        = meses[-2] if incompleto else ultimo_mes
            mes_incompleto   = incompleto
        else:
            mes_corte      = meses[-1] if meses else 12
            mes_incompleto = False
    except Exception:
        mes_corte      = 3
        mes_incompleto = False

    # Aplica corte nas despesas
    df_desp = df_desp[df_desp["mes"] <= mes_corte].copy()

    # ── 4. Orçamento ──────────────────────────────────────────
    try:
        df_orc = pd.read_excel(xl, sheet_name="FatoOrçamento")
        df_orc.columns = [c.strip() for c in df_orc.columns]
        df_orc.rename(columns={
            "Ano":          "ano",
            "Mês":          "mes",
            "Vl_Planejado": "valor_orcado",
        }, inplace=True)
        df_orc["valor_orcado"] = pd.to_numeric(df_orc["valor_orcado"], errors="coerce").fillna(0)
        df_orc["mes_ano"]      = pd.to_datetime(
            df_orc["ano"].astype(str) + "-" + df_orc["mes"].astype(str).str.zfill(2) + "-01"
        ).dt.to_period("M").astype(str)
        df_orc = df_orc[
            (df_orc["ano"] == ANO_DASHBOARD) & (df_orc["mes"] <= mes_corte)
        ].copy()
    except Exception:
        df_orc = pd.DataFrame(columns=["ano", "mes", "mes_ano", "valor_orcado"])

    # ── 5. Faturamento ────────────────────────────────────────
    #   REGRA: coluna D somente quando coluna C = "Considerar"
    #          ano 2026, meses 1 até mes_corte
    try:
        df_fat = pd.read_excel(xl, sheet_name="Faturamento")
        df_fat.columns = ["data", "sub_regiao", "considerar", "valor_faturamento"]
        df_fat["data"]              = pd.to_datetime(df_fat["data"], errors="coerce")
        df_fat["valor_faturamento"] = pd.to_numeric(df_fat["valor_faturamento"], errors="coerce").fillna(0)

        # FILTRO PRINCIPAL: somente "Considerar"
        df_fat = df_fat[df_fat["considerar"] == "Considerar"].copy()

        df_fat["ano"]     = df_fat["data"].dt.year
        df_fat["mes"]     = df_fat["data"].dt.month
        df_fat["mes_ano"] = df_fat["data"].dt.to_period("M").astype(str)

        # Filtro: 2026 + até mes_corte
        df_fat = df_fat[
            (df_fat["ano"] == ANO_DASHBOARD) &
            (df_fat["mes"] <= mes_corte)
        ].copy()

        # Canal por sub-região
        def sub_to_canal(s):
            if pd.isna(s): return "OUTROS"
            s = str(s).upper()
            if "GC" in s or "GRANDE" in s:     return "GRANDES CONTAS"
            if "DR" in s or "PRIME" in s:       return "DISTRIBUIDOR"
            if "CF" in s or "PEQUENO" in s:     return "CF - PEQUENO PORTE"
            if "OP" in s or "OPERADORA" in s:   return "OPERADORAS DE SAÚDE"
            if "LPC" in s:                      return "LPC"
            return s

        df_fat["canal"] = df_fat["sub_regiao"].apply(sub_to_canal)
    except Exception as e:
        df_fat         = pd.DataFrame(columns=["data","sub_regiao","canal",
                                                "valor_faturamento","ano","mes","mes_ano"])
        mes_incompleto = False

    corte_label = f"{MESES_NOMES.get(mes_corte, str(mes_corte))}/{ANO_DASHBOARD}"
    return df_desp, df_orc, df_fat, mes_corte, corte_label, mes_incompleto, None


# ─────────────────────────────────────────────────────────────
# CARREGA OS DADOS
# ─────────────────────────────────────────────────────────────
df_desp, df_orc, df_fat, mes_corte, corte_label, mes_incompleto, erro = load_data(
    st.session_state.last_mtime
)

# ─────────────────────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────────────────────
ultima_mod = ""
try:
    ts = os.path.getmtime(CAMINHO_EXCEL)
    ultima_mod = time.strftime("%d/%m/%Y %H:%M", time.localtime(ts))
except Exception:
    pass

st.markdown(f"""
<div class="header-bar">
    <div class="header-logo">🔬</div>
    <div class="header-title">FATURAMENTO X DESPESA REALIZADA — {ANO_DASHBOARD}</div>
    <div class="header-badge">
        🔄 Auto-reload ativo &nbsp;|&nbsp;
        Dados até: <b>{corte_label if not erro else "—"}</b> &nbsp;|&nbsp;
        Arquivo: {ultima_mod}
    </div>
</div>
""", unsafe_allow_html=True)

if erro:
    st.error(f"❌ {erro}")
    st.info(f"📁 Caminho esperado: `{CAMINHO_EXCEL}`")
    st.stop()

# Aviso mês incompleto
if mes_incompleto:
    prox = mes_corte + 1
    MESES_FULL = {1:"Janeiro",2:"Fevereiro",3:"Março",4:"Abril",5:"Maio",6:"Junho",
                  7:"Julho",8:"Agosto",9:"Setembro",10:"Outubro",11:"Novembro",12:"Dezembro"}
    st.markdown(f"""
    <div class="info-corte">
        ℹ️ <b>{MESES_FULL.get(prox, str(prox))}</b> ainda não fechou —
        dashboard exibindo dados até <b>{corte_label}</b>.
        Quando o mês fechar e a planilha for atualizada, os dados serão incluídos automaticamente.
    </div>
    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# NAVEGAÇÃO
# ─────────────────────────────────────────────────────────────
if "pagina" not in st.session_state:
    st.session_state.pagina = "faturamento"

col_nav1, col_nav2, col_nav3 = st.columns([2, 2, 8])
with col_nav1:
    if st.button("📋 Painel de Despesas", use_container_width=True,
                 type="primary" if st.session_state.pagina == "despesas" else "secondary"):
        st.session_state.pagina = "despesas"
        st.rerun()
with col_nav2:
    if st.button("📈 Faturamento x Despesa", use_container_width=True,
                 type="primary" if st.session_state.pagina == "faturamento" else "secondary"):
        st.session_state.pagina = "faturamento"
        st.rerun()

st.markdown("---")

# ─────────────────────────────────────────────────────────────
# FILTROS
# ─────────────────────────────────────────────────────────────
st.markdown('<div class="sec-title">Filtros</div>', unsafe_allow_html=True)
fc1, fc2, fc3, fc4 = st.columns(4)

meses_disp  = sorted(df_desp["mes"].dropna().unique().astype(int).tolist())
meses_opts  = ["Todos"] + [f"{MESES_NOMES[m]}/{ANO_DASHBOARD}" for m in meses_disp]
tipos_disp  = ["Todos"] + sorted(df_desp["tipo_despesa"].dropna().unique().tolist())
colabs_disp = ["Todos"] + sorted(df_desp["colaborador"].dropna().unique().tolist())
canais_disp = ["Todos"] + sorted(df_desp["canal"].dropna().unique().tolist())

with fc1: f_mes   = st.selectbox("📅 Mês",             meses_opts)
with fc2: f_tipo  = st.selectbox("📂 Tipo de Despesa", tipos_disp)
with fc3: f_colab = st.selectbox("👤 Colaborador",     colabs_disp)
with fc4: f_canal = st.selectbox("📡 Canal",           canais_disp)


def mes_from_label(label):
    """Converte 'Mar/2026' → 3"""
    if label == "Todos":
        return None
    return MESES_INV.get(label.split("/")[0], None)


def filtrar_desp(df):
    d = df.copy()
    m = mes_from_label(f_mes)
    if m:
        d = d[d["mes"] == m]
    if f_tipo != "Todos":
        d = d[d["tipo_despesa"] == f_tipo]
    if f_colab != "Todos":
        d = d[d["colaborador"] == f_colab]
    if f_canal != "Todos" and "canal" in d.columns:
        d = d[d["canal"] == f_canal]
    return d


def filtrar_fat(df):
    d = df.copy()
    m = mes_from_label(f_mes)
    if m:
        d = d[d["mes"] == m]
    if f_canal != "Todos" and "canal" in d.columns:
        d = d[d["canal"] == f_canal]
    return d


def filtrar_orc(df):
    d = df.copy()
    m = mes_from_label(f_mes)
    if m:
        d = d[d["mes"] == m]
    return d


desp_f = filtrar_desp(df_desp)
fat_f  = filtrar_fat(df_fat)
orc_f  = filtrar_orc(df_orc)

# ─────────────────────────────────────────────────────────────
# KPIs
# ─────────────────────────────────────────────────────────────
total_despesa     = desp_f["valor_despesa"].sum()
total_orcado      = orc_f["valor_orcado"].sum()
pct_realizado     = (total_despesa / total_orcado * 100) if total_orcado > 0 else 0
total_faturamento = fat_f["valor_faturamento"].sum()
n_colaboradores   = desp_f["colaborador"].nunique()
media_por_colab   = total_despesa / n_colaboradores if n_colaboradores > 0 else 0
relacao_fat_desp  = total_faturamento / total_despesa if total_despesa > 0 else 0


# ═══════════════════════════════════════════════════════════
# PÁGINA 1: PAINEL DE DESPESAS
# ═══════════════════════════════════════════════════════════
if st.session_state.pagina == "despesas":

    st.markdown('<div class="sec-title">📊 KPIs — Despesas de Viagem</div>', unsafe_allow_html=True)

    k1, k2, k3, k4, k5 = st.columns(5)
    with k1:
        st.markdown(f"""<div class="kpi-card">
            <div class="kpi-label">Orçamento até {corte_label}</div>
            <div class="kpi-value">{fmt_brl(total_orcado)}</div>
            <div class="kpi-sub">Valor planejado no período</div>
        </div>""", unsafe_allow_html=True)
    with k2:
        st.markdown(f"""<div class="kpi-card laranja">
            <div class="kpi-label">Despesa Realizada</div>
            <div class="kpi-value">{fmt_brl(total_despesa)}</div>
            <div class="kpi-sub">{n_colaboradores} colaboradores</div>
        </div>""", unsafe_allow_html=True)
    with k3:
        cor = "verde" if pct_realizado <= 100 else "vermelho"
        st.markdown(f"""<div class="kpi-card {cor}">
            <div class="kpi-label">% Realizado</div>
            <div class="kpi-value">{fmt_pct(pct_realizado)}</div>
            <div class="kpi-sub">Despesa / Orçamento</div>
        </div>""", unsafe_allow_html=True)
    with k4:
        st.markdown(f"""<div class="kpi-card azul">
            <div class="kpi-label">Média por Colaborador</div>
            <div class="kpi-value">{fmt_brl(media_por_colab)}</div>
            <div class="kpi-sub">{n_colaboradores} colaboradores ativos</div>
        </div>""", unsafe_allow_html=True)
    with k5:
        saldo = total_orcado - total_despesa
        st.markdown(f"""<div class="kpi-card {'verde' if saldo >= 0 else 'vermelho'}">
            <div class="kpi-label">Saldo Orçamentário</div>
            <div class="kpi-value">{fmt_brl(saldo)}</div>
            <div class="kpi-sub">Orçamento − Realizado</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Gráficos de barras horizontais ────────────────────────
    st.markdown('<div class="sec-title">🗂 Despesas de Viagem</div>', unsafe_allow_html=True)
    gcol1, gcol2, gcol3 = st.columns(3)

    def bar_h(df_g, col_cat, col_val, titulo, top_n=10):
        if df_g.empty:
            return go.Figure()
        d = df_g.groupby(col_cat)[col_val].sum().nlargest(top_n).reset_index()
        d = d.sort_values(col_val, ascending=True)
        labels = [str(x)[:24]+"..." if len(str(x))>24 else str(x) for x in d[col_cat]]
        fig = go.Figure(go.Bar(
            x=d[col_val], y=labels, orientation="h",
            marker_color="#1a2744",
            text=[fmt_brl(v) for v in d[col_val]],
            textposition="outside",
            textfont=dict(size=10, color="#e95d0c", family="Arial Black"),
        ))
        fig.update_layout(
            title=dict(text=titulo, font=dict(size=12, color="#1a2744"), x=0),
            height=360, margin=dict(l=0, r=110, t=36, b=10),
            xaxis=dict(showticklabels=False, showgrid=False, zeroline=False),
            yaxis=dict(tickfont=dict(size=10)),
            paper_bgcolor="white", plot_bgcolor="white",
        )
        return fig

    with gcol1:
        st.plotly_chart(bar_h(desp_f, "tipo_despesa", "valor_despesa", "Despesa por Tipo"),
                        use_container_width=True)
    with gcol2:
        st.plotly_chart(bar_h(desp_f, "colaborador", "valor_despesa", "Despesa por Colaborador"),
                        use_container_width=True)
    with gcol3:
        st.plotly_chart(bar_h(desp_f, "canal", "valor_despesa", "Despesa por Canal"),
                        use_container_width=True)

    # ── Orçamento x Realizado mensal ──────────────────────────
    st.markdown('<div class="sec-title">📅 Orçamento x Realizado — Mensal</div>', unsafe_allow_html=True)

    real_m = desp_f.groupby("mes_ano")["valor_despesa"].sum().reset_index()
    orc_m  = df_orc.groupby("mes_ano")["valor_orcado"].sum().reset_index()
    df_mensal = pd.merge(
        orc_m.rename(columns={"mes_ano":"periodo"}),
        real_m.rename(columns={"mes_ano":"periodo","valor_despesa":"realizado"}),
        on="periodo", how="outer"
    ).fillna(0).sort_values("periodo")

    if not df_mensal.empty:
        fig_m = go.Figure()
        fig_m.add_trace(go.Bar(
            name="Orçado", x=df_mensal["periodo"], y=df_mensal["valor_orcado"],
            marker_color="#a0b4d8",
            text=[fmt_brl(v) for v in df_mensal["valor_orcado"]],
            textposition="outside", textfont=dict(size=9),
        ))
        fig_m.add_trace(go.Bar(
            name="Realizado", x=df_mensal["periodo"], y=df_mensal["realizado"],
            marker_color="#e95d0c",
            text=[fmt_brl(v) for v in df_mensal["realizado"]],
            textposition="outside", textfont=dict(size=9),
        ))
        fig_m.update_layout(
            barmode="group", height=340, margin=dict(l=0,r=0,t=20,b=10),
            paper_bgcolor="white", plot_bgcolor="white",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            yaxis=dict(showgrid=True, gridcolor="#eee"),
            xaxis=dict(tickangle=-30),
        )
        st.plotly_chart(fig_m, use_container_width=True)

    # ── Acumulado ─────────────────────────────────────────────
    st.markdown('<div class="sec-title">📈 Evolução Acumulada</div>', unsafe_allow_html=True)
    if not df_mensal.empty:
        df_mensal["acum_orc"]  = df_mensal["valor_orcado"].cumsum()
        df_mensal["acum_real"] = df_mensal["realizado"].cumsum()
        fig_a = go.Figure()
        fig_a.add_trace(go.Scatter(
            name="Orçado Acumulado", x=df_mensal["periodo"], y=df_mensal["acum_orc"],
            line=dict(color="#1a2744", width=3, dash="dot"),
            mode="lines+markers", marker=dict(size=7),
        ))
        fig_a.add_trace(go.Scatter(
            name="Realizado Acumulado", x=df_mensal["periodo"], y=df_mensal["acum_real"],
            line=dict(color="#e95d0c", width=3),
            mode="lines+markers",
            fill="tozeroy", fillcolor="rgba(233,93,12,0.07)",
        ))
        fig_a.update_layout(
            height=300, margin=dict(l=0,r=0,t=20,b=10),
            paper_bgcolor="white", plot_bgcolor="white",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            yaxis=dict(showgrid=True, gridcolor="#eee"),
        )
        st.plotly_chart(fig_a, use_container_width=True)

    # ── Tabela analítica ──────────────────────────────────────
    st.markdown('<div class="sec-title">🔎 Tabela Analítica</div>', unsafe_allow_html=True)
    if not desp_f.empty:
        cols = [c for c in ["centro_custo","colaborador","tipo_despesa",
                             "nr_viagem","data","valor_despesa"] if c in desp_f.columns]
        tab = desp_f[cols].copy()
        tab.rename(columns={"centro_custo":"Centro de Custo","colaborador":"Colaborador",
                             "tipo_despesa":"Tipo de Despesa","nr_viagem":"Nº Viagem",
                             "data":"Data","valor_despesa":"Valor (R$)"}, inplace=True)
        if "Data" in tab.columns:
            tab["Data"] = tab["Data"].dt.strftime("%d/%m/%Y")
        if "Valor (R$)" in tab.columns:
            tab["Valor (R$)"] = tab["Valor (R$)"].apply(fmt_brl)
        tab = tab.sort_values(["Centro de Custo","Colaborador","Tipo de Despesa"])
        st.dataframe(tab, use_container_width=True, height=380)
    else:
        st.info("Sem registros para os filtros selecionados.")


# ═══════════════════════════════════════════════════════════
# PÁGINA 2: FATURAMENTO X DESPESA REALIZADA
# ═══════════════════════════════════════════════════════════
else:

    st.markdown('<div class="sec-title">📊 KPIs — Faturamento x Despesa</div>', unsafe_allow_html=True)

    k1, k2, k3, k4, k5 = st.columns(5)
    with k1:
        st.markdown(f"""<div class="kpi-card azul">
            <div class="kpi-label">Faturamento {ANO_DASHBOARD}</div>
            <div class="kpi-value">{fmt_brl(total_faturamento)}</div>
            <div class="kpi-sub">Col C = "Considerar" | até {corte_label}</div>
        </div>""", unsafe_allow_html=True)
    with k2:
        st.markdown(f"""<div class="kpi-card laranja">
            <div class="kpi-label">Despesa Realizada</div>
            <div class="kpi-value">{fmt_brl(total_despesa)}</div>
            <div class="kpi-sub">Viagens até {corte_label}</div>
        </div>""", unsafe_allow_html=True)
    with k3:
        pct_df = (total_despesa / total_faturamento * 100) if total_faturamento > 0 else 0
        cor3   = "verde" if pct_df <= 5 else ("laranja" if pct_df <= 10 else "vermelho")
        st.markdown(f"""<div class="kpi-card {cor3}">
            <div class="kpi-label">Despesa / Faturamento</div>
            <div class="kpi-value">{fmt_pct(pct_df)}</div>
            <div class="kpi-sub">% da receita em viagens</div>
        </div>""", unsafe_allow_html=True)
    with k4:
        st.markdown(f"""<div class="kpi-card verde">
            <div class="kpi-label">Relação Fat / Desp</div>
            <div class="kpi-value">{relacao_fat_desp:.1f}x</div>
            <div class="kpi-sub">R$ faturados por R$ gasto</div>
        </div>""", unsafe_allow_html=True)
    with k5:
        st.markdown(f"""<div class="kpi-card">
            <div class="kpi-label">Orçamento Despesas</div>
            <div class="kpi-value">{fmt_brl(total_orcado)}</div>
            <div class="kpi-sub">% Realizado: {fmt_pct(pct_realizado)}</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Gráficos ──────────────────────────────────────────────
    gc1, gc2 = st.columns([3, 2])

    with gc1:
        st.markdown('<div class="sec-title">🏷 Faturamento por Sub-Região</div>', unsafe_allow_html=True)
        if not fat_f.empty:
            fs = fat_f.groupby("sub_regiao")["valor_faturamento"].sum().nlargest(14).reset_index()
            fs = fs.sort_values("valor_faturamento", ascending=True)
            labels = [str(x)[:32]+"..." if len(str(x))>32 else str(x) for x in fs["sub_regiao"]]
            fig_sub = go.Figure(go.Bar(
                x=fs["valor_faturamento"], y=labels,
                orientation="h", marker_color="#e95d0c",
                text=[fmt_brl(v) for v in fs["valor_faturamento"]],
                textposition="outside",
                textfont=dict(size=10, color="#1a2744", family="Arial Black"),
            ))
            fig_sub.update_layout(
                height=460, margin=dict(l=0, r=120, t=10, b=10),
                xaxis=dict(showticklabels=False, showgrid=False, zeroline=False),
                yaxis=dict(tickfont=dict(size=10)),
                paper_bgcolor="white", plot_bgcolor="white",
            )
            st.plotly_chart(fig_sub, use_container_width=True)
        else:
            st.info("Sem dados de faturamento.")

    with gc2:
        st.markdown('<div class="sec-title">📡 Faturamento por Canal</div>', unsafe_allow_html=True)
        if not fat_f.empty:
            fc_data = fat_f.groupby("canal")["valor_faturamento"].sum().reset_index()
            fig_p = go.Figure(go.Pie(
                labels=fc_data["canal"],
                values=fc_data["valor_faturamento"],
                hole=0.45,
                marker=dict(colors=["#1a2744","#e95d0c","#2980b9","#27ae60","#8e44ad","#f39c12"]),
                textinfo="label+percent", textfont=dict(size=10),
            ))
            fig_p.update_layout(
                height=460, margin=dict(l=0,r=0,t=10,b=10),
                paper_bgcolor="white", showlegend=False,
            )
            st.plotly_chart(fig_p, use_container_width=True)

    # ── Mensal comparativo ────────────────────────────────────
    st.markdown('<div class="sec-title">📅 Faturamento x Despesa — Mensal</div>', unsafe_allow_html=True)

    fat_m  = fat_f.groupby("mes_ano")["valor_faturamento"].sum().reset_index()
    desp_m = desp_f.groupby("mes_ano")["valor_despesa"].sum().reset_index()
    df_comp = pd.merge(
        fat_m.rename(columns={"mes_ano":"periodo"}),
        desp_m.rename(columns={"mes_ano":"periodo","valor_despesa":"despesa"}),
        on="periodo", how="outer"
    ).fillna(0).sort_values("periodo")

    if not df_comp.empty:
        fig_c = go.Figure()
        fig_c.add_trace(go.Bar(
            name="Faturamento", x=df_comp["periodo"], y=df_comp["valor_faturamento"],
            marker_color="#1a2744",
            text=[fmt_brl(v) for v in df_comp["valor_faturamento"]],
            textposition="outside", textfont=dict(size=9),
        ))
        fig_c.add_trace(go.Scatter(
            name="Despesa", x=df_comp["periodo"], y=df_comp["despesa"],
            line=dict(color="#e95d0c", width=3),
            mode="lines+markers+text",
            text=[fmt_brl(v) for v in df_comp["despesa"]],
            textposition="top center", textfont=dict(size=9, color="#e95d0c"),
            yaxis="y2",
        ))
        fig_c.update_layout(
            height=360, margin=dict(l=0,r=80,t=30,b=10),
            paper_bgcolor="white", plot_bgcolor="white",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            yaxis=dict(title="Faturamento (R$)", showgrid=True, gridcolor="#eee"),
            yaxis2=dict(title="Despesa (R$)", overlaying="y", side="right", showgrid=False),
            xaxis=dict(tickangle=-30),
        )
        st.plotly_chart(fig_c, use_container_width=True)

    # ── Tabelas cruzadas ──────────────────────────────────────
    st.markdown(
        f'<div class="sec-title">🔎 Faturamento x Despesa — por Sub-Região / Canal</div>',
        unsafe_allow_html=True
    )

    col_t1, col_t2 = st.columns(2)
    with col_t1:
        st.caption(f"**Faturamento por Sub-Região** — Jan a {corte_label}")
        if not fat_f.empty:
            t = fat_f.groupby("sub_regiao")["valor_faturamento"].sum().reset_index()
            t = t.sort_values("valor_faturamento", ascending=False)
            t.columns = ["Sub-Região", "Faturamento"]
            total = pd.DataFrame([{"Sub-Região":"TOTAL",
                                    "Faturamento": fat_f["valor_faturamento"].sum()}])
            t = pd.concat([t, total], ignore_index=True)
            t["Faturamento"] = t["Faturamento"].apply(fmt_brl)
            st.dataframe(t, use_container_width=True, height=420)

    with col_t2:
        st.caption(f"**Despesa por Canal** — Jan a {corte_label}")
        if not desp_f.empty:
            t2 = desp_f.groupby("canal")["valor_despesa"].sum().reset_index()
            t2 = t2.sort_values("valor_despesa", ascending=False)
            t2.columns = ["Canal", "Despesa"]
            total2 = pd.DataFrame([{"Canal":"TOTAL",
                                     "Despesa": desp_f["valor_despesa"].sum()}])
            t2 = pd.concat([t2, total2], ignore_index=True)
            t2["Despesa"] = t2["Despesa"].apply(fmt_brl)
            st.dataframe(t2, use_container_width=True, height=420)


# ─────────────────────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────────────────────
st.markdown("---")
col_f1, col_f2 = st.columns([4, 1])
with col_f1:
    st.caption(
        f"📁 `{CAMINHO_EXCEL}` · Modificado: **{ultima_mod}** · "
        f"Período: Jan–{corte_label} · "
        f"Faturamento: somente coluna C = 'Considerar' · Ano: {ANO_DASHBOARD}"
    )
with col_f2:
    if st.button("🔄 Forçar recarga", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
