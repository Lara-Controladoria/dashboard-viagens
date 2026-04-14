"""
Dashboard: Despesas de Viagens - FIRSTLAB
Fonte: Base_2.xlsx

REGRAS DE FATURAMENTO:
  - Coluna D (valor) SOMENTE quando coluna C = "Considerar"
  - Ano fixo: 2026
  - Corte automático: último mês COMPLETO de despesas
    (mês com < 60% da média dos anteriores = incompleto)
"""

import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
import os, re, base64, time
from pathlib import Path

st.set_page_config(
    page_title="Despesas de Viagens - FIRSTLAB",
    page_icon="✈️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Cores FIRSTLAB ────────────────────────────
AZUL_ESCURO  = "#0A1F3D"
AZUL_MEDIO   = "#1A3A6B"
LARANJA      = "#F07D00"
CINZA_CLARO  = "#F4F6FA"
BRANCO       = "#FFFFFF"
TEXTO_ESCURO = "#0A1F3D"

st.markdown(f"""
<style>
    .stApp {{ background-color:{CINZA_CLARO}; }}
    .header-bar {{
        background:{AZUL_ESCURO}; padding:18px 28px; border-radius:10px;
        margin-bottom:18px; display:flex; align-items:center; gap:18px;
    }}
    .header-titles {{ display:flex; flex-direction:column; justify-content:center; }}
    .header-bar h1 {{
        color:white; font-size:1.6rem; font-weight:900;
        margin:0; letter-spacing:1px; line-height:1.2;
    }}
    .header-bar .subtitulo {{
        color:{LARANJA}; font-size:.8rem; font-weight:600;
        margin:0 0 2px 0; letter-spacing:2px; text-transform:uppercase;
    }}
    .header-badge {{
        margin-left:auto; background:rgba(255,255,255,0.12);
        color:#a0b4d8; font-size:.7rem; padding:4px 14px;
        border-radius:20px; border:1px solid rgba(255,255,255,0.2);
        white-space:nowrap;
    }}
    .info-corte {{
        background:#e8f4fd; border:1px solid #aed6f1;
        border-radius:8px; padding:8px 16px;
        font-size:.8rem; color:#1a5276; margin-bottom:12px;
    }}
    .kpi-card {{
        background:{BRANCO}; border-radius:10px; padding:16px 20px;
        box-shadow:0 2px 8px rgba(0,0,0,.08);
        border-left:5px solid {LARANJA}; margin-bottom:8px;
    }}
    .kpi-label {{ font-size:.72rem; color:#666; text-transform:uppercase; font-weight:600; margin-bottom:4px; }}
    .kpi-value {{ font-size:1.25rem; font-weight:800; color:{AZUL_ESCURO}; }}
    .kpi-sub   {{ font-size:.7rem; color:#999; margin-top:2px; }}
    .section-title {{
        font-size:.95rem; font-weight:700; color:{AZUL_ESCURO}; text-transform:uppercase;
        letter-spacing:.5px; border-bottom:2px solid {LARANJA};
        padding-bottom:4px; margin-bottom:12px; margin-top:8px;
    }}
    .stSelectbox label {{ font-weight:600; color:{AZUL_ESCURO}; font-size:.8rem; }}
    .stTabs [data-baseweb="tab-list"] {{ gap:8px; }}
    .stTabs [data-baseweb="tab"] {{
        background:{BRANCO}; border-radius:8px 8px 0 0;
        font-weight:700; color:{AZUL_ESCURO}; border:2px solid #dde;
    }}
    .stTabs [aria-selected="true"] {{
        background:{AZUL_ESCURO} !important;
        color:white !important;
        border-color:{AZUL_ESCURO} !important;
    }}
    footer {{ display:none; }}
    #MainMenu {{ visibility:hidden; }}
</style>
""", unsafe_allow_html=True)

ANO_DASHBOARD = 2026

# ── Caminho do arquivo ────────────────────────
# Tenta todos os nomes possíveis (ponto e underscore)
def encontrar_arquivo():
    base = Path(__file__).parent
    candidatos = [
        base / "Base_2.xlsx",
        base / "Base.2.xlsx",
        base / "Base_2.xlsx",
        Path("Base_2.xlsx"),
        Path("Base.2.xlsx"),
    ]
    for p in candidatos:
        if p.exists():
            return str(p)
    return None

# ── Utilitários ───────────────────────────────
def fmt_brl(v):
    if pd.isna(v) or v == 0: return "R$ 0,00"
    try:
        i, d = f"{abs(v):,.2f}".split(".")
        return f"{'-' if v<0 else ''}R$ {i.replace(',','.')},{d}"
    except: return f"R$ {v:.2f}"

def fmt_pct(v):
    return "0,0%" if pd.isna(v) else f"{v:.1f}%".replace(".",",")

def kpi(col, label, valor, sub=""):
    col.markdown(f"""<div class="kpi-card">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{valor}</div>
        <div class="kpi-sub">{sub}</div>
    </div>""", unsafe_allow_html=True)

def limpar_cc(v):
    if pd.isna(v): return v
    return re.sub(r'^[A-Za-z]{2}\d{2}\s*[-–]\s*', '', str(v)).strip()

def bar_h(df, x, y, title, key, cor=None, height=440):
    if df.empty:
        st.info(f"Sem dados para: {title}")
        return
    fig = px.bar(df, x=x, y=y, orientation="h",
                 text=df[x].apply(fmt_brl),
                 color_discrete_sequence=[cor or AZUL_MEDIO],
                 title=f"<b>{title}</b>")
    fig.update_traces(textposition="outside", textfont_size=11,
                      textfont_color=TEXTO_ESCURO)
    fig.update_layout(height=height, margin=dict(l=200,r=130,t=45,b=10),
                      xaxis_title="", yaxis_title="",
                      plot_bgcolor=BRANCO, paper_bgcolor=BRANCO,
                      title_font=dict(size=13, color=AZUL_ESCURO),
                      font=dict(color=AZUL_ESCURO, size=11),
                      yaxis=dict(tickfont=dict(size=11, color=AZUL_ESCURO), automargin=False))
    fig.update_xaxes(showticklabels=False, showgrid=False, range=[0, df[x].max()*1.35])
    st.plotly_chart(fig, use_container_width=True, key=key)

def graf_barras_mes(df_real, df_orc_m, title, key):
    fig = go.Figure()
    if not df_orc_m.empty:
        fig.add_trace(go.Bar(name="Orçamento", x=df_orc_m["mes_ano"],
                             y=df_orc_m["orcamento"], marker_color=LARANJA,
                             text=[fmt_brl(v) for v in df_orc_m["orcamento"]],
                             textposition="outside", textfont=dict(size=11, color=TEXTO_ESCURO)))
    fig.add_trace(go.Bar(name="Realizado", x=df_real["mes_ano"],
                         y=df_real["valor"], marker_color=AZUL_MEDIO,
                         text=[fmt_brl(v) for v in df_real["valor"]],
                         textposition="outside", textfont=dict(size=11, color=TEXTO_ESCURO)))
    fig.update_layout(barmode="group", title=f"<b>{title}</b>",
                      height=440, plot_bgcolor=BRANCO, paper_bgcolor=BRANCO,
                      title_font=dict(size=14, color=AZUL_ESCURO),
                      font=dict(color=TEXTO_ESCURO, size=12),
                      legend=dict(orientation="h", y=1.1, font=dict(size=12)),
                      margin=dict(l=60,r=20,t=60,b=60),
                      xaxis=dict(tickangle=-35, tickfont=dict(size=11)),
                      yaxis=dict(tickformat=",.0f", showgrid=True, gridcolor="#eee",
                                 tickfont=dict(size=11)))
    st.plotly_chart(fig, use_container_width=True, key=key)

def graf_acumulado(df_real, title, key):
    df = df_real.copy().sort_values("mes_ano_sort")
    df["acumulado"] = df["valor"].cumsum()
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=df["mes_ano"], y=df["acumulado"],
                             mode="lines+markers+text",
                             line=dict(color=AZUL_MEDIO, width=3),
                             marker=dict(size=9, color=LARANJA),
                             text=[fmt_brl(v) for v in df["acumulado"]],
                             textposition="top center", textfont=dict(size=11, color=TEXTO_ESCURO),
                             fill="tozeroy", fillcolor="rgba(26,58,107,0.1)"))
    fig.update_layout(title=f"<b>{title}</b>",
                      height=440, plot_bgcolor=BRANCO, paper_bgcolor=BRANCO,
                      title_font=dict(size=14, color=AZUL_ESCURO),
                      font=dict(color=TEXTO_ESCURO, size=12),
                      margin=dict(l=60,r=20,t=60,b=60),
                      xaxis=dict(tickangle=-35, tickfont=dict(size=11)),
                      yaxis=dict(tickformat=",.0f", showgrid=True, gridcolor="#eee",
                                 tickfont=dict(size=11)))
    st.plotly_chart(fig, use_container_width=True, key=key)

def graf_acumulado_comparativo(df_d, df_o, key):
    if "mes_ano" not in df_d.columns or "valor" not in df_d.columns:
        st.info("Dados insuficientes para o gráfico acumulado.")
        return
    real = (df_d.groupby(["mes_ano_sort","mes_ano"])["valor"].sum()
                .reset_index().sort_values("mes_ano_sort"))
    real["acum_real"] = real["valor"].cumsum()
    orc_ac = pd.DataFrame()
    if "mes_ano" in df_o.columns and "orcamento" in df_o.columns:
        orc_ac = (df_o.groupby(["mes_ano_sort","mes_ano"])["orcamento"].sum()
                      .reset_index().sort_values("mes_ano_sort"))
        orc_ac["acum_orc"] = orc_ac["orcamento"].cumsum()
    if not orc_ac.empty:
        merged = real.merge(orc_ac[["mes_ano_sort","mes_ano","acum_orc"]],
                            on=["mes_ano_sort","mes_ano"], how="left").fillna(0)
    else:
        merged = real.copy()
        merged["acum_orc"] = 0
    merged = merged.sort_values("mes_ano_sort")
    fig = go.Figure()
    if not orc_ac.empty:
        fig.add_trace(go.Scatter(
            x=merged["mes_ano"], y=merged["acum_orc"],
            mode="lines+markers+text", name="Orçado Acumulado",
            line=dict(color=LARANJA, width=3, dash="dot"),
            marker=dict(size=9, color=LARANJA, symbol="diamond"),
            text=[fmt_brl(v) for v in merged["acum_orc"]],
            textposition="bottom center", textfont=dict(size=10, color=LARANJA),
        ))
    fig.add_trace(go.Scatter(
        x=merged["mes_ano"], y=merged["acum_real"],
        mode="lines+markers+text", name="Realizado Acumulado",
        line=dict(color=AZUL_MEDIO, width=3),
        marker=dict(size=9, color=AZUL_MEDIO),
        text=[fmt_brl(v) for v in merged["acum_real"]],
        textposition="top center", textfont=dict(size=10, color=AZUL_ESCURO),
        fill="tozeroy", fillcolor="rgba(26,58,107,0.08)",
    ))
    ultimo = merged.iloc[-1]
    pct  = (ultimo["acum_real"] / ultimo["acum_orc"] * 100) if ultimo.get("acum_orc", 0) > 0 else 0
    diff = ultimo["acum_real"] - ultimo.get("acum_orc", 0)
    cor_diff = "#E74C3C" if diff > 0 else "#27AE60"
    sinal = "acima" if diff > 0 else "abaixo"
    fig.update_layout(
        title=f"<b>Acumulado: Realizado x Orçado</b>  "
              f"<span style='font-size:13px; color:{cor_diff}'>"
              f"{fmt_pct(pct)} utilizado | {fmt_brl(abs(diff))} {sinal} do orçado</span>",
        height=460, plot_bgcolor=BRANCO, paper_bgcolor=BRANCO,
        title_font=dict(size=14, color=AZUL_ESCURO),
        font=dict(color=AZUL_ESCURO, size=12),
        legend=dict(orientation="h", y=1.08, font=dict(size=12)),
        margin=dict(l=70,r=30,t=80,b=60),
        xaxis=dict(tickangle=-35, tickfont=dict(size=11), showgrid=False),
        yaxis=dict(tickformat=",.0f", showgrid=True, gridcolor="#eee",
                   tickfont=dict(size=11), zeroline=False),
        hovermode="x unified",
    )
    st.plotly_chart(fig, use_container_width=True, key=key)

# ── Logo ──────────────────────────────────────
def logo_html():
    base = Path(__file__).parent
    for nome in ["Logo.jpg","logo.jpg","Logo.png","logo.png"]:
        p = base / nome
        if p.exists():
            data = base64.b64encode(p.read_bytes()).decode()
            ext  = "jpg" if nome.endswith("jpg") else "png"
            return f'<img src="data:image/{ext};base64,{data}" style="height:52px; object-fit:contain;">'
    return "<span style='font-size:1.5rem;font-weight:900;letter-spacing:1px;'><span style='color:#F07D00'>FIRST</span><span style='color:white'>LAB</span></span>"

# ── Auto-reload ───────────────────────────────
def get_mtime(arq):
    try: return os.path.getmtime(arq)
    except: return 0

arq = encontrar_arquivo()

if "last_mtime" not in st.session_state:
    st.session_state.last_mtime = get_mtime(arq) if arq else 0

if arq:
    mtime_atual = get_mtime(arq)
    if mtime_atual != st.session_state.last_mtime:
        st.session_state.last_mtime = mtime_atual
        st.cache_data.clear()
        st.rerun()

# ── Carregamento ──────────────────────────────
@st.cache_data(show_spinner="Carregando dados...")
def carregar(caminho, _mtime):
    xl   = pd.ExcelFile(caminho)
    abas = xl.sheet_names

    # De/Para tipo despesa
    depara = {}
    aba = next((a for a in abas if "tipodespesa" in a.lower().replace(" ","")), None)
    if aba:
        d = pd.read_excel(caminho, sheet_name=aba)
        depara = dict(zip(d.iloc[:,0].astype(str).str.strip().str.upper(),
                          d.iloc[:,1].astype(str).str.strip().str.upper()))

    # ── Despesas ──────────────────────────────
    aba = next((a for a in abas if "relatorio" in a.lower() or "viagem" in a.lower()), None)
    df_d = pd.DataFrame()
    if aba:
        df = pd.read_excel(caminho, sheet_name=aba)
        df.columns = [c.strip() for c in df.columns]
        cm = {
            "data":        next((c for c in df.columns if c.startswith("DT_")), None),
            "viagem":      next((c for c in df.columns if "NR_VIAGEM" in c), None),
            "colaborador": next((c for c in df.columns if "NM_RAZAOSOC" in c), None),
            "cc_desc":     next((c for c in df.columns if "DS_CENTROCUSTO" in c), None),
            "tipo_orig":   next((c for c in df.columns if "DS_TIPO_DESP" in c), None),
            "valor":       next((c for c in df.columns if "VR_DESPESA" in c), None),
        }
        df = df.rename(columns={v:k for k,v in cm.items() if v})
        if "data" in df.columns:
            df["data"]         = pd.to_datetime(df["data"], errors="coerce")
            df["ano"]          = df["data"].dt.year
            df["mes"]          = df["data"].dt.month
            df["mes_ano"]      = df["data"].dt.strftime("%m/%Y")
            df["mes_ano_sort"] = df["data"].dt.to_period("M").astype(str)
        if "valor"    in df.columns: df["valor"]    = pd.to_numeric(df["valor"], errors="coerce").fillna(0)
        if "tipo_orig" in df.columns:
            df["tipo_desc"] = (df["tipo_orig"].astype(str).str.strip().str.upper()
                               .map(depara).fillna(df["tipo_orig"].astype(str).str.strip().str.upper()))
        else:
            df["tipo_desc"] = "SEM TIPO"
        if "cc_desc" in df.columns:
            df["cc_desc"] = df["cc_desc"].apply(limpar_cc)

        # Canal via dimColaborador
        aba_c = next((a for a in abas if "colaborador" in a.lower()), None)
        if aba_c:
            dc = pd.read_excel(caminho, sheet_name=aba_c)
            dc.columns = [c.strip() for c in dc.columns]
            cn = next((c for c in dc.columns if "colaborador" in c.lower()), None)
            cc = next((c for c in dc.columns if "canal" in c.lower()), None)
            cod_c = next((c for c in dc.columns if "cód" in c.lower() or "cod" in c.lower()), None)
            cod_d = next((c for c in df.columns if "cd_entidade" in c.lower() or "cod_colaborador" in c.lower()), None)
            if cod_c and cod_d:
                dc = dc.rename(columns={cod_c:"cod_colaborador", cc:"canal"})
                df = df.merge(dc[["cod_colaborador","canal"]].drop_duplicates("cod_colaborador"),
                              left_on=cod_d, right_on="cod_colaborador", how="left")
            elif cn and cc:
                dc = dc.rename(columns={cn:"colaborador", cc:"canal"})
                df = df.merge(dc[["colaborador","canal"]].drop_duplicates("colaborador"),
                              on="colaborador", how="left")
        df_d = df

    # ── Corte automático de meses ──────────────
    # Mês com < 60% da média dos anteriores = incompleto
    mes_corte      = None
    mes_incompleto = False
    if not df_d.empty and "ano" in df_d.columns and "mes" in df_d.columns:
        d2026 = df_d[df_d["ano"] == ANO_DASHBOARD]
        if not d2026.empty:
            contagem = d2026.groupby("mes")["valor"].count()
            meses    = sorted(contagem.index.tolist())
            if len(meses) >= 2:
                media_ant      = contagem.iloc[:-1].mean()
                ultimo         = meses[-1]
                mes_incompleto = contagem[ultimo] < media_ant * 0.60
                mes_corte      = meses[-2] if mes_incompleto else ultimo
            else:
                mes_corte = meses[-1] if meses else 12

    # ── Orçamento ─────────────────────────────
    aba = next((a for a in abas if "fato" in a.lower() and "or" in a.lower()), None) or \
          next((a for a in abas if "orc" in a.lower() or "orç" in a.lower()), None)
    df_o = pd.DataFrame()
    if aba:
        d = pd.read_excel(caminho, sheet_name=aba)
        d.columns = [c.strip() for c in d.columns]
        cm = {
            "ano":       next((c for c in d.columns if c.lower()=="ano"), None),
            "mes":       next((c for c in d.columns if c.lower() in ("mês","mes")), None),
            "tipo_desc": next((c for c in d.columns if "tipo" in c.lower()), None),
            "orcamento": next((c for c in d.columns if "plan" in c.lower() or "vl_" in c.lower()), None),
            "cc_desc":   next((c for c in d.columns if "custo" in c.lower()), None),
        }
        d = d.rename(columns={v:k for k,v in cm.items() if v})
        if "orcamento" in d.columns: d["orcamento"] = pd.to_numeric(d["orcamento"], errors="coerce").fillna(0)
        if "ano" in d.columns and "mes" in d.columns:
            d["mes_ano"]      = d.apply(lambda r: f"{int(r['mes']):02d}/{int(r['ano'])}", axis=1)
            d["mes_ano_sort"] = d.apply(lambda r: f"{int(r['ano'])}-{int(r['mes']):02d}", axis=1)
        if "cc_desc" in d.columns: d["cc_desc"] = d["cc_desc"].apply(limpar_cc)
        df_o = d

    # ── Faturamento ───────────────────────────
    # REGRA: coluna D somente quando coluna C = "Considerar"
    # Ano 2026, meses 1 até mes_corte
    aba = next((a for a in abas if "faturamento" in a.lower()), None)
    df_f = pd.DataFrame()
    if aba:
        d = pd.read_excel(caminho, sheet_name=aba)
        d.columns = [c.strip() for c in d.columns]
        cm = {
            "data":        next((c for c in d.columns if "data" in c.lower() or "emiss" in c.lower()), None),
            "sub_regiao":  next((c for c in d.columns if "sub" in c.lower() or "canal" in c.lower()), None),
            "considerar":  next((c for c in d.columns if "consid" in c.lower()), None),
            "faturamento": next((c for c in d.columns if "valor" in c.lower() or "liquid" in c.lower()), None),
        }
        d = d.rename(columns={v:k for k,v in cm.items() if v})
        if "data" in d.columns:
            d["data"]         = pd.to_datetime(d["data"], errors="coerce")
            d["ano"]          = d["data"].dt.year
            d["mes"]          = d["data"].dt.month
            d["mes_ano"]      = d["data"].dt.strftime("%m/%Y")
            d["mes_ano_sort"] = d["data"].dt.to_period("M").astype(str)
        if "faturamento" in d.columns:
            d["faturamento"] = pd.to_numeric(d["faturamento"], errors="coerce").fillna(0)
        # FILTRO PRINCIPAL: somente "Considerar"
        if "considerar" in d.columns:
            d = d[d["considerar"].astype(str).str.strip().str.lower() == "considerar"].copy()
        df_f = d

    return df_d, df_o, df_f, mes_corte, mes_incompleto


# ── Verificação do arquivo ────────────────────
if not arq:
    st.error("❌ Arquivo Base_2.xlsx não encontrado.")
    st.info(f"📁 Pasta esperada: `{Path(__file__).parent}`")
    st.stop()

df_desp_raw, df_orc_raw, df_fat_raw, mes_corte, mes_incompleto = carregar(
    arq, st.session_state.last_mtime
)

# ── Filtro fixo: apenas 2026 + corte de mês ──
MESES_NOMES = {1:"Jan",2:"Fev",3:"Mar",4:"Abr",5:"Mai",6:"Jun",
               7:"Jul",8:"Ago",9:"Set",10:"Out",11:"Nov",12:"Dez"}
corte_label = f"{MESES_NOMES.get(mes_corte, str(mes_corte))}/2026" if mes_corte else "—"

def aplicar_corte(df, eh_fat=False):
    if df.empty: return df
    d = df.copy()
    if "ano" in d.columns:
        d = d[d["ano"] == ANO_DASHBOARD]
    if mes_corte and "mes" in d.columns:
        d = d[d["mes"] <= mes_corte]
    return d

df_desp = aplicar_corte(df_desp_raw)
df_orc  = aplicar_corte(df_orc_raw)
df_fat  = aplicar_corte(df_fat_raw, eh_fat=True)

# ── Header ────────────────────────────────────
ultima_mod = ""
try:
    ts = os.path.getmtime(arq)
    ultima_mod = time.strftime("%d/%m/%Y %H:%M", time.localtime(ts))
except: pass

st.markdown(f"""
<div class="header-bar">
    {logo_html()}
    <div class="header-titles">
        <h1>DESPESAS DE VIAGENS</h1>
    </div>
    <div class="header-badge">
        🔄 Auto-reload ativo &nbsp;|&nbsp;
        Dados até: <b>{corte_label}</b> &nbsp;|&nbsp;
        {ultima_mod}
    </div>
</div>
""", unsafe_allow_html=True)

# Aviso mês incompleto
if mes_incompleto and mes_corte:
    prox = mes_corte + 1
    MESES_FULL = {1:"Janeiro",2:"Fevereiro",3:"Março",4:"Abril",5:"Maio",6:"Junho",
                  7:"Julho",8:"Agosto",9:"Setembro",10:"Outubro",11:"Novembro",12:"Dezembro"}
    st.markdown(f"""
    <div class="info-corte">
        ℹ️ <b>{MESES_FULL.get(prox, str(prox))}</b> ainda não fechou —
        dashboard exibindo dados até <b>{corte_label}</b>.
        Quando o mês fechar e a planilha for atualizada, será incluído automaticamente.
    </div>
    """, unsafe_allow_html=True)

# ── Filtros compartilhados ────────────────────
def opcoes(df, col):
    return sorted(df[col].dropna().unique().tolist()) if col in df.columns else []

tipos_op  = opcoes(df_desp, "tipo_desc")
colabs_op = opcoes(df_desp, "colaborador")
canais_op = opcoes(df_desp, "canal")
ccs_op    = opcoes(df_desp, "cc_desc")
meses_op  = (df_desp[["mes_ano","mes_ano_sort"]].dropna()
             .drop_duplicates().sort_values("mes_ano_sort")["mes_ano"].tolist()
             if "mes_ano" in df_desp.columns else [])

def filtrar(df_d, df_o, df_f, tipo, colab, mano, canal, cc):
    d,o,f = df_d.copy(), df_o.copy(), df_f.copy()
    if tipo  != "Todos" and "tipo_desc"   in d.columns: d = d[d["tipo_desc"]   == tipo]
    if colab != "Todos" and "colaborador" in d.columns: d = d[d["colaborador"] == colab]
    if canal != "Todos" and "canal"       in d.columns: d = d[d["canal"]       == canal]
    if cc    != "Todos" and "cc_desc"     in d.columns: d = d[d["cc_desc"]     == cc]
    if mano  != "Todos":
        if "mes_ano" in d.columns: d = d[d["mes_ano"] == mano]
        if "mes_ano" in o.columns: o = o[o["mes_ano"] == mano]
        if "mes_ano" in f.columns: f = f[f["mes_ano"] == mano]
    return d, o, f

# ══════════════════════════════════════════════
# BLOCO DE CONTEÚDO (reutilizado nas 2 abas)
# ══════════════════════════════════════════════
def renderizar(df_d, df_o, df_f, sufixo):

    # ── KPIs ──────────────────────────────────
    tot_desp     = df_d["valor"].sum()       if "valor"       in df_d.columns else 0
    tot_orc      = df_o["orcamento"].sum()   if "orcamento"   in df_o.columns else 0
    tot_fat      = df_f["faturamento"].sum() if "faturamento" in df_f.columns else 0
    pct_real     = (tot_desp / tot_orc  * 100) if tot_orc  > 0 else 0
    pct_desp_fat = (tot_desp / tot_fat  * 100) if tot_fat  > 0 else 0
    n_col        = df_d["colaborador"].nunique() if "colaborador" in df_d.columns else 1
    media_col    = tot_desp / n_col if n_col > 0 else 0
    n_viag       = df_d["viagem"].nunique() if "viagem" in df_d.columns else 0

    st.markdown('<div class="section-title">Indicadores</div>', unsafe_allow_html=True)
    c1,c2,c3,c4,c5,c6,c7 = st.columns(7)
    kpi(c1, "Orçamento",           fmt_brl(tot_orc))
    kpi(c2, "Despesa Realizada",   fmt_brl(tot_desp))
    kpi(c3, "% Realizado",         fmt_pct(pct_real),     f"de {fmt_brl(tot_orc)}")
    kpi(c4, "Média / Colaborador", fmt_brl(media_col),    f"{n_col} colaboradores")
    kpi(c5, "Faturamento",         fmt_brl(tot_fat))
    kpi(c6, "Despesa / Fat.",      fmt_pct(pct_desp_fat))
    kpi(c7, "Nº Viagens",          f"{n_viag:,}".replace(",","."))

    st.markdown("---")

    # ── Gráficos despesa ──────────────────────
    st.markdown('<div class="section-title">Despesas de Viagem</div>', unsafe_allow_html=True)
    g1,g2,g3 = st.columns(3)

    with g1:
        if "tipo_desc" in df_d.columns and "valor" in df_d.columns:
            dt = df_d.groupby("tipo_desc")["valor"].sum().reset_index().sort_values("valor",ascending=True).tail(12)
            bar_h(dt,"valor","tipo_desc","Despesa por Tipo", key=f"tipo_{sufixo}")

    with g2:
        # Despesa por colaborador com tipo de despesa (stacked)
        if "colaborador" in df_d.columns and "tipo_desc" in df_d.columns and "valor" in df_d.columns:
            dc = (df_d.groupby(["colaborador","tipo_desc"])["valor"]
                      .sum().reset_index())
            # Top 10 colaboradores por total
            top_colabs = (dc.groupby("colaborador")["valor"].sum()
                            .nlargest(10).index.tolist())
            dc = dc[dc["colaborador"].isin(top_colabs)]
            # Ordenar pelo total
            ordem = (dc.groupby("colaborador")["valor"].sum()
                       .sort_values(ascending=True).index.tolist())
            dc["colaborador"] = pd.Categorical(dc["colaborador"], categories=ordem, ordered=True)
            dc = dc.sort_values("colaborador")
            dc["label"] = dc["colaborador"].astype(str).str[:25]

            fig = px.bar(dc, x="valor", y="label", color="tipo_desc",
                         orientation="h",
                         title="<b>Despesa por Colaborador</b>",
                         text=dc["valor"].apply(fmt_brl),
                         height=440)
            fig.update_traces(textposition="inside", textfont_size=9)
            fig.update_layout(
                barmode="stack",
                margin=dict(l=200,r=20,t=45,b=10),
                xaxis_title="", yaxis_title="",
                plot_bgcolor=BRANCO, paper_bgcolor=BRANCO,
                title_font=dict(size=13, color=AZUL_ESCURO),
                font=dict(color=AZUL_ESCURO, size=11),
                yaxis=dict(tickfont=dict(size=11, color=AZUL_ESCURO), automargin=False),
                legend=dict(title="Tipo", font=dict(size=9),
                            orientation="v", x=1.01, y=1),
            )
            fig.update_xaxes(showticklabels=False, showgrid=False)
            st.plotly_chart(fig, use_container_width=True, key=f"colab_{sufixo}")

    with g3:
        if "canal" in df_d.columns and "valor" in df_d.columns:
            dca = df_d.groupby("canal")["valor"].sum().reset_index().sort_values("valor",ascending=True)
            bar_h(dca,"valor","canal","Despesa por Canal", key=f"canal_{sufixo}")

    st.markdown("---")

    # ── Orçamento x Realizado | Acumulado ─────
    st.markdown('<div class="section-title">Orçamento x Realizado — Visão Mensal</div>', unsafe_allow_html=True)
    m1,m2 = st.columns(2)

    with m1:
        if "mes_ano" in df_d.columns and "valor" in df_d.columns:
            real = (df_d.groupby(["mes_ano_sort","mes_ano"])["valor"].sum()
                        .reset_index().sort_values("mes_ano_sort"))
            orc_m = pd.DataFrame()
            if "mes_ano" in df_o.columns and "orcamento" in df_o.columns:
                orc_m = (df_o.groupby(["mes_ano_sort","mes_ano"])["orcamento"].sum()
                             .reset_index().sort_values("mes_ano_sort"))
            graf_barras_mes(real, orc_m, "Orçamento x Realizado por Mês", key=f"orc_mes_{sufixo}")

    with m2:
        if "mes_ano" in df_d.columns and "valor" in df_d.columns:
            real2 = (df_d.groupby(["mes_ano_sort","mes_ano"])["valor"].sum()
                         .reset_index().sort_values("mes_ano_sort"))
            graf_acumulado(real2, "Despesa Acumulada até a Data", key=f"acum_{sufixo}")

    st.markdown("---")

    # ── Acumulado comparativo ─────────────────
    st.markdown('<div class="section-title">Acumulado Realizado x Orçado</div>', unsafe_allow_html=True)
    graf_acumulado_comparativo(df_d, df_o, key=f"acum_comp_{sufixo}")

    st.markdown("---")

    # ── Faturamento por Mês e Canal ───────────
    st.markdown('<div class="section-title">Faturamento por Mês e Canal</div>', unsafe_allow_html=True)
    if "faturamento" in df_f.columns and "mes_ano" in df_f.columns:
        f1,f2 = st.columns(2)
        with f1:
            fm = (df_f.groupby(["mes_ano_sort","mes_ano"])["faturamento"].sum()
                      .reset_index().sort_values("mes_ano_sort"))
            fig = go.Figure()
            fig.add_trace(go.Bar(x=fm["mes_ano"],y=fm["faturamento"],
                                 marker_color=AZUL_MEDIO,
                                 text=[fmt_brl(v) for v in fm["faturamento"]],
                                 textposition="outside",textfont=dict(size=11, color=TEXTO_ESCURO)))
            fig.update_layout(title="<b>Faturamento por Mês</b>",
                              height=440, plot_bgcolor=BRANCO, paper_bgcolor=BRANCO,
                              title_font=dict(size=14, color=AZUL_ESCURO),
                              font=dict(color=TEXTO_ESCURO, size=12),
                              margin=dict(l=60,r=20,t=60,b=60),
                              xaxis=dict(tickangle=-35, tickfont=dict(size=11)),
                              yaxis=dict(tickformat=",.0f", showgrid=True, gridcolor="#eee",
                                         tickfont=dict(size=11)),
                              xaxis_title="", yaxis_title="")
            st.plotly_chart(fig, use_container_width=True, key=f"fat_mes_{sufixo}")
        with f2:
            if "sub_regiao" in df_f.columns:
                fc = (df_f.groupby("sub_regiao")["faturamento"].sum()
                          .reset_index().sort_values("faturamento",ascending=True).tail(15))
                fc["label"] = fc["sub_regiao"].str[:30]
                fig2 = px.bar(fc,x="faturamento",y="label",orientation="h",
                              text=fc["faturamento"].apply(fmt_brl),
                              color_discrete_sequence=[LARANJA],
                              title="<b>Faturamento por Sub-Região</b>")
                fig2.update_traces(textposition="outside", textfont_size=11,
                                  textfont_color=TEXTO_ESCURO)
                fig2.update_layout(height=440, margin=dict(l=230,r=130,t=45,b=10),
                                   xaxis_title="", yaxis_title="",
                                   plot_bgcolor=BRANCO, paper_bgcolor=BRANCO,
                                   title_font=dict(size=13, color=AZUL_ESCURO),
                                   font=dict(color=AZUL_ESCURO, size=11),
                                   yaxis=dict(tickfont=dict(size=11, color=AZUL_ESCURO), automargin=False))
                fig2.update_xaxes(showticklabels=False, showgrid=False,
                                  range=[0, fc["faturamento"].max()*1.35])
                st.plotly_chart(fig2, use_container_width=True, key=f"fat_canal_{sufixo}")

    st.markdown("---")

    # ── Tabela Canal x Faturamento x Despesa ──
    st.markdown('<div class="section-title">Faturamento x Despesa Realizada — por Canal</div>',
                unsafe_allow_html=True)
    if "canal" in df_d.columns and "valor" in df_d.columns:
        td = df_d.groupby("canal")["valor"].sum().reset_index().rename(columns={"canal":"Canal","valor":"Despesa"})
        tf = pd.DataFrame()
        if "sub_regiao" in df_f.columns and "faturamento" in df_f.columns:
            tf = df_f.groupby("sub_regiao")["faturamento"].sum().reset_index().rename(
                    columns={"sub_regiao":"Canal","faturamento":"Faturamento"})
        tab = td.merge(tf, on="Canal", how="outer").fillna(0) if not tf.empty else td.assign(Faturamento=0)
        tot = pd.DataFrame({"Canal":["TOTAL"],"Despesa":[tab["Despesa"].sum()],"Faturamento":[tab["Faturamento"].sum()]})
        tab = pd.concat([tab,tot],ignore_index=True)
        tab["Despesa"]     = tab["Despesa"].apply(fmt_brl)
        tab["Faturamento"] = tab["Faturamento"].apply(fmt_brl)
        st.dataframe(tab, use_container_width=True, hide_index=True)

    st.markdown("---")

    # ── Tabela hierárquica ────────────────────
    st.markdown('<div class="section-title">Detalhamento: Centro de Custo → Colaborador → Tipo → Viagem</div>',
                unsafe_allow_html=True)
    cols_h = [c for c in ["cc_desc","colaborador","tipo_desc","viagem"] if c in df_d.columns]
    if cols_h and "valor" in df_d.columns:
        dh = df_d.groupby(cols_h)["valor"].sum().reset_index().sort_values("valor",ascending=False)
        dh["valor_fmt"] = dh["valor"].apply(fmt_brl)
        rm = {"cc_desc":"Centro de Custo","colaborador":"Colaborador",
              "tipo_desc":"Tipo de Despesa","viagem":"Nº Viagem","valor_fmt":"Valor"}
        dh = dh.rename(columns=rm)[[rm[c] for c in cols_h]+["Valor"]]
        st.dataframe(dh, use_container_width=True, hide_index=True, height=420)

# ══════════════════════════════════════════════
# ABAS
# ══════════════════════════════════════════════
tab1, tab2 = st.tabs(["📊 Faturamento x Despesa Realizada", "🧾 Painel de Despesa"])

with tab1:
    c1,c2,c3,c4,c5 = st.columns([2,2,2,2,2])
    with c1: s_tipo  = st.selectbox("Despesa",         ["Todos"]+tipos_op,  key="s1_tipo")
    with c2: s_colab = st.selectbox("Colaborador",     ["Todos"]+colabs_op, key="s1_colab")
    with c3: s_mano  = st.selectbox("Mês/Ano",         ["Todos"]+meses_op,  key="s1_mano")
    with c4: s_canal = st.selectbox("Canal",           ["Todos"]+canais_op, key="s1_canal")
    with c5: s_cc    = st.selectbox("Centro de Custo", ["Todos"]+ccs_op,    key="s1_cc")
    d1,o1,f1 = filtrar(df_desp, df_orc, df_fat, s_tipo, s_colab, s_mano, s_canal, s_cc)
    renderizar(d1, o1, f1, sufixo="aba1")

with tab2:
    c1,c2,c3,c4,c5 = st.columns([2,2,2,2,2])
    with c1: p_tipo  = st.selectbox("Despesa",         ["Todos"]+tipos_op,  key="s2_tipo")
    with c2: p_colab = st.selectbox("Colaborador",     ["Todos"]+colabs_op, key="s2_colab")
    with c3: p_mano  = st.selectbox("Mês/Ano",         ["Todos"]+meses_op,  key="s2_mano")
    with c4: p_canal = st.selectbox("Canal",           ["Todos"]+canais_op, key="s2_canal")
    with c5: p_cc    = st.selectbox("Centro de Custo", ["Todos"]+ccs_op,    key="s2_cc")
    d2,o2,f2 = filtrar(df_desp, df_orc, df_fat, p_tipo, p_colab, p_mano, p_canal, p_cc)
    renderizar(d2, o2, f2, sufixo="aba2")

st.markdown(f"""
<div style="text-align:center;color:#aaa;font-size:.75rem;margin-top:24px;
            padding:10px;border-top:1px solid #dde;">
    FIRSTLAB · Despesas de Viagens · Dados até {corte_label} · Faturamento: somente "Considerar"
</div>
""", unsafe_allow_html=True)

# Botão forçar recarga
with st.sidebar:
    if st.button("🔄 Forçar recarga"):
        st.cache_data.clear()
        st.rerun()
