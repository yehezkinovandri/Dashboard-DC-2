import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
import os

# ── PAGE CONFIG ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Bank XYZ | CX Dashboard",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── GLOBAL CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
.stApp { background-color: #0d1117; }
footer { visibility: hidden; }

/* Sidebar */
[data-testid="stSidebar"] { background-color: #161b22 !important; }

/* KPI Card */
.kpi { background:#161b22; border:1px solid #30363d; border-radius:10px;
       padding:16px 18px 12px; position:relative; overflow:hidden; min-height:100px; }
.kpi::after { content:''; position:absolute; top:0; left:0; right:0; height:3px; border-radius:10px 10px 0 0; }
.kpi.blue::after   { background: linear-gradient(90deg,#58a6ff,#79c0ff); }
.kpi.green::after  { background: linear-gradient(90deg,#3fb950,#56d364); }
.kpi.amber::after  { background: linear-gradient(90deg,#d29922,#e3b341); }
.kpi.purple::after { background: linear-gradient(90deg,#8957e5,#bc8cff); }
.kpi.red::after    { background: linear-gradient(90deg,#da3633,#f85149); }
.kpi.teal::after   { background: linear-gradient(90deg,#0d7e74,#39d353); }
.kpi-lbl { font-size:10px; font-weight:700; letter-spacing:.8px;
           text-transform:uppercase; color:#8b949e; margin-bottom:4px; }
.kpi-val { font-size:30px; font-weight:800; line-height:1; margin-bottom:4px; }
.kpi-val.blue{color:#58a6ff} .kpi-val.green{color:#3fb950} .kpi-val.amber{color:#e3b341}
.kpi-val.purple{color:#bc8cff} .kpi-val.red{color:#f85149} .kpi-val.teal{color:#39d353}
.kpi-sub { font-size:10px; color:#6e7681; }

/* Section title */
.stitle { font-size:13px; font-weight:700; color:#e6edf3;
          border-left:3px solid #58a6ff; padding-left:9px; margin:12px 0 4px; }

/* Insight */
.icard { background:#161b22; border:1px solid #30363d; border-radius:10px;
         padding:13px 15px; margin-bottom:8px; }
.ihead { font-weight:700; color:#e6edf3; font-size:12px; margin-bottom:4px; }
.ibody { font-size:11px; color:#8b949e; line-height:1.6; }
</style>
""", unsafe_allow_html=True)

# ── CONSTANTS ─────────────────────────────────────────────────────────────────
DATA_PATH = os.path.join(os.path.dirname(__file__), "data", "Data_DC_-_Copy.xlsx")

PROV_COORDS = {
    "DKI Jakarta":        (-6.2088, 106.8456),
    "Jawa Barat":         (-6.9175, 107.6191),
    "Banten":             (-6.4058, 106.0640),
    "Jawa Tengah":        (-7.1500, 110.1403),
    "Jawa Timur":         (-7.5361, 112.2384),
    "Bali":               (-8.3405, 115.0920),
    "Kalimantan Selatan": (-3.0926, 115.2838),
    "Sumatera Selatan":   (-3.3194, 103.9144),
    "Kepulauan Riau":     ( 3.9457, 108.1429),
    "Lampung":            (-4.5585, 105.4068),
    "Sulawesi Selatan":   (-3.6688, 119.9741),
    "Kalimantan Timur":   ( 0.5387, 116.4194),
    "Sumatera Utara":     ( 2.1154,  99.5451),
    "Riau":               ( 0.2933, 101.7068),
}

TP_COLS = {
    "Jam Operasional":  "Dengan mempertimbangkan berbagai hal\t bagaimana penilaian keseluruhan terhadap JUMLAH & WAKTU OPERASIONAL bank - XYZ",
    "Ruang Parkir":     "Dengan mempertimbangkan berbagai hal\t bagaimana penilaian keseluruhan terhadap RUANG PARKIR bank \u2026.. untuk cabang yang ada di gedung sendiri/ruko - XYZ",
    "Banking Hall":     "Dengan mempertimbangkan berbagai hal\t bagaimana penilaian keseluruhan terhadap BANKING HALL bank\u2026.. - XYZ",
    "Toilet":           "Dengan mempertimbangkan berbagai hal\t bagaimana penilaian keseluruhan terhadap TOILET bank\u2026.. - XYZ",
    "Sekuriti":         "Dengan mempertimbangkan berbagai hal\t bagaimana penilaian keseluruhan terhadap SEKURITI Bank\u2026? - XYZ",
    "Teller":           "Dengan mempertimbangkan berbagai hal\t bagaimana penilaian keseluruhan terhadap TELLER bank\u2026? - XYZ",
    "Customer Service": "Dengan mempertimbangkan berbagai hal\t bagaimana penilaian keseluruhan terhadap CUSTOMER SERVICE bank\u2026.. - XYZ",
    "ATM":              "Dengan mempertimbangkan berbagai hal\t bagaimana penilaian keseluruhan terhadap fasilitas ATM bank\u2026. - XYZ",
}
TPK_COLS = {k: v.replace("- XYZ", "- kompetitor") for k, v in TP_COLS.items()}

EMO_POS = {
    "Bahagia":      "Saya merasa Bahagia pada saat menggunakan layanan cabang - XYZ",
    "Percaya":      "Saya Percaya dengan layanan cabang - XYZ",
    "Dihargai":     "Saya merasa Dihargai sebagai nasabah oleh cabang - XYZ",
    "Diperhatikan": "Saya merasa Diperhatikan sebagai nasabah oleh cabang - XYZ",
    "Aman":         "Saya merasa Aman pada saat menggunakan layanan cabang - XYZ",
    "Memanjakan":   "Layanan cabang Memanjakan Saya sebagai nasabah - XYZ",
    "Tertarik":     "Saya merasa Tertarik dengan layanan cabang - XYZ",
    "Semangat":     "Layanan yang diberikan cabang membuat saya merasa Penuh Semangat - XYZ",
}
EMO_NEG = {
    "Tidak Puas":   "Saya merasa Tidak Puas dengan layanan yang diberikan cabang - XYZ",
    "Frustasi":     "Saya merasa Frustasi pada saat menggunakan layanan cabang - XYZ",
    "Kecewa":       "Saya merasa Kecewa dengan layanan cabang - XYZ",
    "Tertekan":     "Saya merasa Tertekan dengan layanan yang diberikan cabang - XYZ",
    "Tidak Bahagia":"Saya merasa Tidak Bahagia dengan layanan yang diberikan cabang - XYZ",
    "Diabaikan":    "Saya merasa Diabaikan pada saat menggunakan layanan cabang - XYZ",
    "Tergesa-gesa": "Saya merasa Tergesa-gesa pada saat menggunakan layanan cabang - XYZ",
}
DIGI_COLS = {
    "Digitalisasi":   "Bank XYZ sudah melakukan proses DIGITALISASI LAYANAN CABANG DENGAN BAIK",
    "Digital Signage":"Keberadaan DIGITAL SIGNAGE di cabang XYZ MEMUDAHKAN saya sebagai nasabah dalam memperoleh informasi produk layanan XYZ",
    "Smart Table":    "Keberadaan SMART TABLE di cabang XYZ MEMUDAHKAN bagi saya sebagai nasabah dalam dalam memperoleh informasi produk layanan XYZ(Note: JAWAB 999 JIKA CABANG TIDAK MEMILIKI SMART TABLE)",
    "Tablet Survey":  "Keberadaan TABLET SURVEY di cabang XYZ MEMUDAHKAN saya sebagai nasabah untuk memberikan penilaian layanan tim cabang",
    "Akses Cabang":   "Akses ke cabang XYZ mudah",
}

# ── HELPERS ───────────────────────────────────────────────────────────────────
def to_num(s):
    return pd.to_numeric(s.astype(str).str.extract(r"^(\d+)")[0], errors="coerce")

def nps_score(series):
    s = pd.to_numeric(series, errors="coerce").dropna()
    if len(s) == 0:
        return 0.0
    return round(((s >= 9).sum() - (s <= 6).sum()) / len(s) * 100, 1)

def mean100(series, scale=6):
    s = pd.to_numeric(series, errors="coerce").dropna()
    if len(s) == 0:
        return 0.0
    return round(s.mean() / scale * 100, 1)

BASE_LAYOUT = dict(
    paper_bgcolor="#0d1117",
    plot_bgcolor="#161b22",
    font=dict(color="#8b949e", size=11),
)

def fig_layout(fig, height=260, **extra):
    layout = dict(
        **BASE_LAYOUT,
        height=height,
        margin=dict(t=8, b=8, l=0, r=0),
        xaxis=dict(gridcolor="#21262d", tickfont=dict(color="#8b949e")),
        yaxis=dict(gridcolor="#21262d", tickfont=dict(color="#8b949e")),
    )
    layout.update(extra)
    fig.update_layout(**layout)
    return fig

def kpi_card(label, value, sub, color):
    st.markdown(
        f'<div class="kpi {color}">'
        f'<div class="kpi-lbl">{label}</div>'
        f'<div class="kpi-val {color}">{value}</div>'
        f'<div class="kpi-sub">{sub}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

def insight(icon, title, body):
    st.markdown(
        f'<div class="icard">'
        f'<div class="ihead">{icon} {title}</div>'
        f'<div class="ibody">{body}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

def section(title):
    st.markdown(f'<p class="stitle">{title}</p>', unsafe_allow_html=True)

# ── LOAD DATA ─────────────────────────────────────────────────────────────────
@st.cache_data
def load_data():
    raw = pd.read_excel(DATA_PATH, header=0)
    # Drop the dummy header row
    df = raw[raw["Provinsi"] != "PROV"].copy().reset_index(drop=True)

    # Identifiers
    df["_prov"]    = df["Provinsi"].astype(str).str.strip()
    df["_area"]    = df["Kab / Kota"].astype(str).str.strip()
    df["_branch"]  = df["NAMA KANTOR CABANG"].astype(str).str.strip()
    df["_gender"]  = df["Jenis kelamin"].astype(str).str.strip()
    df["_age"]     = df["Range Usia"].astype(str).str.strip()
    df["_marital"] = df["Apa status pernikahan Bapak/Ibu saat ini?"].astype(str).str.strip()
    df["_edu"]     = df["Apakah pendidikan terakhir Bapak/Ibu?"].astype(str).str.strip()
    df["_job"]     = df["Apakah pekerjaan Bapak/Ibu saat ini?"].astype(str).str.strip()
    df["_tenure"]  = df["Sudah berapa lamakah Bapak/Ibu menjadi nasabah Bank XYZ?"].astype(str).str.strip()

    income_col = [c for c in df.columns if "PENGHASILAN RUMAH TANGGA" in str(c)][0]
    expense_col = [c for c in df.columns if "PENGELUARAN RUTIN KELUARGA" in str(c)][0]
    df["_income"]  = df[income_col].astype(str).str.strip()
    df["_expense"] = df[expense_col].astype(str).str.strip()

    # Core KPIs
    df["_nps"]      = to_num(df["NPS Bank XYZ"])
    df["_nps_k"]    = to_num(df["NPS bank kompetitor"])
    df["_csi"]      = to_num(df["Tingkat Kepuasan terhadap XYZ"])
    df["_csi_k"]    = to_num(df["Tingkat Kepuasan terhadap Bank kompetitor"])
    df["_loyalty"]  = to_num(df["Ke depannya saya pasti akan tetap menggunakan layanan dari Bank XYZ"])
    df["_ces"]      = to_num(df["Memberikan saya kemudahan untuk bertransaksi dimana pun dan kapan pun - XYZ"])
    df["_nps_seg"]  = pd.cut(df["_nps"], bins=[-1, 6, 8, 10], labels=["Detractor", "Passive", "Promoter"])

    # Touchpoints — nilai 99 adalah kode skip (cabang tidak memiliki fasilitas tsb),
    # dibuang agar tidak mendongkrak rata-rata di atas 100
    for name, col in TP_COLS.items():
        df[f"_tp_{name}"] = to_num(df[col]).replace(99, np.nan)
    for name, col in TPK_COLS.items():
        df[f"_tpk_{name}"] = to_num(df[col]).replace(99, np.nan)

    # Waiting time
    df["_tl_wait"] = pd.to_numeric(df["Berapa lama waktu tunggu/antri Teller Anda untuk kunjungan hari ini? ___ menit"], errors="coerce")
    df["_tl_tol"]  = pd.to_numeric(df["Berapa lamakah waktu tunggu/antri Teller yang Anda bisa terima/toleransi? ___ menit"], errors="coerce")
    df["_cs_wait"] = pd.to_numeric(df["Berapa lama waktu tunggu/antri CS Anda untuk kunjungan hari ini? ___ menit"], errors="coerce")
    df["_cs_tol"]  = pd.to_numeric(df["Berapa lamakah waktu tunggu/antri CS yang Anda bisa terima/toleransi? ___ menit"], errors="coerce")
    peak_tl_col = "Menurut anda\t kapankah saat PALING diperlukan untuk memperbanyak jumlah teller sehingga mengurangi jumlah antrian di bank?"
    peak_cs_col = "Menurut anda\t pada saat kapan yang PALING diperlukan untuk menambah jumlah Customer Service sehingga mengurangi jumlah antrian di bank?"
    df["_peak_tl"] = df[peak_tl_col].astype(str).str.strip()
    df["_peak_cs"] = df[peak_cs_col].astype(str).str.strip()

    # Competitor
    df["_comp"]     = df["Bank mana saja yang saat ini masih aktif Bapak/Ibu gunakan? SELAIN XYZ"].astype(str).str.strip()
    df["_save_bank"]= df["Bank manakah yang merupakan rekening utama untuk Bapak/Ibu menyimpan dana?"].astype(str).str.strip()
    df["_txn_bank"] = df["Bank manakah yang merupakan rekening utama untuk Bapak/Ibu bertransaksi (belanja\t transfer\t dsb)?"].astype(str).str.strip()

    # Emotions
    for name, col in {**EMO_POS, **EMO_NEG}.items():
        df[f"_emo_{name}"] = to_num(df[col])

    # Digitalisasi
    for name, col in DIGI_COLS.items():
        s = to_num(df[col]).replace(999, np.nan)
        df[f"_digi_{name}"] = s

    # Brand IVP
    brand_imp_cols = [c for c in df.columns
                      if "Bank yang" in str(c)
                      and "- XYZ" not in str(c)
                      and "- kompetitor" not in str(c)
                      and len(str(c)) < 100]
    for c in brand_imp_cols:
        df[f"_bimp_{c}"] = to_num(df[c])
        perf_col = c + " - XYZ"
        if perf_col in df.columns:
            df[f"_bperf_{c}"] = to_num(df[perf_col])

    return df, brand_imp_cols

df, BRAND_COLS = load_data()
ALL_TP = list(TP_COLS.keys())

# ── SIDEBAR ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🏦 Bank XYZ")
    st.markdown("**Customer Experience Dashboard**")
    st.markdown("---")
    st.markdown("#### 🗺️ Wilayah")
    all_prov = sorted(df["_prov"].dropna().unique())
    sel_prov = st.multiselect("Provinsi", all_prov, placeholder="Semua Provinsi")
    area_pool = sorted(df[df["_prov"].isin(sel_prov)]["_area"].unique()) if sel_prov \
                else sorted(df["_area"].dropna().unique())
    sel_area = st.multiselect("Kab / Kota", area_pool, placeholder="Semua Area")
    br_pool = sorted(df[df["_area"].isin(sel_area)]["_branch"].unique()) if sel_area \
              else sorted(df["_branch"].dropna().unique())
    sel_br = st.multiselect("Kantor Cabang", br_pool, placeholder="Semua Cabang")
    st.markdown("#### 🏧 Touchpoint")
    sel_tp = st.multiselect("Touchpoint", ALL_TP, placeholder="Semua Touchpoint")
    st.markdown("---")
    st.caption(f"Total Responden: **{len(df):,}**")

# ── APPLY FILTER ──────────────────────────────────────────────────────────────
fdf = df.copy()
if sel_prov: fdf = fdf[fdf["_prov"].isin(sel_prov)]
if sel_area: fdf = fdf[fdf["_area"].isin(sel_area)]
if sel_br:   fdf = fdf[fdf["_branch"].isin(sel_br)]
act_tp = sel_tp if sel_tp else ALL_TP
N = len(fdf)

# ── PAGE HEADER ───────────────────────────────────────────────────────────────
st.markdown(
    f"<h1 style='color:#e6edf3;font-size:22px;font-weight:800;margin:8px 0 2px'>🏦 Bank XYZ — Customer Experience Dashboard</h1>"
    f"<p style='color:#484f58;font-size:12px;margin:0 0 8px'>Survey nasabah cabang &nbsp;|&nbsp; "
    f"<b style='color:#8b949e'>{N:,}</b> dari {len(df):,} responden</p>",
    unsafe_allow_html=True,
)

# ── TABS ──────────────────────────────────────────────────────────────────────
tabs = st.tabs(["📊 KPI Overview", "🏧 Touchpoint", "⏱️ Waktu Tunggu", "🏦 Kompetitor", "😊 Emosi & Digital", "👥 Segmentasi"])

# ══════════════════════════════════════════════════════════════
# TAB 1 — KPI OVERVIEW
# ══════════════════════════════════════════════════════════════
with tabs[0]:
    nps  = nps_score(fdf["_nps"])
    csi  = mean100(fdf["_csi"])
    cli  = mean100(fdf["_loyalty"])
    ces  = mean100(fdf["_ces"])
    nps_k = nps_score(fdf["_nps_k"])
    csi_k = mean100(fdf["_csi_k"])
    seg   = fdf["_nps_seg"].value_counts()
    p_pct  = round(seg.get("Promoter",  0) / max(N, 1) * 100, 1)
    pa_pct = round(seg.get("Passive",   0) / max(N, 1) * 100, 1)
    d_pct  = round(seg.get("Detractor", 0) / max(N, 1) * 100, 1)

    c1, c2, c3, c4 = st.columns(4)
    with c1: kpi_card("Net Promoter Score", nps,
                       f"🟢 {p_pct}% Promoter &nbsp; 🟡 {pa_pct}% Passive &nbsp; 🔴 {d_pct}% Detractor", "blue")
    with c2: kpi_card("Customer Satisfaction", csi,
                       f"Avg {fdf['_csi'].mean():.2f}/6 &nbsp;|&nbsp; Kompetitor: {csi_k}", "green")
    with c3: kpi_card("Customer Loyalty Index", cli,
                       f"Avg {fdf['_loyalty'].mean():.2f}/6 &nbsp;|&nbsp; Skala 0–100", "amber")
    with c4: kpi_card("Customer Effort Score", ces,
                       f"Avg {fdf['_ces'].mean():.2f}/6 &nbsp;|&nbsp; Skala 0–100", "purple")

    st.markdown("<br>", unsafe_allow_html=True)

    col_l, col_r = st.columns([1, 1.2])
    with col_l:
        section("NPS Segmentasi")
        fig = go.Figure(go.Pie(
            labels=["Promoter", "Passive", "Detractor"],
            values=[p_pct, pa_pct, d_pct], hole=0.62,
            marker_colors=["#3fb950", "#e3b341", "#f85149"],
            textfont=dict(color="#e6edf3", size=12),
        ))
        fig.update_layout(paper_bgcolor="#0d1117", plot_bgcolor="#161b22", font=dict(color="#8b949e", size=12), height=360,
            legend=dict(font=dict(color="#c9d1d9", size=13), bgcolor="#0d1117"),
            annotations=[dict(text=f"<b>{nps}</b><br>NPS", x=0.5, y=0.5,
                               showarrow=False, font=dict(size=26, color="#58a6ff"))])
        st.plotly_chart(fig, width='stretch', key="chart_1", config={"displayModeBar": False})

    with col_r:
        section("Distribusi Skor NPS (0–10)")
        nd = fdf["_nps"].value_counts().sort_index()
        fig2 = go.Figure(go.Bar(
            x=nd.index, y=nd.values,
            marker_color=["#f85149" if i <= 6 else "#e3b341" if i <= 8 else "#3fb950" for i in nd.index],
            text=nd.values, textposition="outside", textfont=dict(color="#c9d1d9"),
        ))
        for x0, x1, clr, lbl in [(-0.5,6.5,"#f85149","Detractor"),(6.5,8.5,"#e3b341","Passive"),(8.5,10.5,"#3fb950","Promoter")]:
            fig2.add_vrect(x0=x0, x1=x1, fillcolor=clr, opacity=0.07,
                           annotation_text=lbl, annotation_font_color=clr, annotation_position="top left")
        fig_layout(fig2, 200)
        st.plotly_chart(fig2, width='stretch', key="chart_2", config={"displayModeBar": False})

    # Map
    section("🗺️ Peta Sebaran Responden & NPS per Provinsi")
    prov_df = fdf[fdf["_prov"].isin(PROV_COORDS)].groupby("_prov").agg(
        n=("_nps", "count"),
        NPS=("_nps", nps_score),
        CSI=("_csi", mean100),
    ).reset_index()
    prov_df["lat"] = prov_df["_prov"].map(lambda x: PROV_COORDS[x][0])
    prov_df["lon"] = prov_df["_prov"].map(lambda x: PROV_COORDS[x][1])
    fig_map = px.scatter_map(
        prov_df, lat="lat", lon="lon", size="n", color="NPS",
        hover_name="_prov", hover_data={"n": True, "NPS": True, "CSI": True, "lat": False, "lon": False},
        color_continuous_scale=[[0,"#f85149"],[0.5,"#e3b341"],[1,"#3fb950"]],
        range_color=[0, 100], size_max=50, zoom=4,
        map_style="carto-darkmatter",
        labels={"n": "Responden", "NPS": "NPS Score"},
    )
    fig_map.update_layout(**{k:v for k,v in BASE_LAYOUT.items() if k != 'margin'},
                           height=360, margin=dict(t=0,b=0,l=0,r=0),
                           coloraxis_colorbar=dict(tickfont=dict(color="#8b949e"), thickness=10))
    st.plotly_chart(fig_map, width='stretch', key="chart_3", config={"displayModeBar": False})

    v1, v2 = st.columns(2)
    with v1:
        section("XYZ vs Kompetitor (NPS & CSI)")
        fig_c = go.Figure()
        fig_c.add_trace(go.Bar(name="XYZ", x=["NPS","CSI"], y=[nps, csi], marker_color="#58a6ff",
                                text=[nps,csi], textposition="outside", textfont=dict(color="#e6edf3")))
        fig_c.add_trace(go.Bar(name="Kompetitor", x=["NPS","CSI"], y=[nps_k, csi_k], marker_color="#484f58",
                                text=[nps_k,csi_k], textposition="outside", textfont=dict(color="#8b949e")))
        fig_layout(fig_c, 230, barmode="group", legend=dict(font=dict(color="#c9d1d9"), bgcolor="#0d1117"))
        st.plotly_chart(fig_c, width='stretch', key="chart_4", config={"displayModeBar": False})

    with v2:
        section("Top 12 Cabang by NPS (min 5 responden)")
        br = fdf.groupby("_branch").agg(NPS=("_nps", nps_score), n=("_nps","count")).reset_index()
        br = br[br["n"] >= 5].nlargest(12, "NPS")
        fig_br = go.Figure(go.Bar(
            x=br["NPS"], y=br["_branch"], orientation="h",
            marker_color=["#3fb950" if v>=70 else "#e3b341" if v>=50 else "#f85149" for v in br["NPS"]],
            text=br["NPS"].round(1), textposition="outside", textfont=dict(color="#e6edf3"),
        ))
        fig_layout(fig_br, 320)
        fig_br.update_layout(yaxis=dict(autorange="reversed", tickfont=dict(color="#c9d1d9", size=9)))
        st.plotly_chart(fig_br, width='stretch', key="chart_5", config={"displayModeBar": False})

    section("💡 Insight & Rekomendasi")
    i1, i2 = st.columns(2)
    gap = round(nps - nps_k, 1)
    with i1:
        if nps >= 70:
            insight("🚀","NPS Excellent", f"NPS <b>{nps}</b> sangat kuat. {p_pct}% Promoter bisa diaktifkan sebagai brand ambassador melalui program referral.")
        elif nps >= 50:
            insight("✅","NPS Baik", f"NPS <b>{nps}</b> di level baik. Konversi <b>{pa_pct}%</b> Passive → Promoter adalah peluang terbesar.")
        else:
            insight("⚠️","NPS Perlu Perhatian", f"NPS <b>{nps}</b> perlu tindakan. <b>{d_pct}%</b> Detractor butuh program service recovery segera.")
        insight("📊","CSI & CLI", f"CSI <b>{csi}</b> dan CLI <b>{cli}</b>. CES <b>{ces}</b> mencerminkan kemudahan transaksi — tingkatkan untuk dampak langsung ke NPS.")
    with i2:
        if gap > 0:
            insight("📈",f"Unggul +{gap} poin vs Kompetitor", "Pertahankan keunggulan ini sebagai competitive advantage dan perkuat di komunikasi pemasaran.")
        else:
            insight("📉",f"Tertinggal {abs(gap)} poin vs Kompetitor", "Lakukan benchmarking mendalam terhadap kompetitor dan susun action plan 60 hari.")
        insight("🎯","Quick Win", f"Setiap 10 poin peningkatan CES berpotensi menaikkan NPS ~5–8 poin. Prioritaskan kemudahan transaksi digital dan pengurangan waktu tunggu.")

# ══════════════════════════════════════════════════════════════
# TAB 2 — TOUCHPOINT
# ══════════════════════════════════════════════════════════════
with tabs[1]:
    tp_xyz = {t: mean100(fdf[f"_tp_{t}"]) for t in act_tp}
    tp_kmp = {t: mean100(fdf[f"_tpk_{t}"]) for t in act_tp}

    section("Skor per Touchpoint: XYZ vs Kompetitor")
    fig_tp = go.Figure()
    fig_tp.add_trace(go.Bar(name="XYZ", y=list(tp_xyz.keys()), x=list(tp_xyz.values()),
                             orientation="h", marker_color="#58a6ff",
                             text=[f"{v:.0f}" for v in tp_xyz.values()],
                             textposition="outside", textfont=dict(color="#e6edf3")))
    fig_tp.add_trace(go.Bar(name="Kompetitor", y=list(tp_kmp.keys()), x=list(tp_kmp.values()),
                             orientation="h", marker_color="#484f58",
                             text=[f"{v:.0f}" for v in tp_kmp.values()],
                             textposition="outside", textfont=dict(color="#8b949e")))
    fig_layout(fig_tp, 300, barmode="group",
               legend=dict(font=dict(color="#c9d1d9"), bgcolor="#0d1117"))
    fig_tp.update_layout(xaxis=dict(range=[0, 115]))
    fig_tp.update_yaxes(autorange="reversed")
    st.plotly_chart(fig_tp, width='stretch', key="chart_6", config={"displayModeBar": False})

    r1, r2 = st.columns(2)
    with r1:
        section("Radar: XYZ vs Kompetitor")
        cats_r = act_tp + [act_tp[0]]
        fig_r = go.Figure()
        fig_r.add_trace(go.Scatterpolar(
            r=[tp_xyz[t] for t in act_tp] + [tp_xyz[act_tp[0]]],
            theta=cats_r, fill="toself", name="XYZ",
            fillcolor="rgba(88,166,255,0.15)", line=dict(color="#58a6ff", width=2),
        ))
        fig_r.add_trace(go.Scatterpolar(
            r=[tp_kmp[t] for t in act_tp] + [tp_kmp[act_tp[0]]],
            theta=cats_r, fill="toself", name="Kompetitor",
            fillcolor="rgba(72,79,88,0.15)", line=dict(color="#484f58", width=2),
        ))
        fig_r.update_layout(paper_bgcolor="#0d1117", plot_bgcolor="#161b22", font=dict(color="#8b949e", size=11), height=300,
            polar=dict(
                radialaxis=dict(visible=True, range=[0,100], gridcolor="#21262d", tickfont=dict(color="#484f58", size=9)),
                angularaxis=dict(tickfont=dict(color="#c9d1d9", size=10), gridcolor="#21262d"),
                bgcolor="#161b22",
            ),
            legend=dict(font=dict(color="#c9d1d9"), bgcolor="#0d1117"),
        )
        st.plotly_chart(fig_r, width='stretch', key="chart_7", config={"displayModeBar": False})

    with r2:
        section("Heatmap Touchpoint per Area (Top 10)")
        top10 = fdf["_area"].value_counts().head(10).index
        ht = fdf[fdf["_area"].isin(top10)].groupby("_area")[
            [f"_tp_{t}" for t in act_tp]].mean().multiply(100/6).round(1)
        ht.columns = act_tp
        fig_ht = go.Figure(go.Heatmap(
            z=ht.values, x=ht.columns.tolist(), y=ht.index.tolist(),
            colorscale=[[0,"#f85149"],[0.5,"#e3b341"],[1,"#3fb950"]],
            zmin=60, zmax=100,
            text=ht.values.round(0), texttemplate="%{text}",
            textfont=dict(color="white", size=10),
            colorbar=dict(tickfont=dict(color="#8b949e"), thickness=10),
        ))
        fig_ht.update_layout(paper_bgcolor="#0d1117", plot_bgcolor="#161b22", font=dict(color="#8b949e", size=11), height=420,
            xaxis=dict(tickfont=dict(color="#c9d1d9", size=10)),
            yaxis=dict(tickfont=dict(color="#c9d1d9", size=10)),
        )
        st.plotly_chart(fig_ht, width='stretch', key="chart_8", config={"displayModeBar": False})

    # Importance vs Performance Matrix
    section("🎯 Importance vs Performance Matrix")
    ivp_rows = []
    for c in BRAND_COLS:
        imp_v  = fdf[f"_bimp_{c}"].mean()
        perf_v = fdf[f"_bperf_{c}"].mean() if f"_bperf_{c}" in fdf.columns else np.nan
        if np.isnan(imp_v) or np.isnan(perf_v):
            continue
        lbl = c.replace("Bank yang ", "").strip()
        lbl = (lbl[:35] + "…") if len(lbl) > 35 else lbl
        ivp_rows.append({"label": lbl, "imp": imp_v, "perf": perf_v})

    if ivp_rows:
        ivp = pd.DataFrame(ivp_rows)
        avg_i = ivp["imp"].mean()
        avg_p = ivp["perf"].mean()

        def get_quad(row):
            if row["imp"] >= avg_i and row["perf"] <  avg_p: return "Prioritas Utama",  "#f85149"
            if row["imp"] >= avg_i and row["perf"] >= avg_p: return "Pertahankan",       "#3fb950"
            if row["imp"] <  avg_i and row["perf"] <  avg_p: return "Prioritas Rendah",  "#e3b341"
            return                                                    "Berlebihan",        "#58a6ff"

        ivp[["quad", "clr"]] = ivp.apply(get_quad, axis=1, result_type="expand")
        fig_ivp = go.Figure()
        for q, c in [("Prioritas Utama","#f85149"),("Pertahankan","#3fb950"),
                      ("Prioritas Rendah","#e3b341"),("Berlebihan","#58a6ff")]:
            sub = ivp[ivp["quad"] == q]
            if len(sub) == 0:
                continue
            fig_ivp.add_trace(go.Scatter(
                x=sub["imp"], y=sub["perf"], mode="markers+text", name=q,
                marker=dict(color=c, size=11, opacity=0.85, line=dict(color="#0d1117", width=1)),
                text=sub["label"], textposition="top center", textfont=dict(color=c, size=8),
                hovertemplate="<b>%{text}</b><br>Importance: %{x:.2f}<br>Performance: %{y:.2f}<extra></extra>",
            ))
        fig_ivp.add_vline(x=avg_i, line=dict(color="#30363d", dash="dash", width=1))
        fig_ivp.add_hline(y=avg_p, line=dict(color="#30363d", dash="dash", width=1))
        fig_layout(fig_ivp, 380,
            xaxis=dict(title="Importance", gridcolor="#21262d", tickfont=dict(color="#8b949e"), title_font=dict(color="#8b949e")),
            yaxis=dict(title="Performance", gridcolor="#21262d", tickfont=dict(color="#8b949e"), title_font=dict(color="#8b949e")),
            legend=dict(font=dict(color="#c9d1d9"), bgcolor="#161b22", bordercolor="#30363d", borderwidth=1),
        )
        st.plotly_chart(fig_ivp, width='stretch', key="chart_9", config={"displayModeBar": False})

        priority = ivp[ivp["quad"] == "Prioritas Utama"].nlargest(3, "imp")["label"].tolist()
        best_tp  = max(tp_xyz, key=tp_xyz.get)
        worst_tp = min(tp_xyz, key=tp_xyz.get)
        ti1, ti2 = st.columns(2)
        with ti1:
            insight("⭐", f"Touchpoint Terbaik: {best_tp}", f"Skor <b>{tp_xyz[best_tp]:.1f}/100</b>. Jadikan benchmark internal dan replikasikan best practice ke touchpoint lain.")
        with ti2:
            insight("🎯", "Prioritas I-P Matrix", f"Atribut penting tapi performa rendah: <b>{', '.join(priority) if priority else 'Semua sudah baik'}</b>. Alokasikan resources ke area ini untuk dampak maksimal.")

# ══════════════════════════════════════════════════════════════
# TAB 3 — WAKTU TUNGGU
# ══════════════════════════════════════════════════════════════
with tabs[2]:
    tl_w = fdf["_tl_wait"].dropna()
    tl_t = fdf["_tl_tol"].dropna()
    cs_w = fdf["_cs_wait"].dropna()
    cs_t = fdf["_cs_tol"].dropna()
    tl_gap = round(tl_t.mean() - tl_w.mean(), 1)
    cs_gap = round(cs_t.mean() - cs_w.mean(), 1)

    k1,k2,k3,k4,k5,k6 = st.columns(6)
    with k1: kpi_card("Median Tunggu Teller",  f"{tl_w.median():.0f} min", f"Avg {tl_w.mean():.1f} min", "blue")
    with k2: kpi_card("Toleransi Teller",       f"{tl_t.median():.0f} min", "Median toleransi nasabah",    "green")
    with k3: kpi_card("Gap Teller",              f"{tl_gap:+.1f} min",       "Toleransi − Aktual",          "teal" if tl_gap >= 0 else "red")
    with k4: kpi_card("Median Tunggu CS",        f"{cs_w.median():.0f} min", f"Avg {cs_w.mean():.1f} min", "amber")
    with k5: kpi_card("Toleransi CS",            f"{cs_t.median():.0f} min", "Median toleransi nasabah",    "green")
    with k6: kpi_card("Gap CS",                   f"{cs_gap:+.1f} min",       "Toleransi − Aktual",          "teal" if cs_gap >= 0 else "red")

    st.markdown("<br>", unsafe_allow_html=True)
    w1, w2 = st.columns(2)
    with w1:
        section("Distribusi Waktu Tunggu Teller (cap 60 min)")
        fig_tw = go.Figure()
        fig_tw.add_trace(go.Histogram(x=tl_w.clip(0,60), name="Aktual",    nbinsx=20, marker_color="#58a6ff", opacity=0.7))
        fig_tw.add_trace(go.Histogram(x=tl_t.clip(0,60), name="Toleransi", nbinsx=20, marker_color="#3fb950", opacity=0.5))
        fig_tw.add_vline(x=tl_w.median(), line=dict(color="#58a6ff", dash="dash", width=1.5),
                          annotation_text=f"Median {tl_w.median():.0f}m", annotation_font_color="#58a6ff")
        fig_layout(fig_tw, 240, barmode="overlay", legend=dict(font=dict(color="#c9d1d9"), bgcolor="#0d1117"))
        st.plotly_chart(fig_tw, width='stretch', key="chart_10", config={"displayModeBar": False})
    with w2:
        section("Distribusi Waktu Tunggu CS (cap 60 min)")
        fig_cw = go.Figure()
        fig_cw.add_trace(go.Histogram(x=cs_w.clip(0,60), name="Aktual",    nbinsx=20, marker_color="#e3b341", opacity=0.7))
        fig_cw.add_trace(go.Histogram(x=cs_t.clip(0,60), name="Toleransi", nbinsx=20, marker_color="#3fb950", opacity=0.5))
        fig_cw.add_vline(x=cs_w.median(), line=dict(color="#e3b341", dash="dash", width=1.5),
                          annotation_text=f"Median {cs_w.median():.0f}m", annotation_font_color="#e3b341")
        fig_layout(fig_cw, 240, barmode="overlay", legend=dict(font=dict(color="#c9d1d9"), bgcolor="#0d1117"))
        st.plotly_chart(fig_cw, width='stretch', key="chart_11", config={"displayModeBar": False})

    w3, w4 = st.columns(2)
    with w3:
        section("Jam Tersibuk — Teller")
        pk_tl = fdf["_peak_tl"].replace(" ", np.nan).dropna().value_counts()
        pk_tl = pk_tl[pk_tl.index.str.len() > 3]
        fig_pt = go.Figure(go.Bar(y=pk_tl.index, x=pk_tl.values, orientation="h",
                                   marker_color="#58a6ff", text=pk_tl.values,
                                   textposition="outside", textfont=dict(color="#e6edf3")))
        fig_layout(fig_pt, 230)
        fig_pt.update_yaxes(autorange="reversed", tickfont=dict(color="#c9d1d9", size=9))
        st.plotly_chart(fig_pt, width='stretch', key="chart_12", config={"displayModeBar": False})
    with w4:
        section("Jam Tersibuk — CS")
        pk_cs = fdf["_peak_cs"].replace(" ", np.nan).dropna().value_counts()
        pk_cs = pk_cs[pk_cs.index.str.len() > 3]
        fig_pc = go.Figure(go.Bar(y=pk_cs.index, x=pk_cs.values, orientation="h",
                                   marker_color="#e3b341", text=pk_cs.values,
                                   textposition="outside", textfont=dict(color="#e6edf3")))
        fig_layout(fig_pc, 230)
        fig_pc.update_yaxes(autorange="reversed", tickfont=dict(color="#c9d1d9", size=9))
        st.plotly_chart(fig_pc, width='stretch', key="chart_13", config={"displayModeBar": False})

    section("Rata-rata Waktu Tunggu per Area (Top 10)")
    top10a = fdf["_area"].value_counts().head(10).index
    wa = fdf[fdf["_area"].isin(top10a)].groupby("_area").agg(
        Teller=("_tl_wait","median"), CS=("_cs_wait","median")).reset_index().sort_values("Teller")
    fig_wa = go.Figure()
    fig_wa.add_trace(go.Bar(name="Teller", y=wa["_area"], x=wa["Teller"], orientation="h",
                             marker_color="#58a6ff", text=wa["Teller"].round(1),
                             textposition="outside", textfont=dict(color="#e6edf3")))
    fig_wa.add_trace(go.Bar(name="CS", y=wa["_area"], x=wa["CS"], orientation="h",
                             marker_color="#e3b341", text=wa["CS"].round(1),
                             textposition="outside", textfont=dict(color="#e6edf3")))
    fig_layout(fig_wa, 300, barmode="group", legend=dict(font=dict(color="#c9d1d9"), bgcolor="#0d1117"))
    fig_wa.update_yaxes(tickfont=dict(color="#c9d1d9", size=10))
    st.plotly_chart(fig_wa, width='stretch', key="chart_14", config={"displayModeBar": False})

    section("💡 Insight Waktu Tunggu")
    wi1, wi2 = st.columns(2)
    with wi1:
        insight("⏱️","Teller", f"Median tunggu <b>{tl_w.median():.0f} menit</b>, toleransi <b>{tl_t.median():.0f} menit</b>. Gap <b>{tl_gap:+.1f} menit</b> — {'dalam batas toleransi ✅' if tl_gap >= 0 else 'MELEBIHI toleransi ⚠️'}.")
        insight("📋","Rekomendasi Teller", f"Peak hour: <b>{pk_tl.index[0] if len(pk_tl) else '-'}</b>. Tambah Teller di jam tersebut. Target: rata-rata tunggu di bawah 5 menit.")
    with wi2:
        insight("⏱️","CS", f"Median tunggu <b>{cs_w.median():.0f} menit</b>, toleransi <b>{cs_t.median():.0f} menit</b>. Gap <b>{cs_gap:+.1f} menit</b> — {'dalam batas toleransi ✅' if cs_gap >= 0 else 'MELEBIHI toleransi ⚠️'}.")
        insight("📋","Rekomendasi CS", f"Peak hour: <b>{pk_cs.index[0] if len(pk_cs) else '-'}</b>. Optimalkan CS digital untuk transaksi sederhana dan tambah staf di jam pagi.")

# ══════════════════════════════════════════════════════════════
# TAB 4 — KOMPETITOR
# ══════════════════════════════════════════════════════════════
with tabs[3]:
    def parse_banks(series):
        rows = []
        for val in series.dropna():
            v = str(val).strip()
            if v and v not in ("nan", " ", ""):
                for b in v.split(";"):
                    b = b.strip()
                    if b and len(b) > 2:
                        rows.append(b)
        return pd.Series(rows).value_counts()

    comp_vc = parse_banks(fdf["_comp"])
    save_vc = fdf["_save_bank"].replace(" ", np.nan).dropna().value_counts()
    txn_vc  = fdf["_txn_bank"].replace(" ", np.nan).dropna().value_counts()

    b1, b2, b3 = st.columns(3)
    with b1:
        section("Bank Lain yang Aktif Digunakan")
        top_c = comp_vc.head(8)
        fig_bc = go.Figure(go.Bar(y=top_c.index, x=top_c.values, orientation="h",
                                   marker_color="#484f58", text=top_c.values,
                                   textposition="outside", textfont=dict(color="#e6edf3")))
        fig_layout(fig_bc, 280)
        fig_bc.update_yaxes(autorange="reversed", tickfont=dict(color="#c9d1d9", size=9))
        st.plotly_chart(fig_bc, width='stretch', key="chart_15", config={"displayModeBar": False})
    with b2:
        section("Rekening Utama Simpanan")
        ts = save_vc.head(6)
        fig_sv = go.Figure(go.Pie(labels=ts.index, values=ts.values, hole=0.5,
                                   marker_colors=["#3fb950","#58a6ff","#e3b341","#f85149","#bc8cff","#39d353"],
                                   textfont=dict(color="#e6edf3")))
        fig_sv.update_layout(paper_bgcolor="#0d1117", plot_bgcolor="#161b22", font=dict(color="#8b949e", size=11), height=280, legend=dict(font=dict(color="#c9d1d9", size=9), bgcolor="#0d1117"))
        st.plotly_chart(fig_sv, width='stretch', key="chart_16", config={"displayModeBar": False})
    with b3:
        section("Rekening Utama Transaksi")
        tt = txn_vc.head(6)
        fig_tx = go.Figure(go.Pie(labels=tt.index, values=tt.values, hole=0.5,
                                   marker_colors=["#3fb950","#58a6ff","#e3b341","#f85149","#bc8cff","#39d353"],
                                   textfont=dict(color="#e6edf3")))
        fig_tx.update_layout(paper_bgcolor="#0d1117", plot_bgcolor="#161b22", font=dict(color="#8b949e", size=11), height=280, legend=dict(font=dict(color="#c9d1d9", size=9), bgcolor="#0d1117"))
        st.plotly_chart(fig_tx, width='stretch', key="chart_17", config={"displayModeBar": False})

    nps2  = nps_score(fdf["_nps"]); nps_k2 = nps_score(fdf["_nps_k"])
    csi2  = mean100(fdf["_csi"]);   csi_k2  = mean100(fdf["_csi_k"])
    pct_xyz = round((fdf["_save_bank"].str.contains("XYZ", na=False)).mean() * 100, 1)
    ck1, ck2, ck3 = st.columns(3)
    with ck1: kpi_card("NPS Kompetitor",  nps_k2, f"XYZ: {nps2} (gap: {nps2-nps_k2:+.1f})", "blue" if nps2 >= nps_k2 else "red")
    with ck2: kpi_card("CSI Kompetitor",  csi_k2, f"XYZ: {csi2} (gap: {csi2-csi_k2:+.1f})", "green" if csi2 >= csi_k2 else "amber")
    with ck3: kpi_card("Simpanan Utama di XYZ", f"{pct_xyz}%", "Nasabah yang menjadikan XYZ bank simpanan utama", "teal")

    st.markdown("<br>", unsafe_allow_html=True)
    section("Gap Touchpoint XYZ vs Kompetitor")
    tp_xyz2 = {t: mean100(fdf[f"_tp_{t}"]) for t in act_tp}
    tp_kmp2 = {t: mean100(fdf[f"_tpk_{t}"]) for t in act_tp}
    gap_df = pd.DataFrame({
        "Touchpoint": act_tp,
        "Gap": [tp_xyz2[t] - tp_kmp2[t] for t in act_tp],
    }).sort_values("Gap")
    fig_gap = go.Figure(go.Bar(
        x=gap_df["Gap"], y=gap_df["Touchpoint"], orientation="h",
        marker_color=["#3fb950" if v >= 0 else "#f85149" for v in gap_df["Gap"]],
        text=[f"{v:+.1f}" for v in gap_df["Gap"]],
        textposition="outside", textfont=dict(color="#e6edf3"),
    ))
    fig_gap.add_vline(x=0, line=dict(color="#484f58", width=1))
    fig_layout(fig_gap, 280)
    fig_gap.update_yaxes(tickfont=dict(color="#c9d1d9", size=10))
    st.plotly_chart(fig_gap, width='stretch', key="chart_18", config={"displayModeBar": False})
    st.caption("Gap positif = XYZ unggul · Gap negatif = Kompetitor unggul")

    section("💡 Insight Kompetitor")
    top_rival = comp_vc.index[0] if len(comp_vc) > 0 else "BCA"
    unggul = gap_df[gap_df["Gap"] > 0]["Touchpoint"].tolist()
    lemah  = gap_df[gap_df["Gap"] < 0]["Touchpoint"].tolist()
    ci1, ci2 = st.columns(2)
    with ci1:
        insight("🏦","Kompetitor Utama", f"Bank yang paling banyak digunakan selain XYZ: <b>{top_rival}</b>. XYZ menjadi simpanan utama bagi <b>{pct_xyz}%</b> nasabah.")
        insight("📈","Keunggulan XYZ", f"XYZ unggul di: <b>{', '.join(unggul) if unggul else 'semua touchpoint'}</b>. Perkuat sebagai selling point dalam komunikasi pemasaran.")
    with ci2:
        insight("⚠️","Area Tertinggal", f"XYZ tertinggal di: <b>{', '.join(lemah) if lemah else 'tidak ada — XYZ unggul semua!'}</b>. Lakukan benchmarking proses dan fasilitas kompetitor.")
        insight("🎯","Rekomendasi", f"Fokus cegah migrasi ke <b>{top_rival}</b> dengan program loyalitas, kemudahan digital, dan layanan personal yang menjadi pembeda XYZ.")

# ══════════════════════════════════════════════════════════════
# TAB 5 — EMOSI & DIGITAL
# ══════════════════════════════════════════════════════════════
with tabs[4]:
    pos_s = {e: mean100(fdf[f"_emo_{e}"]) for e in EMO_POS}
    neg_s = {e: mean100(fdf[f"_emo_{e}"]) for e in EMO_NEG}
    avg_pos = round(np.mean(list(pos_s.values())), 1)
    avg_neg = round(np.mean(list(neg_s.values())), 1)
    gap_e   = round(avg_pos - avg_neg, 1)

    em1, em2, em3 = st.columns(3)
    with em1: kpi_card("Rata-rata Emosi Positif", avg_pos, "8 dimensi · Skala 0–100", "green")
    with em2: kpi_card("Rata-rata Emosi Negatif", avg_neg, "7 dimensi · Idealnya rendah", "red")
    with em3: kpi_card("Sentiment Gap", f"+{gap_e}", "Positif − Negatif · makin besar makin baik", "teal")

    st.markdown("<br>", unsafe_allow_html=True)
    e1, e2 = st.columns(2)
    with e1:
        section("😊 Emosi Positif (0–100)")
        fig_pe = go.Figure(go.Bar(
            x=list(pos_s.values()), y=list(pos_s.keys()), orientation="h",
            marker_color="#3fb950", text=[f"{v:.0f}" for v in pos_s.values()],
            textposition="outside", textfont=dict(color="#e6edf3"),
        ))
        fig_layout(fig_pe, 260)
        fig_pe.update_layout(xaxis=dict(range=[0, 115]))
        fig_pe.update_yaxes(autorange="reversed", tickfont=dict(color="#c9d1d9"))
        st.plotly_chart(fig_pe, width='stretch', key="chart_19", config={"displayModeBar": False})
    with e2:
        section("😠 Emosi Negatif (lebih rendah = lebih baik)")
        fig_ne = go.Figure(go.Bar(
            x=list(neg_s.values()), y=list(neg_s.keys()), orientation="h",
            marker_color="#f85149", text=[f"{v:.0f}" for v in neg_s.values()],
            textposition="outside", textfont=dict(color="#e6edf3"),
        ))
        fig_layout(fig_ne, 240)
        fig_ne.update_layout(xaxis=dict(range=[0, 115]))
        fig_ne.update_yaxes(autorange="reversed", tickfont=dict(color="#c9d1d9"))
        st.plotly_chart(fig_ne, width='stretch', key="chart_20", config={"displayModeBar": False})

    section("Heatmap Emosi per Area (Top 8)")
    top8 = fdf["_area"].value_counts().head(8).index
    all_e = list(EMO_POS.keys()) + list(EMO_NEG.keys())
    eh = fdf[fdf["_area"].isin(top8)].groupby("_area")[
        [f"_emo_{e}" for e in all_e]].mean().multiply(100/6).round(1)
    eh.columns = all_e
    fig_eh = go.Figure(go.Heatmap(
        z=eh.values, x=eh.columns.tolist(), y=eh.index.tolist(),
        colorscale=[[0,"#f85149"],[0.5,"#e3b341"],[1,"#3fb950"]],
        zmin=0, zmax=100,
        text=eh.values.round(0), texttemplate="%{text}",
        textfont=dict(color="white", size=9),
        colorbar=dict(tickfont=dict(color="#8b949e"), thickness=10),
    ))
    fig_eh.update_layout(paper_bgcolor="#0d1117", plot_bgcolor="#161b22", font=dict(color="#8b949e", size=11), height=380,
        xaxis=dict(tickfont=dict(color="#c9d1d9", size=10), tickangle=-30),
        yaxis=dict(tickfont=dict(color="#c9d1d9", size=11)),
    )
    st.plotly_chart(fig_eh, width='stretch', key="chart_21", config={"displayModeBar": False})

    st.markdown("---")
    section("📱 Digitalisasi Layanan Cabang")
    digi_s = {d: mean100(fdf[f"_digi_{d}"]) for d in DIGI_COLS}
    dc = st.columns(5)
    for i, (lbl, sc) in enumerate(digi_s.items()):
        with dc[i]: kpi_card(lbl, f"{sc:.0f}", "Skor 0–100", ["blue","green","amber","purple","teal"][i])

    st.markdown("<br>", unsafe_allow_html=True)
    dg1, dg2 = st.columns([1.2, 1])
    with dg1:
        section("Heatmap Digitalisasi per Area (Top 10)")
        top10d = fdf["_area"].value_counts().head(10).index
        da = fdf[fdf["_area"].isin(top10d)].groupby("_area")[
            [f"_digi_{d}" for d in DIGI_COLS]].mean().multiply(100/6).round(1)
        da.columns = list(DIGI_COLS.keys())
        fig_dh = go.Figure(go.Heatmap(
            z=da.values, x=da.columns.tolist(), y=da.index.tolist(),
            colorscale=[[0,"#f85149"],[0.5,"#e3b341"],[1,"#3fb950"]],
            zmin=50, zmax=100,
            text=da.values.round(0), texttemplate="%{text}",
            textfont=dict(color="white", size=10),
            colorbar=dict(tickfont=dict(color="#8b949e"), thickness=10),
        ))
        fig_dh.update_layout(paper_bgcolor="#0d1117", plot_bgcolor="#161b22", font=dict(color="#8b949e", size=11), height=400,
            xaxis=dict(tickfont=dict(color="#c9d1d9", size=11)),
            yaxis=dict(tickfont=dict(color="#c9d1d9", size=11)),
        )
        st.plotly_chart(fig_dh, width='stretch', key="chart_22", config={"displayModeBar": False})
    with dg2:
        section("Radar Digitalisasi")
        dv = list(digi_s.values()); dl = list(digi_s.keys())
        fig_dr = go.Figure(go.Scatterpolar(
            r=dv + [dv[0]], theta=dl + [dl[0]], fill="toself",
            fillcolor="rgba(188,140,255,0.15)",
            line=dict(color="#bc8cff", width=2), marker=dict(color="#bc8cff", size=7),
        ))
        fig_dr.update_layout(paper_bgcolor="#0d1117", plot_bgcolor="#161b22", font=dict(color="#8b949e", size=11), height=400,
            polar=dict(
                radialaxis=dict(visible=True, range=[0,100], gridcolor="#21262d", tickfont=dict(color="#484f58", size=9)),
                angularaxis=dict(tickfont=dict(color="#c9d1d9", size=10), gridcolor="#21262d"),
                bgcolor="#161b22",
            ),
        )
        st.plotly_chart(fig_dr, width='stretch', key="chart_23", config={"displayModeBar": False})

# ══════════════════════════════════════════════════════════════
# TAB 6 — SEGMENTASI
# ══════════════════════════════════════════════════════════════
with tabs[5]:
    section("👥 Profil Demografis")
    s1, s2 = st.columns(2)
    with s1:
        g = fdf["_gender"].value_counts()
        fig_g = go.Figure(go.Pie(labels=g.index, values=g.values, hole=0.55,
                                  marker_colors=["#58a6ff","#bc8cff"], textfont=dict(color="#e6edf3", size=13)))
        fig_g.update_layout(paper_bgcolor="#0d1117", plot_bgcolor="#161b22", font=dict(color="#8b949e", size=12),
                             height=280, title=dict(text="Jenis Kelamin", font=dict(color="#8b949e", size=13)),
                             legend=dict(font=dict(color="#c9d1d9", size=12), bgcolor="#0d1117"))
        st.plotly_chart(fig_g, width='stretch', key="chart_24", config={"displayModeBar": False})
    with s2:
        m = fdf["_marital"].value_counts()
        fig_m = go.Figure(go.Pie(labels=m.index, values=m.values, hole=0.55,
                                  marker_colors=["#3fb950","#e3b341","#f85149"], textfont=dict(color="#e6edf3", size=13)))
        fig_m.update_layout(paper_bgcolor="#0d1117", plot_bgcolor="#161b22", font=dict(color="#8b949e", size=12),
                             height=280, title=dict(text="Status Pernikahan", font=dict(color="#8b949e", size=13)),
                             legend=dict(font=dict(color="#c9d1d9", size=12), bgcolor="#0d1117"))
        st.plotly_chart(fig_m, width='stretch', key="chart_25", config={"displayModeBar": False})

    s3, s4 = st.columns(2)
    with s3:
        age = fdf["_age"].value_counts().sort_index()
        fig_a = go.Figure(go.Bar(x=age.index, y=age.values, marker_color="#58a6ff",
                                  text=age.values, textposition="outside", textfont=dict(color="#e6edf3", size=12)))
        fig_layout(fig_a, 280)
        fig_a.update_layout(title=dict(text="Range Usia", font=dict(color="#8b949e", size=13)))
        fig_a.update_xaxes(tickangle=-30, tickfont=dict(color="#c9d1d9", size=11))
        st.plotly_chart(fig_a, width='stretch', key="chart_26", config={"displayModeBar": False})
    with s4:
        ten = fdf["_tenure"].value_counts()
        fig_t = go.Figure(go.Bar(x=ten.values, y=ten.index, orientation="h",
                                  marker_color="#3fb950", text=ten.values,
                                  textposition="outside", textfont=dict(color="#e6edf3", size=12)))
        fig_layout(fig_t, 280)
        fig_t.update_layout(title=dict(text="Lama Menjadi Nasabah", font=dict(color="#8b949e", size=13)))
        fig_t.update_yaxes(autorange="reversed", tickfont=dict(color="#c9d1d9", size=11))
        st.plotly_chart(fig_t, width='stretch', key="chart_27", config={"displayModeBar": False})

    st.markdown("---")
    section("Profil Sosial Ekonomi")
    se1, se2 = st.columns(2)
    with se1:
        edu_order = ["SD","SLTP","SLTA","Akademi/D3","Sarjana (S1)","Paska Sarjana (S2)","Doktor (S3)"]
        edu = fdf["_edu"].value_counts().reindex([e for e in edu_order if e in fdf["_edu"].values])
        fig_edu = go.Figure(go.Bar(x=edu.values, y=edu.index, orientation="h",
                                    marker_color="#e3b341", text=edu.values,
                                    textposition="outside", textfont=dict(color="#e6edf3", size=12)))
        fig_layout(fig_edu, 300)
        fig_edu.update_layout(title=dict(text="Pendidikan Terakhir", font=dict(color="#8b949e", size=13)))
        fig_edu.update_yaxes(autorange="reversed", tickfont=dict(color="#c9d1d9", size=11))
        st.plotly_chart(fig_edu, width='stretch', key="chart_28", config={"displayModeBar": False})
    with se2:
        job = fdf["_job"].value_counts().head(8)
        fig_j = go.Figure(go.Bar(x=job.values, y=job.index, orientation="h",
                                  marker_color="#bc8cff", text=job.values,
                                  textposition="outside", textfont=dict(color="#e6edf3", size=12)))
        fig_layout(fig_j, 300)
        fig_j.update_layout(title=dict(text="Jenis Pekerjaan", font=dict(color="#8b949e", size=13)))
        fig_j.update_yaxes(autorange="reversed", tickfont=dict(color="#c9d1d9", size=11))
        st.plotly_chart(fig_j, width='stretch', key="chart_29", config={"displayModeBar": False})

    se3, se4 = st.columns(2)
    with se3:
        # Clean income labels: normalize "Rp." -> "Rp" and strip spaces
        inc_raw = fdf["_income"].astype(str).str.strip().replace([" ", "nan"], None).dropna()
        inc_raw = inc_raw[inc_raw.str.len() > 3]
        inc_clean = inc_raw.str.replace(r"Rp\.\s*", "Rp ", regex=True).str.replace(r"\s+", " ", regex=True).str.strip()
        inc = inc_clean.value_counts().head(10)
        fig_i = go.Figure(go.Bar(x=inc.values, y=inc.index, orientation="h",
                                  marker_color="#58a6ff", text=inc.values,
                                  textposition="outside", textfont=dict(color="#e6edf3", size=11)))
        fig_layout(fig_i, 360)
        fig_i.update_layout(title=dict(text="Penghasilan Rumah Tangga / Bulan", font=dict(color="#8b949e", size=13)),
                             margin=dict(t=30, b=8, l=0, r=60))
        fig_i.update_yaxes(autorange="reversed", tickfont=dict(color="#c9d1d9", size=10))
        if len(inc) > 0:
            fig_i.update_xaxes(range=[0, inc.values.max() * 1.25])
        st.plotly_chart(fig_i, width='stretch', key="chart_30", config={"displayModeBar": False})
    with se4:
        # Clean expense: remove SES code prefix like "A2 : ", "B : " etc
        exp_raw = fdf["_expense"].astype(str).str.strip().replace([" ", "nan"], None).dropna()
        exp_raw = exp_raw[exp_raw.str.len() > 3]
        exp_clean = exp_raw.str.replace(r"^[A-Z0-9\.]+\s*:\s*", "", regex=True).str.strip()
        exp = exp_clean.value_counts().head(10)
        fig_ex = go.Figure(go.Bar(x=exp.values, y=exp.index, orientation="h",
                                   marker_color="#39d353", text=exp.values,
                                   textposition="outside", textfont=dict(color="#e6edf3", size=11)))
        fig_layout(fig_ex, 360)
        fig_ex.update_layout(title=dict(text="Pengeluaran Rutin / Bulan", font=dict(color="#8b949e", size=13)),
                              margin=dict(t=30, b=8, l=0, r=60))
        fig_ex.update_yaxes(autorange="reversed", tickfont=dict(color="#c9d1d9", size=10))
        if len(exp) > 0:
            fig_ex.update_xaxes(range=[0, exp.values.max() * 1.25])
        st.plotly_chart(fig_ex, width='stretch', key="chart_31", config={"displayModeBar": False})

    st.markdown("---")
    section("NPS per Segmen")
    ns1, ns2, ns3 = st.columns(3)

    def nps_bar(series_name, label_col):
        grp = fdf.groupby(label_col)["_nps"].apply(nps_score).reset_index()
        grp.columns = ["seg","NPS"]
        grp = grp.sort_values("NPS")
        fig = go.Figure(go.Bar(
            x=grp["NPS"], y=grp["seg"], orientation="h",
            marker_color=["#f85149" if v<50 else "#e3b341" if v<70 else "#3fb950" for v in grp["NPS"]],
            text=grp["NPS"].round(1), textposition="outside", textfont=dict(color="#e6edf3", size=12),
        ))
        fig_layout(fig, 300)
        fig.update_yaxes(tickfont=dict(color="#c9d1d9", size=10))
        return fig, grp

    with ns1:
        fig_na, age_grp = nps_bar("Usia", "_age")
        fig_na.update_layout(title=dict(text="NPS per Range Usia", font=dict(color="#8b949e", size=13)))
        st.plotly_chart(fig_na, width='stretch', key="chart_32", config={"displayModeBar": False})
    with ns2:
        fig_nj, job_grp = nps_bar("Pekerjaan", "_job")
        fig_nj.update_layout(title=dict(text="NPS per Pekerjaan", font=dict(color="#8b949e", size=13)))
        st.plotly_chart(fig_nj, width='stretch', key="chart_33", config={"displayModeBar": False})
    with ns3:
        fig_nt, ten_grp = nps_bar("Tenure", "_tenure")
        fig_nt.update_layout(title=dict(text="NPS per Lama Nasabah", font=dict(color="#8b949e", size=13)))
        st.plotly_chart(fig_nt, width='stretch', key="chart_34", config={"displayModeBar": False})

    section("💡 Insight Segmentasi")
    si1, si2 = st.columns(2)
    best_age  = age_grp.nlargest(1,"NPS").iloc[0]
    worst_age = age_grp.nsmallest(1,"NPS").iloc[0]
    best_job  = job_grp.nlargest(1,"NPS").iloc[0]
    with si1:
        insight("🎯","Segmen Paling Loyal",
                f"Usia <b>{best_age['seg']}</b> (NPS <b>{best_age['NPS']:.1f}</b>) dan profesi <b>{best_job['seg']}</b> paling puas. Kembangkan program khusus untuk segmen ini sebagai brand advocate.")
        insight("📊","Profil Dominan",
                f"Mayoritas adalah <b>{fdf['_edu'].value_counts().index[0]}</b>, bekerja sebagai <b>{fdf['_job'].value_counts().index[0]}</b>. Sesuaikan komunikasi produk dengan profil ini.")
    with si2:
        insight("⚠️","Segmen Perlu Perhatian",
                f"Usia <b>{worst_age['seg']}</b> memiliki NPS terendah (<b>{worst_age['NPS']:.1f}</b>). Rancang pendekatan layanan yang lebih relevan untuk segmen ini.")
        insight("💰","Profil Ekonomi",
                f"Kelompok penghasilan dominan: <b>{fdf['_income'].value_counts().index[0]}</b>. Sesuaikan produk dan promosi dengan kapasitas finansial segmen utama untuk meningkatkan engagement.")

st.markdown("<div style='text-align:center;color:#484f58;font-size:12px;font-weight:500;padding:20px 0 12px;border-top:1px solid #21262d;margin-top:16px'>Dashboard Bank XYZ &nbsp;·&nbsp; Data Challenge 2026 &nbsp;·&nbsp; Kelompok 1</div>", unsafe_allow_html=True)
