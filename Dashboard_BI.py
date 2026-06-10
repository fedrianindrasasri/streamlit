"""
Dashboard Kualitas Udara Jakarta 2022 — ISPU
=============================================
Dashboard BI interaktif untuk masyarakat umum, dibangun dengan Streamlit + Plotly.

Cara menjalankan:
    pip install streamlit plotly pandas openpyxl
    streamlit run dashboard_ispu_streamlit.py

Letakkan file Excel ISPU di folder yang sama, atau unggah lewat sidebar.

Arah desain (frontend-design):
    Subjek = warna langit di atas kita. Hero adalah "pita cakrawala" yang
    warnanya mengikuti kategori ISPU terkini; kalender 365 hari menjadi
    tekstur langit setahun. Tipografi: Space Grotesk (display/angka),
    Inter (body), Space Mono (label monitoring).
"""

import os
import glob
import pandas as pd
import numpy as np
import streamlit as st
import plotly.graph_objects as go

# ──────────────────────────────────────────────────────────────────────────
# Konstanta domain (vernakular subjek: skala resmi ISPU)
# ──────────────────────────────────────────────────────────────────────────
KAT_ORDER = ["BAIK", "SEDANG", "TIDAK SEHAT", "SANGAT TIDAK SEHAT", "BERBAHAYA"]
KAT_COLOR = {
    "BAIK": "#1FA971",
    "SEDANG": "#2E7FC4",
    "TIDAK SEHAT": "#E69A2E",
    "SANGAT TIDAK SEHAT": "#D1483B",
    "BERBAHAYA": "#2B2733",
}
# Gradien atmosferik per kategori (terang -> gelap) untuk pita cakrawala
KAT_SKY = {
    "BAIK": ("#27C489", "#0E6E48"),
    "SEDANG": ("#4596DA", "#1C5689"),
    "TIDAK SEHAT": ("#F0AC46", "#B06E12"),
    "SANGAT TIDAK SEHAT": ("#E25B4C", "#8E2A20"),
    "BERBAHAYA": ("#544E5C", "#221F29"),
}
KAT_REC = {
    "BAIK": ["Udara sehat — nikmati aktivitas luar ruangan.",
             "Waktu tepat untuk berolahraga dan membuka ventilasi alami."],
    "SEDANG": ["Aman untuk umum; aktivitas normal bisa dijalankan.",
               "Kelompok sangat sensitif kurangi aktivitas berat berkepanjangan."],
    "TIDAK SEHAT": ["Kurangi aktivitas berat di luar ruangan.",
                    "Kelompok sensitif (anak, lansia, penderita asma) gunakan masker.",
                    "Tutup jendela saat polusi memuncak."],
    "SANGAT TIDAK SEHAT": ["Hindari aktivitas di luar ruangan.",
                           "Gunakan masker N95 bila harus keluar.",
                           "Nyalakan penyaring udara di dalam ruangan."],
    "BERBAHAYA": ["Tetap berada di dalam ruangan.",
                  "Tutup semua ventilasi ke luar.",
                  "Ikuti arahan otoritas kesehatan setempat."],
}
AREA = {"DKI1": "Bundaran HI", "DKI2": "Kelapa Gading", "DKI3": "Jagakarsa",
        "DKI4": "Lubang Buaya", "DKI5": "Kebon Jeruk"}
BULAN = ["", "Jan", "Feb", "Mar", "Apr", "Mei", "Jun",
         "Jul", "Agu", "Sep", "Okt", "Nov", "Des"]
POLL_COLS = ["pm_10", "pm_duakomalima", "so2", "co", "o3", "no2"]
POLL_NAME = {"pm_10": "PM10", "pm_duakomalima": "PM2,5", "so2": "SO2",
             "co": "CO", "o3": "O3", "no2": "NO2"}
POLL_COLOR = {"PM2,5": "#C2410C", "PM10": "#EA580C", "O3": "#E69A2E",
              "SO2": "#2E7FC4", "CO": "#0E8C9E", "NO2": "#7C5BD0"}

# ──────────────────────────────────────────────────────────────────────────
# Halaman + tipografi & token desain
# ──────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Kualitas Udara Jakarta 2022",
                   page_icon="🌫️", layout="wide",
                   initial_sidebar_state="expanded")

CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;700&family=Inter:wght@400;500;600&family=Space+Mono:wght@400;700&display=swap');

:root{
  --kabut:#EEF0F4; --ink:#161B26; --ink-2:#5E6B7E; --card:#FFFFFF;
  --hair:#E2E6EC; --baik:#1FA971; --sedang:#2E7FC4; --tidak:#E69A2E;
  --sangat:#D1483B; --bahaya:#2B2733;
}
html, body, [class*="css"], .stApp{ font-family:'Inter',sans-serif; }
.stApp{ background:var(--kabut); color:var(--ink); }
.block-container{ padding-top:1.4rem; padding-bottom:3rem; max-width:1280px; }
#MainMenu, footer, header[data-testid="stHeader"]{ visibility:hidden; height:0; }

h1,h2,h3,h4{ font-family:'Space Grotesk',sans-serif; letter-spacing:-.5px; color:var(--ink); }

/* eyebrow / utility line */
.eyebrow{ font-family:'Space Mono',monospace; font-size:11px; letter-spacing:2px;
  text-transform:uppercase; color:var(--ink-2); }

/* HERO — pita cakrawala */
.hero{ border-radius:20px; padding:34px 38px; color:#fff; position:relative;
  overflow:hidden; box-shadow:0 18px 40px -18px rgba(20,30,50,.45); }
.hero::after{ content:""; position:absolute; inset:0;
  background:radial-gradient(120% 140% at 85% -10%, rgba(255,255,255,.28), transparent 55%);
  pointer-events:none; }
.hero .eyebrow{ color:rgba(255,255,255,.82); }
.hero .num{ font-family:'Space Grotesk',sans-serif; font-weight:700;
  font-size:96px; line-height:.9; letter-spacing:-4px; margin:6px 0 0; }
.hero .cat{ font-family:'Space Grotesk',sans-serif; font-weight:500; font-size:26px; }
.hero .meta{ font-family:'Space Mono',monospace; font-size:12.5px;
  color:rgba(255,255,255,.88); margin-top:12px; letter-spacing:.3px; }
.scale{ display:flex; height:10px; border-radius:99px; overflow:hidden;
  margin-top:20px; border:1px solid rgba(255,255,255,.35); max-width:460px; }
.scale span{ display:block; height:100%; }
.scale-x{ display:flex; justify-content:space-between; max-width:460px;
  font-family:'Space Mono',monospace; font-size:10px;
  color:rgba(255,255,255,.8); margin-top:5px; }

/* recommendation card */
.rec{ background:var(--card); border-radius:20px; padding:28px 30px;
  border:1px solid var(--hair); height:100%; }
.rec h4{ font-size:16px; margin:0 0 16px; display:flex; align-items:center; gap:10px;}
.rec .dot{ width:13px; height:13px; border-radius:50%; display:inline-block; }
.rec ul{ margin:0; padding:0; list-style:none; display:flex;
  flex-direction:column; gap:13px; }
.rec li{ font-size:14.5px; line-height:1.45; padding-left:24px; position:relative; color:var(--ink); }
.rec li::before{ content:"→"; position:absolute; left:0; font-family:'Space Mono',monospace; }

/* KPI cards */
.kpi{ background:var(--card); border:1px solid var(--hair); border-radius:16px;
  padding:20px 22px; border-top:4px solid var(--ink-2); height:100%; }
.kpi .lbl{ font-family:'Space Mono',monospace; font-size:11px; letter-spacing:1px;
  text-transform:uppercase; color:var(--ink-2); }
.kpi .v{ font-family:'Space Grotesk',sans-serif; font-weight:700; font-size:40px;
  line-height:1; margin:10px 0 6px; letter-spacing:-2px; }
.kpi .s{ font-size:12.5px; color:var(--ink-2); }

/* section panel */
.panel{ background:var(--card); border:1px solid var(--hair); border-radius:18px;
  padding:22px 26px; }
.panel h3{ font-size:17px; margin:0 0 2px; }
.panel .sub{ font-size:12.5px; color:var(--ink-2); margin-bottom:14px;
  font-family:'Space Mono',monospace; letter-spacing:.3px; }
.legend{ display:flex; gap:18px; flex-wrap:wrap; font-size:12px; color:var(--ink-2);
  font-family:'Space Mono',monospace; }
.legend i{ width:12px; height:12px; border-radius:3px; display:inline-block;
  margin-right:6px; vertical-align:-1px; }

/* sidebar */
section[data-testid="stSidebar"]{ background:var(--ink); }
section[data-testid="stSidebar"] *{ color:#D7DEE8 !important; }
section[data-testid="stSidebar"] h2{ font-family:'Space Grotesk',sans-serif; color:#fff !important; }
.stDataFrame{ border-radius:12px; overflow:hidden; }
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────
# Data: muat + validasi + bersihkan (cached)
# ──────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_clean(file) -> pd.DataFrame:
    df = pd.read_excel(file)
    df["tanggal"] = pd.to_datetime(df["tanggal"], errors="coerce")

    # 1) perbaiki tahun yang salah ketik (mis. 2020 -> ikut tahun periode)
    per_year = df["periode_data"] // 100
    mask = df["tanggal"].notna() & (df["tanggal"].dt.year != per_year)
    df.loc[mask, "tanggal"] = df.loc[mask, "tanggal"].apply(lambda d: d.replace(year=2022))

    # 2) hitung ulang 'critical' yang kosong dari polutan tertinggi
    def crit(r):
        if isinstance(r["critical"], str) and r["critical"].strip():
            return r["critical"]
        col = r[POLL_COLS].astype(float).idxmax()
        return POLL_NAME[col]
    df["critical"] = df.apply(crit, axis=1)

    # 3) buang baris tanpa tanggal valid atau lokasi tidak dikenal
    df = df[df["lokasi_spku"].isin(AREA) & df["tanggal"].notna()].copy()
    df["wilayah"] = df["lokasi_spku"].map(AREA)
    df["bln"] = df["tanggal"].dt.month
    df["ispu"] = df["max"].astype(int)
    df["categori"] = pd.Categorical(df["categori"], categories=KAT_ORDER, ordered=True)
    return df.sort_values("tanggal").reset_index(drop=True)


def find_default():
    cands = glob.glob("*ISPU*xlsx") + glob.glob("*Pencemaran*xlsx") + glob.glob("Filedata*xlsx")
    return cands[0] if cands else None


# ──────────────────────────────────────────────────────────────────────────
# Sidebar: sumber data + filter
# ──────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## Kualitas Udara")
    st.markdown('<span class="eyebrow">Stasiun pemantau · DKI Jakarta</span>',
                unsafe_allow_html=True)
    st.markdown("---")

    default = find_default()
    src = st.file_uploader("Sumber data (Excel ISPU)", type=["xlsx"])
    if src is None and default:
        src = default
        st.caption(f"Memakai berkas: `{os.path.basename(default)}`")

if src is None:
    st.info("Unggah berkas Excel ISPU pada panel kiri untuk memulai.")
    st.stop()

df = load_clean(src)

with st.sidebar:
    st.markdown("### Filter")
    wil_opt = ["Semua wilayah"] + [f"{AREA[c]} ({c})" for c in sorted(df["lokasi_spku"].unique())]
    f_wil = st.selectbox("Wilayah", wil_opt)
    bln_opt = ["Sepanjang tahun"] + [BULAN[m] for m in range(1, 13)]
    f_bln = st.selectbox("Bulan", bln_opt)
    kat_opt = ["Semua kategori"] + [k for k in KAT_ORDER if k in df["categori"].values]
    f_kat = st.selectbox("Kategori", kat_opt,
                         format_func=lambda x: x if x.startswith("Semua") else x.title())

# terapkan filter
d = df.copy()
if not f_wil.startswith("Semua"):
    code = f_wil.split("(")[-1].strip(")")
    d = d[d["lokasi_spku"] == code]
if not f_bln.startswith("Sepanjang"):
    d = d[d["bln"] == BULAN.index(f_bln)]
if not f_kat.startswith("Semua"):
    d = d[d["categori"] == f_kat]

if d.empty:
    st.warning("Tidak ada data untuk kombinasi filter ini. Coba longgarkan filternya.")
    st.stop()

# ISPU tingkat kota per hari = nilai terburuk antar stasiun
daily = (d.sort_values("ispu").drop_duplicates("tanggal", keep="last")
           .sort_values("tanggal"))


def fmt_tgl(ts):
    return f"{ts.day} {BULAN[ts.month]} {ts.year}"


# ──────────────────────────────────────────────────────────────────────────
# HERO — pita cakrawala (signature)
# ──────────────────────────────────────────────────────────────────────────
st.markdown('<span class="eyebrow">Indeks Standar Pencemaran Udara · Jakarta · 2022</span>',
            unsafe_allow_html=True)
st.markdown("<h1 style='margin:.1rem 0 1rem;font-size:30px;'>Seberapa bersih udara kita?</h1>",
            unsafe_allow_html=True)

last = daily.iloc[-1]
kat = last["categori"]
sky_a, sky_b = KAT_SKY[kat]
col_hero, col_rec = st.columns([1.25, 1], gap="medium")

with col_hero:
    st.markdown(f"""
    <div class="hero" style="background:linear-gradient(135deg,{sky_a},{sky_b});">
      <span class="eyebrow">ISPU terkini · {fmt_tgl(last['tanggal'])}</span>
      <div class="num">{last['ispu']}</div>
      <div class="cat">{str(kat).title()}</div>
      <div class="meta">Polutan kritis {last['critical']} · {last['wilayah']} ({last['lokasi_spku']})</div>
      <div class="scale">
        <span style="flex:1;background:var(--baik)"></span>
        <span style="flex:1;background:var(--sedang)"></span>
        <span style="flex:2;background:var(--tidak)"></span>
        <span style="flex:1;background:var(--sangat)"></span>
        <span style="flex:.6;background:var(--bahaya)"></span>
      </div>
      <div class="scale-x"><span>0</span><span>50</span><span>100</span><span>200</span><span>300+</span></div>
    </div>""", unsafe_allow_html=True)

with col_rec:
    recs = "".join(f"<li>{t}</li>" for t in KAT_REC[kat])
    st.markdown(f"""
    <div class="rec">
      <h4><span class="dot" style="background:{KAT_COLOR[kat]}"></span>Yang sebaiknya dilakukan</h4>
      <ul>{recs}</ul>
    </div>""", unsafe_allow_html=True)

st.write("")

# ──────────────────────────────────────────────────────────────────────────
# KPI
# ──────────────────────────────────────────────────────────────────────────
n = len(d)
tidak = int((d["categori"].map(lambda k: KAT_ORDER.index(k) >= 2)).sum())
baik = int((d["categori"] == "BAIK").sum())
mx = d.loc[d["ispu"].idxmax()]
avg = round(d["ispu"].mean())
dom = d["critical"].value_counts()
dom_name, dom_n = dom.index[0], int(dom.iloc[0])

kpis = [
    ("Hari tidak sehat", f"{tidak/n*100:.1f}%", f"{tidak} dari {n} hari", "#E69A2E"),
    ("Hari baik", f"{baik}", 'kategori "Baik"', "#1FA971"),
    ("ISPU tertinggi", f"{int(mx['ispu'])}", fmt_tgl(mx["tanggal"]), "#D1483B"),
    ("Polutan dominan", dom_name, f"{round(dom_n/n*100)}% hari · rata² ISPU {avg}", "#2E7FC4"),
]
for col, (lbl, v, s, c) in zip(st.columns(4, gap="medium"), kpis):
    col.markdown(f"""<div class="kpi" style="border-top-color:{c}">
        <div class="lbl">{lbl}</div><div class="v" style="color:{c}">{v}</div>
        <div class="s">{s}</div></div>""", unsafe_allow_html=True)

st.write("")

# ──────────────────────────────────────────────────────────────────────────
# Plotly helper styling
# ──────────────────────────────────────────────────────────────────────────
def style_fig(fig, h=300):
    fig.update_layout(
        height=h, margin=dict(l=8, r=8, t=8, b=8),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter, sans-serif", size=12, color="#161B26"),
        showlegend=False, hoverlabel=dict(font_family="Space Mono, monospace"),
    )
    fig.update_xaxes(gridcolor="#EEF0F4", zeroline=False)
    fig.update_yaxes(gridcolor="#EEF0F4", zeroline=False)
    return fig


# ──────────────────────────────────────────────────────────────────────────
# KALENDER — tekstur langit setahun (signature)
# ──────────────────────────────────────────────────────────────────────────
st.markdown('<div class="panel">', unsafe_allow_html=True)
cleg = ("".join(f'<span><i style="background:{KAT_COLOR[k]}"></i>{k.title()}</span>'
                for k in KAT_ORDER if k in df["categori"].values))
st.markdown(f"""<div style="display:flex;justify-content:space-between;align-items:flex-start;flex-wrap:wrap;gap:10px">
  <div><h3>Setahun langit</h3>
  <div class="sub">{len(daily)} hari tercatat · tiap kotak = kategori udara terburuk hari itu</div></div>
  <div class="legend">{cleg}</div></div>""", unsafe_allow_html=True)

dmap = {r["tanggal"].strftime("%Y-%m-%d"): r for _, r in daily.iterrows()}
z, hov = [], []
for m in range(1, 13):
    zr, hr = [], []
    for day in range(1, 32):
        try:
            key = f"2022-{m:02d}-{day:02d}"
            pd.Timestamp(key)
        except ValueError:
            zr.append(None); hr.append(""); continue
        r = dmap.get(key)
        if r is None:
            zr.append(None); hr.append("")
        else:
            zr.append(KAT_ORDER.index(r["categori"]))
            hr.append(f"{day} {BULAN[m]}<br>ISPU {r['ispu']} · {str(r['categori']).title()}<br>Kritis: {r['critical']}")
    z.append(zr); hov.append(hr)

cs = [[0.0, "#1FA971"], [0.2, "#1FA971"], [0.2, "#2E7FC4"], [0.4, "#2E7FC4"],
      [0.4, "#E69A2E"], [0.6, "#E69A2E"], [0.6, "#D1483B"], [0.8, "#D1483B"],
      [0.8, "#2B2733"], [1.0, "#2B2733"]]
cal = go.Figure(go.Heatmap(
    z=z, customdata=hov, hovertemplate="%{customdata}<extra></extra>",
    colorscale=cs, zmin=-0.5, zmax=4.5, showscale=False, xgap=3, ygap=3))
cal.update_yaxes(tickvals=list(range(12)), ticktext=BULAN[1:],
                 autorange="reversed", showgrid=False)
cal.update_xaxes(tickvals=[0, 4, 9, 14, 19, 24, 29],
                 ticktext=["1", "5", "10", "15", "20", "25", "30"], showgrid=False)
cal.update_layout(height=340, margin=dict(l=8, r=8, t=8, b=8),
                  paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                  font=dict(family="Space Mono, monospace", size=11, color="#5E6B7E"),
                  hoverlabel=dict(font_family="Space Mono, monospace"))
st.plotly_chart(cal, width='stretch', config={"displayModeBar": False})
st.markdown("</div>", unsafe_allow_html=True)

st.write("")

# ──────────────────────────────────────────────────────────────────────────
# Baris grafik 1: tren bulanan + donut kategori
# ──────────────────────────────────────────────────────────────────────────
c1, c2 = st.columns(2, gap="medium")

with c1:
    st.markdown('<div class="panel"><h3>Rata-rata ISPU per bulan</h3>'
                '<div class="sub">Garis = ambang tidak sehat (100)</div>', unsafe_allow_html=True)
    mvals, mcol = [], []
    for m in range(1, 13):
        rows = d[d["bln"] == m]
        v = round(rows["ispu"].mean()) if len(rows) else None
        mvals.append(v)
        mcol.append("#E69A2E" if (v or 0) > 100 else "#1FA971" if (v is not None and v <= 50) else "#2E7FC4")
    fig = go.Figure(go.Bar(x=BULAN[1:], y=mvals, marker_color=mcol,
                           marker_line_width=0,
                           hovertemplate="%{x}: rata² %{y}<extra></extra>"))
    fig.add_hline(y=100, line_dash="dash", line_color="#D1483B", line_width=1.5)
    fig.update_traces(marker=dict(cornerradius=6))
    style_fig(fig)
    st.plotly_chart(fig, width='stretch', config={"displayModeBar": False})
    st.markdown("</div>", unsafe_allow_html=True)

with c2:
    st.markdown('<div class="panel"><h3>Sebaran kategori</h3>'
                '<div class="sub">Proporsi hari menurut kualitas udara</div>', unsafe_allow_html=True)
    present = [k for k in KAT_ORDER if (d["categori"] == k).sum() > 0]
    vals = [int((d["categori"] == k).sum()) for k in present]
    fig = go.Figure(go.Pie(labels=[k.title() for k in present], values=vals, hole=.6,
                           marker=dict(colors=[KAT_COLOR[k] for k in present],
                                       line=dict(color="#fff", width=3)),
                           sort=False, direction="clockwise",
                           textinfo="percent", textfont=dict(family="Space Grotesk", size=13),
                           hovertemplate="%{label}: %{value} hari (%{percent})<extra></extra>"))
    fig.update_layout(height=300, margin=dict(l=8, r=8, t=8, b=8),
                      paper_bgcolor="rgba(0,0,0,0)", showlegend=True,
                      legend=dict(orientation="v", x=1, y=.5, font=dict(size=12)),
                      font=dict(family="Inter, sans-serif", color="#161B26"),
                      annotations=[dict(text=f"<b>{n}</b><br>hari", x=.5, y=.5,
                                        showarrow=False, font=dict(family="Space Grotesk", size=20))])
    st.plotly_chart(fig, width='stretch', config={"displayModeBar": False})
    st.markdown("</div>", unsafe_allow_html=True)

st.write("")

# ──────────────────────────────────────────────────────────────────────────
# Baris grafik 2: per wilayah + polutan dominan
# ──────────────────────────────────────────────────────────────────────────
c3, c4 = st.columns(2, gap="medium")

with c3:
    st.markdown('<div class="panel"><h3>Kualitas udara per wilayah</h3>'
                '<div class="sub">% hari tidak sehat per stasiun pemantau</div>', unsafe_allow_html=True)
    stats = []
    for w, g in d.groupby("wilayah"):
        pct = (g["categori"].map(lambda k: KAT_ORDER.index(k) >= 2)).mean() * 100
        stats.append((f"{w} ({g['lokasi_spku'].iloc[0]})", round(pct), len(g)))
    stats.sort(key=lambda x: x[1])
    fig = go.Figure(go.Bar(
        y=[s[0] for s in stats], x=[s[1] for s in stats], orientation="h",
        marker_color=["#E69A2E" if s[1] >= 50 else "#2E7FC4" if s[1] > 0 else "#1FA971" for s in stats],
        marker=dict(cornerradius=6),
        customdata=[s[2] for s in stats],
        hovertemplate="%{y}<br>%{x}% tidak sehat (n=%{customdata})<extra></extra>"))
    style_fig(fig)
    fig.update_xaxes(ticksuffix="%", range=[0, max([s[1] for s in stats]) * 1.25 + 5])
    st.plotly_chart(fig, width='stretch', config={"displayModeBar": False})
    st.markdown("</div>", unsafe_allow_html=True)

with c4:
    st.markdown('<div class="panel"><h3>Penyebab utama polusi</h3>'
                '<div class="sub">Polutan kritis penentu nilai ISPU</div>', unsafe_allow_html=True)
    pc = d["critical"].value_counts()
    fig = go.Figure(go.Bar(
        y=list(pc.index)[::-1], x=(pc.values / n * 100).round()[::-1], orientation="h",
        marker_color=[POLL_COLOR.get(p, "#5E6B7E") for p in list(pc.index)[::-1]],
        marker=dict(cornerradius=6),
        customdata=list(pc.values)[::-1],
        hovertemplate="%{y}: %{x}% hari (%{customdata} hari)<extra></extra>"))
    style_fig(fig)
    fig.update_xaxes(ticksuffix="%", range=[0, 108])
    st.plotly_chart(fig, width='stretch', config={"displayModeBar": False})
    st.markdown("</div>", unsafe_allow_html=True)

st.write("")

# ──────────────────────────────────────────────────────────────────────────
# Tabel rinci (sortable bawaan st.dataframe)
# ──────────────────────────────────────────────────────────────────────────
st.markdown('<div class="panel"><h3>Data harian rinci</h3>'
            '<div class="sub">Klik judul kolom untuk mengurutkan</div>', unsafe_allow_html=True)
tbl = d[["tanggal", "wilayah", "lokasi_spku", "categori", "critical",
         "pm_duakomalima", "o3", "ispu"]].copy()
tbl["tanggal"] = tbl["tanggal"].dt.strftime("%Y-%m-%d")
tbl["categori"] = tbl["categori"].astype(str).str.title()
tbl = tbl.rename(columns={"tanggal": "Tanggal", "wilayah": "Wilayah",
                          "lokasi_spku": "Stasiun", "categori": "Kategori",
                          "critical": "Polutan kritis", "pm_duakomalima": "PM2,5",
                          "o3": "O3", "ispu": "ISPU"})
st.dataframe(tbl, width='stretch', hide_index=True, height=380,
             column_config={"ISPU": st.column_config.NumberColumn(format="%d")})
st.markdown("</div>", unsafe_allow_html=True)

st.markdown(
    '<div style="text-align:center;margin-top:22px" class="eyebrow">'
    'Sumber: data ISPU 2022 · SPKU DKI Jakarta &nbsp;·&nbsp; data divalidasi & dibersihkan &nbsp;·&nbsp; dibuat untuk edukasi publik'
    '</div>', unsafe_allow_html=True)