"""
==============================================================================
🌷 GEOMETRI JALAN INTERAKTIF - TULIP EDITION
Standar Acuan: TPGJAK Bina Marga 1997 / Pedoman Desain Geometrik Jalan RAE 2021
==============================================================================
"""

import io
import math
from datetime import date
import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from docx import Document
from docx.shared import Mm, Pt, RGBColor
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# ----------------------------------------------------------------------------
# 1. KONFIGURASI HALAMAN STREAMLIT & TEMA TULIP EDITION
# ----------------------------------------------------------------------------
st.set_page_config(
    page_title="Geometri Jalan Interaktif - Tulip Edition",
    layout="wide",
    page_icon="🌷",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;600;700&display=swap');
    
    .stApp {
        background: linear-gradient(135deg, #fff9e6 0%, #ffe6f0 50%, #e8f4f8 100%);
        color: #2D3142;
        font-family: 'Poppins', sans-serif;
    }
    section[data-testid="stSidebar"] {
        background-color: rgba(255, 235, 243, 0.98) !important;
        border-right: 2px solid #F8BBD0;
    }
    section[data-testid="stSidebar"] h1, 
    section[data-testid="stSidebar"] h2, 
    section[data-testid="stSidebar"] h3, 
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] span {
        color: #3A0056 !important;
        font-weight: 700 !important;
    }
    .stTabs [data-baseweb="tab-list"] {
        background-color: #FFFFFF !important;
        border-radius: 15px;
        padding: 8px;
        border: 2px solid #F48FB1;
        box-shadow: 0 4px 12px rgba(136, 14, 79, 0.08);
        gap: 6px;
    }
    .stTabs [data-baseweb="tab"],
    .stTabs [data-baseweb="tab"] p,
    .stTabs [data-baseweb="tab"] span,
    .stTabs button[role="tab"],
    .stTabs button[role="tab"] p {
        color: #880E4F !important;
        font-weight: 800 !important;
        font-size: 14.5px !important;
        border-radius: 10px;
        padding: 8px 16px;
        background-color: transparent !important;
    }
    .stTabs [data-baseweb="tab"]:hover,
    .stTabs button[role="tab"]:hover {
        background-color: #FCE4EC !important;
        color: #AD1457 !important;
    }
    .stTabs [aria-selected="true"],
    .stTabs button[aria-selected="true"] {
        background: linear-gradient(45deg, #AD1457, #D81B60) !important;
        box-shadow: 0px 4px 12px rgba(173, 20, 87, 0.35);
    }
    .stTabs [aria-selected="true"] p,
    .stTabs [aria-selected="true"] span,
    .stTabs button[aria-selected="true"] p {
        color: #FFFFFF !important;
    }
    div[data-baseweb="input"] {
        background-color: #FFFFFF !important;
        border-radius: 10px;
        color: #111111 !important;
        border: 1.5px solid #F48FB1 !important;
    }
    .stButton>button {
        background: linear-gradient(45deg, #C2185B, #E91E63);
        color: white !important;
        border: none;
        border-radius: 12px;
        font-weight: bold;
        box-shadow: 0px 4px 10px rgba(194, 24, 91, 0.3);
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background: linear-gradient(45deg, #880E4F, #C2185B);
        transform: translateY(-2px);
        box-shadow: 0px 6px 14px rgba(136, 14, 79, 0.4);
    }
    
    /* KPI Card Tulip Edition */
    .kpi-container {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(175px, 1fr));
        gap: 14px;
        margin-bottom: 22px;
    }
    .kpi-card {
        background: #FFFFFF;
        border: 1.5px solid #F48FB1;
        border-radius: 14px;
        padding: 14px 16px;
        text-align: left;
        position: relative;
        overflow: hidden;
        box-shadow: 0 4px 14px rgba(136, 14, 79, 0.07);
    }
    .kpi-card::before {
        content: '';
        position: absolute;
        top: 0; left: 0; right: 0; height: 3.5px;
        background: linear-gradient(90deg, #C2185B, #E91E63);
    }
    .kpi-title {
        font-size: 11px;
        font-weight: 800;
        color: #880E4F;
        text-transform: uppercase;
        letter-spacing: 0.6px;
        margin-bottom: 4px;
    }
    .kpi-value {
        font-size: 21px;
        font-weight: 800;
        color: #3A0056;
        font-family: 'JetBrains Mono', monospace;
        line-height: 1.2;
    }
    .kpi-sub {
        font-size: 11px;
        color: #AD1457;
        margin-top: 3px;
        font-weight: 600;
    }
    
    .tulip-card {
        background: #FFFFFF;
        border: 1.5px solid #F48FB1;
        border-radius: 16px;
        padding: 18px;
        box-shadow: 0 4px 16px rgba(136, 14, 79, 0.06);
        margin-bottom: 18px;
    }
    
    /* Status Cards */
    .status-card {
        padding: 15px;
        border-radius: 14px;
        margin-bottom: 15px;
        font-weight: 600;
        box-shadow: 0 4px 15px rgba(0,0,0,0.05);
    }
    .status-oke {
        background-color: rgba(200, 230, 201, 0.98);
        color: #1B5E20;
        border: 2px solid #81C784;
    }
    .status-sedang {
        background-color: rgba(255, 245, 157, 0.98);
        color: #E65100;
        border: 2px solid #FFF176;
    }
    .status-bahaya {
        background-color: rgba(255, 205, 210, 0.98);
        color: #B71C1C;
        border: 2px solid #E57373;
    }
</style>
""", unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# 2. KEAMANAN & PORTAL LOGIN TULIP
# ----------------------------------------------------------------------------
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    col_top1, col_top2 = st.columns([4, 1])
    with col_top2:
        st.markdown(
            "<p style='text-align: right; font-size: 13px; color: #880E4F; margin: 0;'>pw: "
            "<span style='font-family: monospace; font-weight: bold; color: #C2185B;'>cantikaaa27</span></p>",
            unsafe_allow_html=True
        )
    st.markdown("<h1 style='text-align: center; color: #3A0056;'>🌷 Perencanaan Geometrik Jalan Interaktif</h1>", unsafe_allow_html=True)
    st.markdown("<h4 style='text-align: center; color: #880E4F;'>Standar TPGJAK 1997 / RAE 2021</h4>", unsafe_allow_html=True)
    
    st.image(
        "https://images.unsplash.com/photo-1544620347-c4fd4a3d5957?q=80&w=1200&auto=format&fit=crop",
        caption="Sistem Geometri Jalan Interaktif (Tulip Edition)",
        use_container_width=True
    )
    st.markdown("---")
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown(
            """
            <div style="background-color: #FCE4EC; border: 1.5px solid #F48FB1; border-radius: 10px; 
                        padding: 12px 16px; margin-bottom: 12px; color: #880E4F; font-size: 14px; text-align: center;">
                💡 <b>Petunjuk password:</b> lihat bagian pojok kanan atas (<b style="color: #C2185B;">cantikaaa27</b>)
            </div>
            """,
            unsafe_allow_html=True
        )
        st.markdown("<p style='color: #3A0056; font-weight: 700; margin-bottom: 4px;'>Masukkan Password Akses:</p>", unsafe_allow_html=True)
        pwd_input = st.text_input(
            "Masukkan Password Akses:",
            type="password",
            placeholder="Ketik password di sini...",
            label_visibility="collapsed"
        )
        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            if st.button("Masuk Aplikasi 🚀", use_container_width=True):
                if pwd_input.strip() == "cantikaaa27":
                    st.session_state.authenticated = True
                    st.rerun()
                else:
                    st.error("❌ Password salah! Silakan periksa kembali tulisan di pojok kanan atas.")
        with col_btn2:
            if st.button("Akses Cepat (Demo) ⚡", use_container_width=True):
                st.session_state.authenticated = True
                st.rerun()
    st.stop()

# ----------------------------------------------------------------------------
# 3. ENGINE PERHITUNGAN GEOMETRIK & MATEMATIS (BM 1997 / RAE 2021)
# ----------------------------------------------------------------------------
V_TAB = [20, 30, 40, 50, 60, 80, 100, 120]
R_FC_TAB = [60, 130, 250, 350, 500, 900, 1500, 2500]

def sta_str(s: float, d: int = 2) -> str:
    sign = "-" if s < 0 else ""
    s_abs = abs(s)
    return f"{sign}{int(s_abs // 1000)}+{s_abs % 1000:0{4 + d}.{d}f}"

def f_max(V: float) -> float:
    return -0.00065 * V + 0.192 if V < 80 else -0.00125 * V + 0.24

def r_min(V: float, emax: float) -> float:
    return V**2 / (127 * (emax + f_max(V)))

def hitung_tikungan(V: float, delta: float, emax: float, en: float, Rc: float, sta_pi: float):
    fm = f_max(V)
    Rmin = r_min(V, emax)
    Dmax = 181913.53 * (emax + fm) / (V**2)
    D = 1432.39 / Rc
    e = -emax * (D**2) / (Dmax**2) + 2 * emax * D / Dmax
    e = max(e, en)
    re = 0.035 if V <= 70 else 0.025
    Ls_waktu = V * 3 / 3.6
    Ls_shortt = 0.022 * (V**3) / (Rc * 0.4) - 2.727 * V * e / 0.4
    Ls_re = (e + en) * V / (3.6 * re)
    Ls_min = max(Ls_waktu, Ls_shortt, Ls_re)
    Ls_pakai = math.ceil(Ls_min / 5) * 5
    R_fc = float(np.interp(V, V_TAB, R_FC_TAB))
    d2 = math.radians(delta / 2)
    
    res = dict(
        V=V, delta=delta, emax=emax, en=en, Rc=Rc, fm=fm, Rmin=Rmin,
        Dmax=Dmax, D=D, e=e, Ls_waktu=Ls_waktu, Ls_shortt=Ls_shortt,
        Ls_re=Ls_re, Ls_min=Ls_min, R_fc=R_fc, sta_pi=sta_pi, Ls_pakai=Ls_pakai
    )
    if Rc >= R_fc:
        Tc = Rc * math.tan(d2)
        Ec = Tc * math.tan(d2 / 2)
        Lc = math.radians(delta) * Rc
        res.update(
            jenis="Full Circle (FC)", Ls=0.0, Ls_sup=Ls_pakai, ths=0.0,
            dc=delta, Lc=Lc, T=Tc, E=Ec, Xs=0.0, Ys=0.0, p=0.0, k=0.0
        )
    else:
        ths = math.degrees(Ls_pakai / (2 * Rc))
        dc = delta - 2 * ths
        Lc = math.radians(dc) * Rc
        Ls = float(Ls_pakai)
        jenis = "Spiral-Circle-Spiral (S-C-S)"
        
        if Lc < 25 or dc <= 0:
            jenis = "Spiral-Spiral (S-S)"
            ths = delta / 2
            Ls = math.radians(ths) * 2 * Rc
            Lc, dc = 0.0, 0.0
        Xs = Ls * (1 - Ls**2 / (40 * Rc**2))
        Ys = (Ls**2) / (6 * Rc)
        p = Ys - Rc * (1 - math.cos(math.radians(ths)))
        k = Xs - Rc * math.sin(math.radians(ths))
        Ts = (Rc + p) * math.tan(d2) + k
        Es = (Rc + p) / math.cos(d2) - Rc
        res.update(
            jenis=jenis, Ls=Ls, Ls_sup=Ls, ths=ths, dc=dc, Lc=Lc, T=Ts,
            E=Es, Xs=Xs, Ys=Ys, p=p, k=k
        )
    res["Ltot"] = res["Lc"] + 2 * res["Ls"]
    res["sta_awal"] = sta_pi - res["T"]
    res["sta_akhir"] = res["sta_awal"] + res["Ltot"]
    res["v_safe"] = math.sqrt(127 * Rc * (res["e"] + fm))
    return res

def heading(s, Ls, Lc, Rc, dl):
    Ltot = 2 * Ls + Lc
    th = np.zeros_like(s)
    if Ls > 0:
        ths = Ls / (2 * Rc)
        m1 = (s > 0) & (s <= Ls)
        th[m1] = (s[m1]**2) / (2 * Rc * Ls)
        m2 = (s > Ls) & (s < Ls + Lc)
        th[m2] = ths + (s[m2] - Ls) / Rc
        m3 = (s >= Ls + Lc) & (s <= Ltot)
        th[m3] = dl - ((Ltot - s[m3])**2) / (2 * Rc * Ls)
    else:
        m = (s > 0) & (s <= Ltot)
        th[m] = s[m] / Rc
    th[s > Ltot] = dl
    return th

def hitung_profil(pvi_df, V):
    d = pvi_df.dropna().sort_values(pvi_df.columns[0]).reset_index(drop=True)
    if len(d) < 2:
        return None
    sta = d.iloc[:, 0].to_numpy(float)
    el = d.iloc[:, 1].to_numpy(float)
    lvm = d.iloc[:, 2].to_numpy(float)
    g = np.diff(el) / np.diff(sta)
    Jh = 0.278 * V * 2.5 + (V**2) / (254 * 0.40)
    kurva = []
    for i in range(1, len(sta) - 1):
        A = (g[i] - g[i - 1]) * 100
        if abs(A) < 1e-9:
            continue
        a = abs(A)
        if A < 0:
            L = a * (Jh**2) / 399
            if L < Jh:
                L = 2 * Jh - 399 / a
        else:
            L = a * (Jh**2) / (120 + 3.5 * Jh)
            if L < Jh:
                L = 2 * Jh - (120 + 3.5 * Jh) / a
        
        L = math.ceil(max(L, 0.6 * V, V * 3 / 3.6) / 10) * 10
        if lvm[i] > 0:
            L = lvm[i]
            
        kurva.append(dict(
            i=i, sta=sta[i], el=el[i], A=A, Lv=L, Ev=a * L / 800,
            bvc=sta[i] - L / 2, evc=sta[i] + L / 2,
            jenis="Cembung" if A < 0 else "Cekung",
            g1=g[i - 1], g2=g[i]
        ))
    x = np.arange(sta[0], sta[-1] + 0.01, 1.0)
    y = np.interp(x, sta, el)
    for k in kurva:
        m = (x >= k["bvc"]) & (x <= k["evc"])
        dx = x[m] - k["bvc"]
        yb = k["el"] - k["g1"] * k["Lv"] / 2
        y[m] = yb + k["g1"] * dx + (k["g2"] - k["g1"]) / (2 * k["Lv"]) * (dx**2)
        
    return dict(sta=sta, el=el, g=g * 100, kurva=kurva, x=x, y=y, Jh=Jh)

# ----------------------------------------------------------------------------
# 4. SIDEBAR INPUT & KONTROL
# ----------------------------------------------------------------------------
with st.sidebar:
    st.markdown("<h2 style='color: #3A0056;'>⚙️ Parameter Rencana</h2>", unsafe_allow_html=True)
    if st.button("🚪 Keluar / Logout", use_container_width=True):
        st.session_state.authenticated = False
        st.rerun()
        
    st.markdown("---")
    st.subheader("1. Kriteria Desain")
    V = st.number_input("Kecepatan Rencana Vr (km/jam)", 20, 120, 60, 10)
    v_sim = st.number_input("Kecepatan Kendaraan Simulasi (km/jam)", 10, 150, 65, 5)
    delta = st.number_input("Sudut Tikungan Δ (°)", 1.0, 179.0, 45.0, 0.5)
    emax = st.number_input("e maks (%)", 4.0, 12.0, 10.0, 0.5) / 100
    en = st.number_input("Kemiringan Normal en (%)", 1.0, 4.0, 2.0, 0.5) / 100
    
    Rmin = r_min(V, emax)
    st.caption(f"ℹ️ *R minimum acuan TPGJAK = {Rmin:.2f} m*")
    mode = st.radio("Jari-jari Tikungan Rc", ["Otomatis (Dibulatkan)", "Manual Input"], horizontal=True)
    if mode.startswith("Oto"):
        Rc = float(math.ceil(Rmin / 10) * 10)
    else:
        Rc = st.number_input("Rc (m)", 10.0, 5000.0, max(150.0, float(math.ceil(Rmin / 10) * 10)), 5.0)
        
    sta_pi = st.number_input("STA PI (m)", 0.0, 100000.0, 300.0, 10.0)
    arah = st.radio("Arah Tikungan", ["Kiri", "Kanan"], horizontal=True)
    B = st.number_input("Lebar Perkerasan Jalan B (m)", 3.0, 30.0, 7.0, 0.5)
    Bahu = st.number_input("Lebar Bahu Jalan (m)", 0.5, 5.0, 1.5, 0.25)

r = hitung_tikungan(V, delta, emax, en, Rc, sta_pi)
sg = 1 if arah == "Kiri" else -1
Ls, Lc = r["Ls"], r["Lc"]
dl = math.radians(r["delta"])
Ltot = r["Ltot"]
ext = math.ceil(max(120, 1.4 * r["T"]) / 10) * 10
s_vals = np.unique(np.concatenate([np.arange(-ext, Ltot + ext, 0.5), [0, Ls, Ls + Lc, Ltot]]))
th_vals = heading(s_vals, Ls, Lc, Rc, dl)
dS = np.diff(s_vals)
x_horiz = np.concatenate([[0], np.cumsum(dS * (np.cos(th_vals[:-1]) + np.cos(th_vals[1:])) / 2)])
y_horiz = np.concatenate([[0], np.cumsum(dS * (np.sin(th_vals[:-1]) + np.sin(th_vals[1:])) / 2)])
i0 = np.argmin(np.abs(s_vals))
x_horiz, y_horiz = x_horiz - x_horiz[i0], (y_horiz - y_horiz[i0]) * sg
ths_ = th_vals * sg
nx, ny = -np.sin(ths_), np.cos(ths_)

if "pvi_df" not in st.session_state:
    st.session_state.pvi_df = pd.DataFrame({
        "STA (m)": [0.0, 250.0, 500.0, 750.0],
        "Elevasi Rencana (m)": [100.0, 105.0, 101.0, 107.0],
        "Lv Manual (m)": [0.0, 0.0, 0.0, 0.0]
    })
if "tanah_df" not in st.session_state:
    st.session_state.tanah_df = pd.DataFrame({
        "STA (m)": [0.0, 125.0, 250.0, 375.0, 500.0, 625.0, 750.0],
        "Elevasi Tanah (m)": [101.0, 103.5, 104.0, 102.5, 101.8, 105.0, 106.5]
    })

P_global = hitung_profil(st.session_state.pvi_df, V)
if P_global is not None:
    z_elev = np.interp(s_vals, P_global["x"], P_global["y"])
else:
    z_elev = np.zeros_like(s_vals)

# ----------------------------------------------------------------------------
# 5. HEADER UTAMA & KPI EXECUTIVE CARDS (TULIP EDITION)
# ----------------------------------------------------------------------------
st.markdown("<h1 style='color: #3A0056; text-align: center; font-weight: 800; margin-bottom: 2px;'>🌸 Geometri Jalan Interaktif & Simulasi 3D</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #880E4F; font-weight: 600; margin-bottom: 20px;'>Standar TPGJAK Bina Marga 1997 / Pedoman Desain Geometrik Jalan RAE 2021</p>", unsafe_allow_html=True)

# KPI Cards Tulip Style
fc_flag = r["jenis"].startswith("Full")
lbl_start = "TC" if fc_flag else "TS"
lbl_end = "CT" if fc_flag else "ST"

kpi_html = f"""
<div class="kpi-container">
    <div class="kpi-card">
        <div class="kpi-title">Jenis Tikungan</div>
        <div class="kpi-value" style="font-size: 16px; color: #C2185B;">{r['jenis']}</div>
        <div class="kpi-sub">Arah: Belok {arah}</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-title">Kecepatan Vr / Aman</div>
        <div class="kpi-value">{r['V']:.0f} <span style="font-size: 14px;">km/h</span></div>
        <div class="kpi-sub">Vaman: {r['v_safe']:.1f} km/jam</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-title">Jari-Jari (Rc / Rmin)</div>
        <div class="kpi-value">{r['Rc']:.0f} <span style="font-size: 14px;">m</span></div>
        <div class="kpi-sub">Rmin: {r['Rmin']:.1f} m (SF {(r['Rc']/r['Rmin']):.2f})</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-title">Superelevasi Desain</div>
        <div class="kpi-value">{r['e']*100:.2f} <span style="font-size: 14px;">%</span></div>
        <div class="kpi-sub">emax: {r['emax']*100:.1f}% | en: {r['en']*100:.1f}%</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-title">Panjang Tikungan (Ltot)</div>
        <div class="kpi-value">{r['Ltot']:.2f} <span style="font-size: 14px;">m</span></div>
        <div class="kpi-sub">Ls: {r['Ls_sup']:.1f}m | Lc: {r['Lc']:.1f}m</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-title">Rentang Stasioning</div>
        <div class="kpi-value" style="font-size: 15px;">{sta_str(r['sta_awal'])}</div>
        <div class="kpi-sub">s/d {sta_str(r['sta_akhir'])}</div>
    </div>
</div>
"""
st.markdown(kpi_html, unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# 6. ENGINE RENDER GAMBAR CAD SKALATIS & WORD (.DOCX)
# ----------------------------------------------------------------------------
def draw_dim_line(ax, p1, p2, label, offset_vec=(0.0, 0.0), text_offset=(0.0, 0.0), fontsize=7.5):
    q1 = (p1[0] + offset_vec[0], p1[1] + offset_vec[1])
    q2 = (p2[0] + offset_vec[0], p2[1] + offset_vec[1])
    ax.plot([p1[0], q1[0]], [p1[1], q1[1]], color='#000000', linewidth=0.55, zorder=4)
    ax.plot([p2[0], q2[0]], [p2[1], q2[1]], color='#000000', linewidth=0.55, zorder=4)
    ax.annotate(
        "", xy=q2, xytext=q1,
        arrowprops=dict(arrowstyle='<->', color='#000000', lw=0.75, shrinkA=0, shrinkB=0, mutation_scale=8),
        zorder=5
    )
    mid_x = 0.5 * (q1[0] + q2[0]) + text_offset[0]
    mid_y = 0.5 * (q1[1] + q2[1]) + text_offset[1]
    angle_deg = math.degrees(math.atan2(q2[1] - q1[1], q2[0] - q1[0]))
    if angle_deg > 90:
        angle_deg -= 180
    elif angle_deg < -90:
        angle_deg += 180
    ax.text(
        mid_x, mid_y, label,
        fontsize=fontsize, color='#000000', ha='center', va='center',
        rotation=angle_deg,
        bbox=dict(boxstyle='square,pad=0.12', facecolor='white', edgecolor='none'),
        zorder=6
    )

def render_cad_denah_bw(r, B, arah, draw_w_mm=330.0, draw_h_mm=277.0):
    Rc = r["Rc"]
    Ls = r["Ls"]
    Lc = r["Lc"]
    Ts = r["T"]
    Es = r["E"]
    Xs = r["Xs"]
    Ys = r["Ys"]
    p_val = r["p"]
    k_val = r["k"]
    delta_deg = r["delta"]
    ths_deg = r["ths"]
    fc = r["jenis"].startswith("Full")
    is_ss = r["jenis"].startswith("Spiral-Spiral")

    mirror = 1.0 if arah == "Kanan" else -1.0
    d_rad = math.radians(delta_deg)
    half_d = d_rad / 2.0

    R_eff = Rc if fc else (Rc + p_val)
    y_O = 0.0
    x_O = 0.0
    x_PI = 0.0
    y_PI = R_eff / math.cos(half_d)

    ang_in = math.radians(90.0 - delta_deg / 2.0)
    u_in = np.array([mirror * math.cos(ang_in), math.sin(ang_in)])
    n_in = np.array([mirror * math.sin(ang_in), -math.cos(ang_in)])

    ang_out = math.radians(-90.0 + delta_deg / 2.0)
    u_out = np.array([mirror * math.cos(ang_out), math.sin(ang_out)])
    n_out = np.array([-mirror * math.sin(ang_out), math.cos(ang_out)])

    p_PI = np.array([x_PI, y_PI])
    p_TS = p_PI - Ts * u_in
    p_ST = p_PI + Ts * u_out
    p_O = np.array([x_O, y_O])

    ext_tan = max(0.22 * Ts, 15.0)
    p_tan_start = p_TS - ext_tan * u_in
    p_tan_end = p_ST + ext_tan * u_out
    p_SCS = np.array([0.0, Rc])

    if fc:
        p_SC = p_TS.copy()
        p_CS = p_ST.copy()
        ang_sc = math.atan2(p_SC[1] - p_O[1], p_SC[0] - p_O[0])
        ang_cs = math.atan2(p_CS[1] - p_O[1], p_CS[0] - p_O[0])
        angs_arc = np.linspace(ang_sc, ang_cs, 120)
        curve_x = p_O[0] + Rc * np.cos(angs_arc)
        curve_y = p_O[1] + Rc * np.sin(angs_arc)
        arc_ext_x, arc_ext_y = np.array([]), np.array([])
        p_proj_SC = p_SC.copy()
        p_pc_circle = p_SC.copy()
    else:
        s_sp = np.linspace(0.0, Ls, 60)
        x_loc = s_sp - (s_sp**5) / (40.0 * (Rc**2) * (Ls**2))
        y_loc = (s_sp**3) / (6.0 * Rc * Ls)
        sp1_pts = np.array([p_TS + xl_i * u_in + yl_i * n_in for xl_i, yl_i in zip(x_loc, y_loc)])
        sp2_pts = np.array([p_ST - xl_i * u_out + yl_i * n_out for xl_i, yl_i in zip(x_loc, y_loc)])[::-1]
        p_SC = sp1_pts[-1]
        p_CS = sp2_pts[0]

        ang_sc = math.atan2(p_SC[1] - p_O[1], p_SC[0] - p_O[0])
        ang_cs = math.atan2(p_CS[1] - p_O[1], p_CS[0] - p_O[0])
        angs_arc = np.linspace(ang_sc, ang_cs, 100)
        arc_x = p_O[0] + Rc * np.cos(angs_arc)
        arc_y = p_O[1] + Rc * np.sin(angs_arc)

        curve_x = np.concatenate([sp1_pts[:, 0], arc_x, sp2_pts[:, 0]])
        curve_y = np.concatenate([sp1_pts[:, 1], arc_y, sp2_pts[:, 1]])

        angs_ext = np.linspace(ang_sc + mirror * math.radians(ths_deg), ang_sc, 40)
        arc_ext_x = p_O[0] + Rc * np.cos(angs_ext)
        arc_ext_y = p_O[1] + Rc * np.sin(angs_ext)

        p_proj_SC = p_TS + Xs * u_in
        p_pc_circle = p_TS + k_val * u_in + p_val * n_in

    all_x = np.concatenate([curve_x, [p_tan_start[0], p_tan_end[0], p_PI[0], p_O[0]]])
    all_y = np.concatenate([curve_y, [p_tan_start[1], p_tan_end[1], p_PI[1], p_O[1]]])
    raw_dx = max(float(np.max(all_x) - np.min(all_x)), 10.0)
    raw_dy = max(float(np.max(all_y) - np.min(all_y)), 10.0)

    pad_l_mm, pad_r_mm, pad_t_mm, pad_b_mm = 16.0, 14.0, 16.0, 18.0
    plot_w_mm = draw_w_mm - pad_l_mm - pad_r_mm
    plot_h_mm = draw_h_mm - pad_t_mm - pad_b_mm

    need_dx_m = raw_dx * 1.48
    need_dy_m = raw_dy * 1.32

    req_scale_denom = max((need_dx_m * 1000.0) / plot_w_mm, (need_dy_m * 1000.0) / plot_h_mm)
    standard_scales = [100, 200, 250, 500, 750, 1000, 1250, 1500, 2000, 2500, 3000, 4000, 5000, 7500, 10000, 15000, 20000, 25000]
    scale_denom = next((sc for sc in standard_scales if sc >= req_scale_denom), int(math.ceil(req_scale_denom / 5000.0) * 5000))

    world_w_m = (plot_w_mm / 1000.0) * scale_denom
    world_h_m = (plot_h_mm / 1000.0) * scale_denom

    cx = 0.5 * (float(np.min(all_x)) + float(np.max(all_x))) - mirror * 0.06 * world_w_m
    cy = 0.5 * (float(np.min(all_y)) + float(np.max(all_y)))
    x_min_lim, x_max_lim = cx - world_w_m / 2.0, cx + world_w_m / 2.0
    y_min_lim, y_max_lim = cy - world_h_m / 2.0, cy + world_h_m / 2.0

    fig = plt.figure(figsize=((draw_w_mm - 2.0) / 25.4, (draw_h_mm - 2.0) / 25.4), dpi=300)
    ax = fig.add_axes([pad_l_mm / draw_w_mm, pad_b_mm / draw_h_mm, plot_w_mm / draw_w_mm, plot_h_mm / draw_h_mm])
    ax.set_xlim(x_min_lim, x_max_lim)
    ax.set_ylim(y_min_lim, y_max_lim)
    ax.set_aspect('equal', adjustable='box')
    ax.axis('off')

    ax.plot([p_tan_start[0], p_PI[0]], [p_tan_start[1], p_PI[1]], color='#000000', linewidth=0.9, zorder=3)
    ax.plot([p_PI[0], p_tan_end[0]], [p_PI[1], p_tan_end[1]], color='#000000', linewidth=0.9, zorder=3)
    ext_pi = 0.14 * Ts
    p_PI_ext = p_PI + ext_pi * u_in
    ax.plot([p_PI[0], p_PI_ext[0]], [p_PI[1], p_PI_ext[1]], color='#000000', linewidth=0.75, zorder=3)

    ax.plot(curve_x, curve_y, color='#000000', linewidth=1.6, zorder=4)
    if len(arc_ext_x) > 0:
        ax.plot(arc_ext_x, arc_ext_y, color='#000000', linewidth=0.75, linestyle='--', zorder=3)

    ax.plot([p_O[0], p_PI[0]], [p_O[1], p_PI[1]], color='#000000', linewidth=0.75, zorder=3)
    if not fc:
        ax.plot([p_O[0], p_SC[0]], [p_O[1], p_SC[1]], color='#000000', linewidth=0.75, zorder=3)
        ax.plot([p_O[0], p_CS[0]], [p_O[1], p_CS[1]], color='#000000', linewidth=0.75, zorder=3)
        ax.plot([p_O[0], p_TS[0] + k_val * u_in[0]], [p_O[1], p_TS[1] + k_val * u_in[1]], color='#000000', linewidth=0.65, zorder=3)
        ax.plot([p_SC[0], p_proj_SC[0]], [p_SC[1], p_proj_SC[1]], color='#000000', linewidth=0.7, zorder=3)
    else:
        ax.plot([p_O[0], p_TS[0]], [p_O[1], p_TS[1]], color='#000000', linewidth=0.75, zorder=3)
        ax.plot([p_O[0], p_ST[0]], [p_O[1], p_ST[1]], color='#000000', linewidth=0.75, zorder=3)

    nodes_x = [p_TS[0], p_ST[0], p_SCS[0]]
    nodes_y = [p_TS[1], p_ST[1], p_SCS[1]]
    if not fc:
        nodes_x.extend([p_SC[0], p_CS[0], p_pc_circle[0]])
        nodes_y.extend([p_SC[1], p_CS[1], p_pc_circle[1]])
    ax.scatter(nodes_x, nodes_y, facecolor='white', edgecolor='#000000', s=20, linewidths=0.9, zorder=7)

    lbl_a, lbl_b = ("TC", "CT") if fc else ("TS", "ST")
    ax.text(p_PI[0], p_PI[1] + 0.025 * world_h_m, f"PI (STA {sta_str(r['sta_pi'])})", fontsize=8, fontweight='bold', ha='center', va='bottom', color='#000000')
    ax.text(p_TS[0] - mirror * 0.025 * world_w_m, p_TS[1] - 0.015 * world_h_m, f"{lbl_a}\n({sta_str(r['sta_awal'])})", fontsize=7.5, fontweight='bold', ha='center', va='top', color='#000000')
    ax.text(p_ST[0] + mirror * 0.025 * world_w_m, p_ST[1] - 0.015 * world_h_m, f"{lbl_b}\n({sta_str(r['sta_akhir'])})", fontsize=7.5, fontweight='bold', ha='center', va='top', color='#000000')

    mid_es = 0.5 * (p_PI + p_SCS)
    ax.text(mid_es[0] + mirror * 0.015 * world_w_m, mid_es[1], f"Es = {Es:.2f} m", fontsize=7, ha='left' if mirror > 0 else 'right', va='center', color='#000000')

    mid_r = 0.5 * (p_O + p_SCS)
    ax.text(mid_r[0] + mirror * 0.015 * world_w_m, mid_r[1], f"R = {Rc:.2f} m\n(Rmin = {r['Rmin']:.2f} m)", fontsize=7, fontweight='bold', ha='left' if mirror > 0 else 'right', va='center', color='#000000')

    out_dir = -n_in
    step_dim = 0.055 * world_w_m
    draw_dim_line(ax, p_TS, p_PI, f"Ts = {Ts:.2f} m", offset_vec=tuple(out_dir * (3.1 * step_dim)), text_offset=tuple(out_dir * (0.015 * world_w_m)), fontsize=7)

    if not fc:
        draw_dim_line(ax, p_TS, p_proj_SC, f"Xc = {Xs:.2f} m", offset_vec=tuple(out_dir * (2.0 * step_dim)), text_offset=tuple(out_dir * (0.014 * world_w_m)), fontsize=6.8)
        p_k_end = p_TS + k_val * u_in
        draw_dim_line(ax, p_TS, p_k_end, f"k = {k_val:.2f} m", offset_vec=tuple(out_dir * (0.95 * step_dim)), text_offset=tuple(out_dir * (0.014 * world_w_m)), fontsize=6.5)

    bar_phys_mm = 40.0
    bar_real_m = (bar_phys_mm / 1000.0) * scale_denom
    sb_x0 = x_max_lim - bar_real_m - 0.03 * world_w_m
    sb_y0 = y_min_lim + 0.035 * world_h_m
    sb_h = 0.011 * world_h_m
    for seg_i in range(4):
        seg_w = bar_real_m / 4.0
        fc_seg = '#000000' if seg_i % 2 == 0 else '#FFFFFF'
        rect = patches.Rectangle((sb_x0 + seg_i * seg_w, sb_y0), seg_w, sb_h, linewidth=0.7, edgecolor='#000000', facecolor=fc_seg, zorder=8)
        ax.add_patch(rect)
    ax.text(sb_x0, sb_y0 - 0.01 * world_h_m, "0", fontsize=5.8, ha='center', va='top', color='#000000')
    ax.text(sb_x0 + bar_real_m, sb_y0 - 0.01 * world_h_m, f"{bar_real_m:g} m", fontsize=5.8, ha='center', va='top', color='#000000')
    ax.text(sb_x0 + bar_real_m / 2.0, sb_y0 + sb_h + 0.007 * world_h_m, f"SKALA GRAFIS 1 : {scale_denom:,}", fontsize=6, fontweight='bold', ha='center', va='bottom', color='#000000')

    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=300)
    plt.close(fig)
    buf.seek(0)
    return buf, scale_denom

def render_cad_superelevasi_bw(r, draw_w_mm=330.0, draw_h_mm=277.0):
    e_des = r["e"] * 100.0
    en_val = r["en"] * 100.0
    ls_val = r["Ls_sup"]
    lc_val = r["Lc"]
    ltot_val = r["Ltot"]
    fc = r["jenis"].startswith("Full")

    if fc:
        s_start = -(2.0 / 3.0) * ls_val
        s_tc = 0.0
        s_full1 = (1.0 / 3.0) * ls_val
        s_full2 = ltot_val - (1.0 / 3.0) * ls_val
        s_ct = ltot_val
        s_end = ltot_val + (2.0 / 3.0) * ls_val
    else:
        s_start = 0.0
        s_tc = 0.0
        s_full1 = ls_val
        s_full2 = ltot_val - ls_val
        s_ct = ltot_val
        s_end = ltot_val

    total_span_m = max(s_end - s_start, 20.0)
    pad_l_mm, pad_r_mm, pad_t_mm, pad_b_mm = 22.0, 22.0, 22.0, 22.0
    plot_w_mm = draw_w_mm - pad_l_mm - pad_r_mm
    plot_h_mm = draw_h_mm - pad_t_mm - pad_b_mm

    need_dx_m = total_span_m * 1.25
    req_scale_h = (need_dx_m * 1000.0) / plot_w_mm
    standard_scales = [100, 200, 250, 500, 750, 1000, 1250, 1500, 2000, 2500, 3000, 4000, 5000]
    scale_h = next((sc for sc in standard_scales if sc >= req_scale_h), int(math.ceil(req_scale_h / 1000.0) * 1000))

    world_w_m = (plot_w_mm / 1000.0) * scale_h
    cx = 0.5 * (s_start + s_end)
    x_min_lim, x_max_lim = cx - world_w_m / 2.0, cx + world_w_m / 2.0
    y_lim = max(e_des * 1.75, 14.0)

    fig = plt.figure(figsize=((draw_w_mm - 2.0) / 25.4, (draw_h_mm - 2.0) / 25.4), dpi=300)
    ax = fig.add_axes([pad_l_mm / draw_w_mm, pad_b_mm / draw_h_mm, plot_w_mm / draw_w_mm, plot_h_mm / draw_h_mm])
    ax.set_xlim(x_min_lim, x_max_lim)
    ax.set_ylim(-y_lim, y_lim)
    ax.axis('off')

    ax.plot([s_start - 0.04 * world_w_m, s_end + 0.04 * world_w_m], [0.0, 0.0], color='#000000', linestyle='-.', linewidth=0.85)
    ax.plot([s_start, s_end], [-en_val, -en_val], color='#000000', linestyle='--', linewidth=0.75)

    trans_len = max(s_full1 - s_start, 1e-6)
    s_zero1 = s_start + trans_len * (en_val / (e_des + en_val))
    s_zero2 = s_end - trans_len * (en_val / (e_des + en_val))
    s_crown1 = s_start + trans_len * ((2.0 * en_val) / (e_des + en_val))
    s_crown2 = s_end - trans_len * ((2.0 * en_val) / (e_des + en_val))

    x_luar = [s_start, s_zero1, s_crown1, s_full1, s_full2, s_crown2, s_zero2, s_end]
    y_luar = [-en_val, 0.0, en_val, e_des, e_des, en_val, 0.0, -en_val]
    ax.plot(x_luar, y_luar, color='#000000', linewidth=1.3)

    x_dalam = [s_start, s_crown1, s_full1, s_full2, s_crown2, s_end]
    y_dalam = [-en_val, -en_val, -e_des, -e_des, -en_val, -en_val]
    ax.plot(x_dalam, y_dalam, color='#000000', linewidth=1.3)

    mid_lc = 0.5 * (s_full1 + s_full2)
    ax.text(mid_lc, e_des + 0.035 * y_lim, f"e = {e_des:.2f} %", fontsize=7.5, fontweight='bold', ha='center', va='bottom', color='#000000')
    ax.text(mid_lc, -e_des - 0.035 * y_lim, f"-e = -{e_des:.2f} %", fontsize=7.5, fontweight='bold', ha='center', va='top', color='#000000')

    y_dim_b2 = -y_lim * 0.72
    draw_dim_line(ax, (s_start, y_dim_b2), (s_full1, y_dim_b2), f"LS = {ls_val:.2f} m", text_offset=(0, 0.035 * y_lim), fontsize=7)
    if lc_val > 0:
        draw_dim_line(ax, (s_full1, y_dim_b2), (s_full2, y_dim_b2), f"LC = {lc_val:.2f} m", text_offset=(0, 0.035 * y_lim), fontsize=7)
    draw_dim_line(ax, (s_full2, y_dim_b2), (s_end, y_dim_b2), f"LS = {ls_val:.2f} m", text_offset=(0, 0.035 * y_lim), fontsize=7)

    bar_phys_mm = 40.0
    bar_real_m = (bar_phys_mm / 1000.0) * scale_h
    sb_x0 = x_max_lim - bar_real_m - 0.02 * world_w_m
    sb_y0 = -y_lim * 0.92
    sb_h = 0.022 * y_lim
    for seg_i in range(4):
        seg_w = bar_real_m / 4.0
        fc_seg = '#000000' if seg_i % 2 == 0 else '#FFFFFF'
        rect = patches.Rectangle((sb_x0 + seg_i * seg_w, sb_y0), seg_w, sb_h, linewidth=0.7, edgecolor='#000000', facecolor=fc_seg, zorder=8)
        ax.add_patch(rect)
    ax.text(sb_x0, sb_y0 - 0.02 * y_lim, "0", fontsize=5.8, ha='center', va='top', color='#000000')
    ax.text(sb_x0 + bar_real_m, sb_y0 - 0.02 * y_lim, f"{bar_real_m:g} m", fontsize=5.8, ha='center', va='top', color='#000000')
    ax.text(sb_x0 + bar_real_m / 2.0, sb_y0 + sb_h + 0.015 * y_lim, f"SKALA HORIZONTAL 1 : {scale_h:,}", fontsize=6, fontweight='bold', ha='center', va='bottom', color='#000000')

    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=300)
    plt.close(fig)
    buf.seek(0)
    return buf, scale_h

def build_etiket_table_word(
    parent_cell, r, B, arah,
    logo_text, catatan_text, kegiatan_text, pekerjaan_text, lokasi_text,
    diperiksa_text, digambar_text, judul_gambar, skala_str, tanggal_str, no_lembar, jml_lembar
):
    row_heights_mm = [37.5, 7.5, 50.0, 7.5, 15.0, 7.5, 15.0, 7.5, 15.0, 7.5, 27.5, 7.5, 27.5, 7.5, 17.5, 7.5, 12.0]
    et_tbl = parent_cell.add_table(rows=len(row_heights_mm), cols=6)
    et_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    et_tbl.autofit = False

    for r_idx, h_mm in enumerate(row_heights_mm):
        row = et_tbl.rows[r_idx]
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:trHeight {nsdecls("w")} w:val="{int(h_mm * 56.7)}" w:hRule="exact"/>'))
        for c_idx in range(6):
            cell = row.cells[c_idx]
            cell.width = Mm(10.0)
            tcPr = cell._tc.get_or_add_tcPr()
            tcPr.append(parse_xml(
                f'<w:tcBorders {nsdecls("w")}>'
                f'<w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
                f'<w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
                f'<w:left w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
                f'<w:right w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
                f'</w:tcBorders>'
            ))
            tcPr.append(parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="20" w:type="dxa"/><w:bottom w:w="20" w:type="dxa"/><w:left w:w="45" w:type="dxa"/><w:right w:w="45" w:type="dxa"/></w:tcMar>'))

    def put_txt(cell, txt, size_pt=7.0, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER):
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = align
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(txt)
        run.font.name = "Arial"
        run.font.size = Pt(size_pt)
        run.font.bold = bold
        run.font.color.rgb = RGBColor(0, 0, 0)

    c_logo = et_tbl.cell(0, 0).merge(et_tbl.cell(0, 5))
    put_txt(c_logo, logo_text, size_pt=10, bold=True)
    c_cat_h = et_tbl.cell(1, 0).merge(et_tbl.cell(1, 5))
    put_txt(c_cat_h, "CATATAN", size_pt=6.5, bold=True)
    c_cat_b = et_tbl.cell(2, 0).merge(et_tbl.cell(2, 5))
    put_txt(c_cat_b, catatan_text, size_pt=6.2, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT)
    c_keg_h = et_tbl.cell(3, 0).merge(et_tbl.cell(3, 5))
    put_txt(c_keg_h, "KEGIATAN", size_pt=6.5, bold=True)
    c_keg_b = et_tbl.cell(4, 0).merge(et_tbl.cell(4, 5))
    put_txt(c_keg_b, kegiatan_text, size_pt=6.8, bold=True)
    c_pek_h = et_tbl.cell(5, 0).merge(et_tbl.cell(5, 5))
    put_txt(c_pek_h, "PEKERJAAN", size_pt=6.5, bold=True)
    c_pek_b = et_tbl.cell(6, 0).merge(et_tbl.cell(6, 5))
    put_txt(c_pek_b, pekerjaan_text, size_pt=6.8, bold=True)
    c_lok_h = et_tbl.cell(7, 0).merge(et_tbl.cell(7, 5))
    put_txt(c_lok_h, "LOKASI", size_pt=6.5, bold=True)
    c_lok_b = et_tbl.cell(8, 0).merge(et_tbl.cell(8, 5))
    put_txt(c_lok_b, lokasi_text, size_pt=6.8, bold=False)
    c_per_h = et_tbl.cell(9, 0).merge(et_tbl.cell(9, 5))
    put_txt(c_per_h, "DI PERIKSA OLEH :", size_pt=6.5, bold=True)
    c_per_b = et_tbl.cell(10, 0).merge(et_tbl.cell(10, 5))
    put_txt(c_per_b, diperiksa_text, size_pt=7.0, bold=False)
    c_dig_h = et_tbl.cell(11, 0).merge(et_tbl.cell(11, 5))
    put_txt(c_dig_h, "DI GAMBAR OLEH :", size_pt=6.5, bold=True)
    c_dig_b = et_tbl.cell(12, 0).merge(et_tbl.cell(12, 5))
    put_txt(c_dig_b, digambar_text, size_pt=7.0, bold=False)
    c_jd_h = et_tbl.cell(13, 0).merge(et_tbl.cell(13, 2))
    c_sk_h = et_tbl.cell(13, 3).merge(et_tbl.cell(13, 5))
    put_txt(c_jd_h, "JUDUL GAMBAR", size_pt=6.2, bold=True)
    put_txt(c_sk_h, "SKALA", size_pt=6.2, bold=True)
    c_jd_b = et_tbl.cell(14, 0).merge(et_tbl.cell(14, 2))
    c_sk_b = et_tbl.cell(14, 3).merge(et_tbl.cell(14, 5))
    put_txt(c_jd_b, judul_gambar, size_pt=6.5, bold=True)
    put_txt(c_sk_b, skala_str, size_pt=7.5, bold=True)
    c_tg_h = et_tbl.cell(15, 0).merge(et_tbl.cell(15, 1))
    c_no_h = et_tbl.cell(15, 2).merge(et_tbl.cell(15, 3))
    c_jm_h = et_tbl.cell(15, 4).merge(et_tbl.cell(15, 5))
    put_txt(c_tg_h, "TANGGAL", size_pt=5.8, bold=True)
    put_txt(c_no_h, "NO. LEMBAR", size_pt=5.8, bold=True)
    put_txt(c_jm_h, "JML. LEMBAR", size_pt=5.8, bold=True)
    c_tg_b = et_tbl.cell(16, 0).merge(et_tbl.cell(16, 1))
    c_no_b = et_tbl.cell(16, 2).merge(et_tbl.cell(16, 3))
    c_jm_b = et_tbl.cell(16, 4).merge(et_tbl.cell(16, 5))
    put_txt(c_tg_b, tanggal_str, size_pt=6.2, bold=False)
    put_txt(c_no_b, no_lembar, size_pt=7.0, bold=True)
    put_txt(c_jm_b, jml_lembar, size_pt=7.0, bold=True)

def generate_word_2_sheets_landscape(r, B, arah, meta):
    page_w_mm, page_h_mm = 420.0, 297.0
    margin_left_mm, margin_other_mm = 20.0, 10.0
    inner_w_mm = page_w_mm - margin_left_mm - margin_other_mm
    inner_h_mm = page_h_mm - 2.0 * margin_other_mm
    etiket_w_mm = 60.0
    draw_w_mm = inner_w_mm - etiket_w_mm

    img1_buf, scale1 = render_cad_denah_bw(r, B, arah, draw_w_mm=draw_w_mm, draw_h_mm=inner_h_mm)
    img2_buf, scale2 = render_cad_superelevasi_bw(r, draw_w_mm=draw_w_mm, draw_h_mm=inner_h_mm)

    doc = Document()
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Mm(page_w_mm)
    section.page_height = Mm(page_h_mm)
    section.left_margin = Mm(margin_left_mm)
    section.right_margin = Mm(margin_other_mm)
    section.top_margin = Mm(margin_other_mm)
    section.bottom_margin = Mm(margin_other_mm)

    catatan_default = (
        f"1. Standar: TPGJAK 1997 / BM 2021\n"
        f"2. Jenis: {r['jenis']}\n"
        f"3. Vr = {r['V']:.0f} km/jam, Δ = {r['delta']:.2f}°\n"
        f"4. Rmin = {r['Rmin']:.2f} m, Rc = {r['Rc']:.2f} m\n"
        f"5. e = {r['e']*100:.2f}%, en = {r['en']*100:.1f}%\n"
        f"6. Ls = {r['Ls_sup']:.2f} m, Ltot = {r['Ltot']:.2f} m\n"
        f"7. Satuan ukuran dalam Meter"
    )
    catatan_final = meta["catatan"].strip() if meta["catatan"].strip() else catatan_default

    sheets_spec = [
        (img1_buf, f"DENAH TIKUNGAN\n({r['jenis'].upper()})", f"1 : {scale1:,}", "01", "02"),
        (img2_buf, "DIAGRAM\nSUPERELEVASI", f"1 : {scale2:,}", "02", "02"),
    ]

    for idx_s, (img_stream, jdl_gbr, skl_txt, no_lbr, jml_lbr) in enumerate(sheets_spec):
        if idx_s > 0:
            doc.add_page_break()

        outer_tbl = doc.add_table(rows=1, cols=2)
        outer_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        outer_tbl.autofit = False

        row_o = outer_tbl.rows[0]
        trPr = row_o._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:trHeight {nsdecls("w")} w:val="{int(inner_h_mm * 56.7)}" w:hRule="exact"/>'))

        c_draw = row_o.cells[0]
        c_etiket = row_o.cells[1]
        c_draw.width = Mm(draw_w_mm)
        c_etiket.width = Mm(etiket_w_mm)

        for cell_box, b_left, b_right in [(c_draw, "14", "8"), (c_etiket, "8", "14")]:
            tcPr = cell_box._tc.get_or_add_tcPr()
            tcPr.append(parse_xml(
                f'<w:tcBorders {nsdecls("w")}>'
                f'<w:top w:val="single" w:sz="14" w:space="0" w:color="000000"/>'
                f'<w:bottom w:val="single" w:sz="14" w:space="0" w:color="000000"/>'
                f'<w:left w:val="single" w:sz="{b_left}" w:space="0" w:color="000000"/>'
                f'<w:right w:val="single" w:sz="{b_right}" w:space="0" w:color="000000"/>'
                f'</w:tcBorders>'
            ))

        c_draw.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p_img = c_draw.paragraphs[0]
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.add_run().add_picture(img_stream, width=Mm(draw_w_mm - 2.0))

        build_etiket_table_word(
            c_etiket, r, B, arah,
            logo_text=meta["logo"],
            catatan_text=catatan_final,
            kegiatan_text=meta["kegiatan"],
            pekerjaan_text=meta["pekerjaan"],
            lokasi_text=meta["lokasi"],
            diperiksa_text=meta["diperiksa"],
            digambar_text=meta["digambar"],
            judul_gambar=jdl_gbr,
            skala_str=skl_txt,
            tanggal_str=meta["tanggal"],
            no_lembar=no_lbr,
            jml_lembar=jml_lbr
        )

    out_buf = io.BytesIO()
    doc.save(out_buf)
    return out_buf.getvalue(), scale1, scale2

def generate_excel_bm2021(r, pvi_data):
    wb = openpyxl.Workbook()
    font_title = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    font_sec = Font(name="Calibri", size=11, bold=True, color="880E4F")
    font_reg = Font(name="Calibri", size=10, color="2D3142")
    font_bold = Font(name="Calibri", size=10, bold=True, color="2D3142")
    fill_header = PatternFill(start_color="AD1457", end_color="AD1457", fill_type="solid")
    fill_sec = PatternFill(start_color="FCE4EC", end_color="FCE4EC", fill_type="solid")
    thin = Side(style="thin", color="F48FB1")
    border_cell = Border(left=thin, right=thin, top=thin, bottom=thin)
    align_c = Alignment(horizontal="center", vertical="center")
    align_l = Alignment(horizontal="left", vertical="center")

    ws = wb.active
    ws.title = "Perhitungan Tikungan"
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 32
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 14

    ws.merge_cells("A1:E1")
    ws["A1"] = f"ANALISIS GEOMETRIK TIKUNGAN ({r['jenis'].upper()}) - STANDAR TPGJAK 1997"
    ws["A1"].font = font_title
    ws["A1"].fill = fill_header
    ws["A1"].alignment = align_c
    ws.row_dimensions[1].height = 30

    headers = ["No", "Parameter Desain", "Rumus / Rujukan Standar", "Nilai Calculated", "Satuan"]
    for col_idx, h in enumerate(headers, 1):
        c = ws.cell(row=2, column=col_idx, value=h)
        c.font = font_bold
        c.fill = fill_sec
        c.alignment = align_c
        c.border = border_cell
    ws.row_dimensions[2].height = 24

    data_rows = [
        (1, "Kecepatan Rencana (Vr)", "Kriteria Desain Klasifikasi Jalan", r["V"], "km/jam"),
        (2, "Sudut Tikungan (Δ)", "Pengukuran Trase Peta Kontur", r["delta"], "derajat"),
        (3, "Superelevasi Maksimum (e_max)", "Standar Wilayah Perencanaan", r["emax"] * 100, "%"),
        (4, "Kemiringan Normal (en)", "Kebutuhan Drainase Permukaan", r["en"] * 100, "%"),
        (5, "Koefisien Gesek Maks (f_max)", "-0.00065*Vr + 0.192", round(r["fm"], 4), "-"),
        (6, "Jari-Jari Minimum (Rmin)", "Vr² / (127 * (emax + f_max))", round(r["Rmin"], 2), "m"),
        (7, "Jari-Jari Rencana (Rc)", "Rc Rencana ≥ Rmin", round(r["Rc"], 2), "m"),
        (8, "Derajat Lengkung (D)", "1432.39 / Rc", round(r["D"], 3), "derajat"),
        (9, "Superelevasi Desain (e)", "-emax*(D/Dmax)² + 2*emax*(D/Dmax)", round(r["e"] * 100, 2), "%"),
        (10, "Panjang Spiral Kritis (Ls pakai)", "MAX(Ls1, Ls2, Ls3) dibulatkan 5m", round(r["Ls_sup"], 2), "m"),
        (11, "Sudut Spiral (θs)", "(Ls * 90) / (π * Rc)", round(r["ths"], 3), "derajat"),
        (12, "Panjang Busur Lingkaran (Lc)", "((Δ - 2θs)/360) * 2π * Rc", round(r["Lc"], 2), "m"),
        (13, "Panjang Tangen Tikungan (Ts)", "(Rc + p) * tan(Δ/2) + k", round(r["T"], 2), "m"),
        (14, "Jarak Eksternal (Es)", "(Rc + p) / cos(Δ/2) - Rc", round(r["E"], 2), "m"),
        (15, "Panjang Total Lengkung (Ltot)", "Lc + 2 * Ls", round(r["Ltot"], 2), "m"),
        (16, "Kecepatan Aman Tikungan (V_safe)", "√(127 * Rc * (e + f_max))", round(r["v_safe"], 1), "km/jam"),
        (17, "STA Awal Tikungan", "STA PI - Ts", sta_str(r["sta_awal"]), "m"),
        (18, "STA Akhir Tikungan", "STA Awal + Ltot", sta_str(r["sta_akhir"]), "m"),
    ]

    for idx, (no, param, formula, val, unit) in enumerate(data_rows, 3):
        ws.row_dimensions[idx].height = 20
        ws.cell(row=idx, column=1, value=no).alignment = align_c
        ws.cell(row=idx, column=2, value=param).alignment = align_l
        ws.cell(row=idx, column=3, value=formula).alignment = align_l
        c_val = ws.cell(row=idx, column=4, value=val)
        c_val.alignment = align_c
        c_val.font = font_bold
        ws.cell(row=idx, column=5, value=unit).alignment = align_c
        for col_idx in range(1, 6):
            ws.cell(row=idx, column=col_idx).border = border_cell

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ----------------------------------------------------------------------------
# 7. TAB UTAMA APLIKASI
# ----------------------------------------------------------------------------
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "📍 Denah Tikungan (2D)", 
    "📈 Diagram Superelevasi", 
    "📐 Potongan Melintang Interaktif",
    "⛰️ Profil Vertikal & Galian/Timbunan",
    "🚘 Simulasi 3D Mobil Jalan", 
    "📐 Gambar Kerja Word (Hitam Putih)",
    "📋 Tabel Hitungan & Excel"
])

# ----------------------------------------------------------------------------
# TAB 1: DENAH TIKUNGAN 2D INTERAKTIF
# ----------------------------------------------------------------------------
with tab1:
    st.markdown("<h3 style='color: #3A0056;'>📍 Visualisasi Denah Trase Tikungan 2D</h3>", unsafe_allow_html=True)
    
    col_t1, col_t2 = st.columns([3, 1])
    with col_t2:
        st.markdown("""
        <div class="tulip-card">
            <h4 style="color: #3A0056; margin: 0 0 10px 0; font-size: 15px;">⚙️ Opsi Tampilan</h4>
        """, unsafe_allow_html=True)
        show_pavement = st.checkbox("Badan Perkerasan Aspal", True)
        show_shoulder = st.checkbox("Bahu Jalan", True)
        show_tangents = st.checkbox("Garis Tangen & Titik PI", True)
        show_radius = st.checkbox("Jari-jari Rc & Pusat (O)", True)
        show_stations = st.checkbox("Stasioning Ticks Tiap 50m", True)
        st.markdown("</div>", unsafe_allow_html=True)

    with col_t1:
        fig_2d = go.Figure()
        xl, yl = x_horiz + (B / 2) * nx, y_horiz + (B / 2) * ny
        xr, yr = x_horiz - (B / 2) * nx, y_horiz - (B / 2) * ny
        x_sh_l, y_sh_l = x_horiz + (B / 2 + Bahu) * nx, y_horiz + (B / 2 + Bahu) * ny
        x_sh_r, y_sh_r = x_horiz - (B / 2 + Bahu) * nx, y_horiz - (B / 2 + Bahu) * ny
        
        if show_shoulder:
            fig_2d.add_trace(go.Scatter(
                x=np.concatenate([x_sh_l, x_sh_r[::-1]]),
                y=np.concatenate([y_sh_l, y_sh_r[::-1]]),
                fill='toself', fillcolor='rgba(244, 143, 177, 0.18)',
                line=dict(color='rgba(194, 24, 91, 0.35)', dash='dot', width=1),
                name='Bahu Jalan', hoverinfo='skip'
            ))

        if show_pavement:
            fig_2d.add_trace(go.Scatter(
                x=np.concatenate([xl, xr[::-1]]),
                y=np.concatenate([yl, yr[::-1]]),
                fill='toself', fillcolor='rgba(200, 200, 210, 0.28)',
                line=dict(color='#AD1457', width=1.5), name='Badan Perkerasan', hoverinfo='skip'
            ))

        fig_2d.add_trace(go.Scatter(
            x=x_horiz, y=y_horiz, mode='lines', name='As Jalan (Centerline)',
            line=dict(color='#3A0056', width=3),
            hovertemplate='<b>STA:</b> %{customdata}<br><b>X:</b> %{x:.2f} m<br><b>Y:</b> %{y:.2f} m<extra></extra>',
            customdata=[sta_str(r['sta_awal'] + sv) for sv in s_vals]
        ))

        idx_start = np.argmin(np.abs(s_vals - 0))
        idx_end = np.argmin(np.abs(s_vals - r["Ltot"]))
        x_ts, y_ts = x_horiz[idx_start], y_horiz[idx_start]
        x_st, y_st = x_horiz[idx_end], y_horiz[idx_end]
        x_pi = x_ts + r["T"] * np.cos(ths_[idx_start])
        y_pi = y_ts + r["T"] * np.sin(ths_[idx_start])

        if show_tangents:
            fig_2d.add_trace(go.Scatter(
                x=[x_ts, x_pi, x_st], y=[y_ts, y_pi, y_st],
                mode='lines', name='Garis Singgung Tangen', line=dict(color='#1565C0', width=1.8, dash='dash')
            ))
            fig_2d.add_trace(go.Scatter(
                x=[x_pi], y=[y_pi], mode='markers', name='Titik PI',
                marker=dict(size=10, color='#111111', symbol='square')
            ))
            fig_2d.add_annotation(
                x=x_pi, y=y_pi,
                text=f"<b>PI</b><br>Δ = {r['delta']:.1f}°<br>STA {sta_str(r['sta_pi'])}",
                showarrow=False, xshift=12, yshift=32 if arah == "Kanan" else -32, xanchor="left",
                font=dict(color="#111111", size=11), bgcolor="rgba(255,255,255,0.92)",
                bordercolor="#F48FB1", borderwidth=1, borderpad=4
            )

        mid_idx = idx_start if r["jenis"].startswith("Full") else np.argmin(np.abs(s_vals - (r["Ltot"] / 2)))
        x_center = x_horiz[mid_idx] + sg * r["Rc"] * nx[mid_idx]
        y_center = y_horiz[mid_idx] + sg * r["Rc"] * ny[mid_idx]

        if show_radius:
            fig_2d.add_trace(go.Scatter(
                x=[x_horiz[mid_idx], x_center], y=[y_horiz[mid_idx], y_center],
                mode='lines', name=f'Jari-jari Rc = {r["Rc"]:.0f} m (Rmin = {r["Rmin"]:.1f} m)',
                line=dict(color='#2E7D32', width=1.5, dash='dot')
            ))
            fig_2d.add_trace(go.Scatter(
                x=[x_center], y=[y_center], mode='markers', name='Pusat Lengkungan (O)',
                marker=dict(size=12, color='#2E7D32', symbol='cross-thin', line=dict(width=2, color='#2E7D32'))
            ))

        if show_stations:
            sta_min_plot = math.ceil((r["sta_awal"] + s_vals[0]) / 50.0) * 50.0
            sta_max_plot = math.floor((r["sta_awal"] + s_vals[-1]) / 50.0) * 50.0
            tick_stas = np.arange(sta_min_plot, sta_max_plot + 1, 50.0)
            tick_half_len = max(B * 0.85, 5.5)
            for st_val in tick_stas:
                s_rel = st_val - r["sta_awal"]
                if s_vals[0] <= s_rel <= s_vals[-1]:
                    idx_t = np.argmin(np.abs(s_vals - s_rel))
                    xt0, yt0 = x_horiz[idx_t], y_horiz[idx_t]
                    nxt, nyt = nx[idx_t], ny[idx_t]
                    fig_2d.add_trace(go.Scatter(
                        x=[xt0 - tick_half_len * nxt, xt0 + tick_half_len * nxt],
                        y=[yt0 - tick_half_len * nyt, yt0 + tick_half_len * nyt],
                        mode='lines', showlegend=False, hoverinfo='skip', line=dict(color='#757575', width=1.2)
                    ))
                    lbl_side = -sg
                    fig_2d.add_annotation(
                        x=xt0 + lbl_side * (tick_half_len + 2.5) * nxt,
                        y=yt0 + lbl_side * (tick_half_len + 2.5) * nyt,
                        text=sta_str(st_val, d=0), showarrow=False,
                        font=dict(color="#424242", size=9.5), bgcolor="rgba(255,255,255,0.75)", borderpad=1
                    )

        fc = r["jenis"].startswith("Full")
        is_ss = r["jenis"].startswith("Spiral-Spiral")
        pts_x, pts_y = [x_ts, x_st], [y_ts, y_st]
        if fc:
            pts_lbls = [f"<b>TC</b><br>{sta_str(r['sta_awal'])}", f"<b>CT</b><br>{sta_str(r['sta_akhir'])}"]
        else:
            pts_lbls = [f"<b>TS</b><br>{sta_str(r['sta_awal'])}", f"<b>ST</b><br>{sta_str(r['sta_akhir'])}"]
            if is_ss:
                idx_mid = np.argmin(np.abs(s_vals - r["Ls"]))
                pts_x.append(x_horiz[idx_mid])
                pts_y.append(y_horiz[idx_mid])
                pts_lbls.append(f"<b>SC=CS</b><br>{sta_str(r['sta_awal'] + r['Ls'])}")
            else:
                idx_sc = np.argmin(np.abs(s_vals - r["Ls"]))
                idx_cs = np.argmin(np.abs(s_vals - (r["Ls"] + r["Lc"])))
                pts_x.extend([x_horiz[idx_sc], x_horiz[idx_cs]])
                pts_y.extend([y_horiz[idx_sc], y_horiz[idx_cs]])
                pts_lbls.extend([
                    f"<b>SC</b><br>{sta_str(r['sta_awal'] + r['Ls'])}",
                    f"<b>CS</b><br>{sta_str(r['sta_awal'] + r['Ls'] + r['Lc'])}"
                ])

        fig_2d.add_trace(go.Scatter(
            x=pts_x, y=pts_y, mode='markers', name='Titik Utama (TS/SC/CS/ST)',
            marker=dict(size=10, color='#D81B60', symbol='circle', line=dict(width=1.5, color='#FFFFFF'))
        ))
        for px, py, plbl in zip(pts_x, pts_y, pts_lbls):
            fig_2d.add_annotation(
                x=px, y=py, text=plbl, showarrow=False,
                yshift=-26 if arah == "Kanan" else 26,
                font=dict(color="#880E4F", size=10.5), bgcolor="rgba(255,255,255,0.92)",
                bordercolor="#F48FB1", borderwidth=1, borderpad=3
            )

        arah_txt = "tikungan ke kiri" if arah == "Kiri" else "tikungan ke kanan"
        fig_2d.update_layout(
            title=dict(text=f"Denah Trase Tikungan - {r['jenis']} ({arah_txt})", font=dict(color='#3A0056', size=16), x=0.5, xanchor='center'),
            xaxis=dict(title="X (m)", gridcolor='#E0E0E0', showline=True, linecolor='#333333', mirror=True),
            yaxis=dict(title="Y (m)", gridcolor='#E0E0E0', showline=True, linecolor='#333333', mirror=True, scaleanchor="x", scaleratio=1),
            legend=dict(bgcolor="rgba(255, 255, 255, 0.95)", bordercolor="#F48FB1", borderwidth=1, font=dict(color="#3A0056", size=11), x=0.99, y=0.99, xanchor="right", yanchor="top"),
            plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF", margin=dict(l=60, r=40, t=60, b=50), height=600
        )
        st.plotly_chart(fig_2d, use_container_width=True)

# ----------------------------------------------------------------------------
# TAB 2: DIAGRAM SUPERELEVASI INTERAKTIF
# ----------------------------------------------------------------------------
with tab2:
    st.markdown("<h3 style='color: #3A0056;'>📈 Diagram Superelevasi (Rotasi pada Sumbu Jalan)</h3>", unsafe_allow_html=True)
    
    e_desain = r["e"] * 100
    en_val = r["en"] * 100
    ls_val = r["Ls_sup"]
    ltot_val = r["Ltot"]
    sta_0 = r["sta_awal"]
    
    if r["jenis"].startswith("Full"):
        s_start_trans = -(2.0 / 3.0) * ls_val
        s_full_start = (1.0 / 3.0) * ls_val
        s_full_end = ltot_val - (1.0 / 3.0) * ls_val
        s_end_trans = ltot_val + (2.0 / 3.0) * ls_val
        krit_s = [s_start_trans, 0.0, s_full_start, s_full_end, ltot_val, s_end_trans]
        krit_lbl = ["Awal Transisi", "TC", "e penuh", "e penuh", "CT", "Akhir Transisi"]
    else:
        s_start_trans = 0.0
        s_full_start = ls_val
        s_full_end = ltot_val - ls_val
        s_end_trans = ltot_val
        if r["jenis"].startswith("Spiral-Spiral"):
            krit_s = [0.0, ls_val, ltot_val]
            krit_lbl = ["TS", "SC=CS", "ST"]
        else:
            krit_s = [0.0, ls_val, ltot_val - ls_val, ltot_val]
            krit_lbl = ["TS", "SC", "CS", "ST"]

    trans_len = s_full_start - s_start_trans
    s_zero_1 = s_start_trans + trans_len * (en_val / (e_desain + en_val))
    s_zero_2 = s_end_trans - trans_len * (en_val / (e_desain + en_val))
    s_crown_1 = s_start_trans + trans_len * ((2 * en_val) / (e_desain + en_val))
    s_crown_2 = s_end_trans - trans_len * ((2 * en_val) / (e_desain + en_val))

    margin_s = max(40.0, 0.25 * ltot_val)
    s_min_sup = min(0.0, s_start_trans) - margin_s
    s_max_sup = max(ltot_val, s_end_trans) + margin_s
    
    key_nodes = np.unique(np.sort([
        s_min_sup, s_start_trans, s_zero_1, s_crown_1,
        s_full_start, s_full_end, s_crown_2, s_zero_2, s_end_trans, s_max_sup
    ]))
    
    e_luar = np.zeros_like(key_nodes)
    e_dalam = np.zeros_like(key_nodes)
    for i, sv in enumerate(key_nodes):
        if sv <= s_start_trans or sv >= s_end_trans:
            e_luar[i] = -en_val
            e_dalam[i] = -en_val
        elif s_start_trans < sv < s_full_start:
            frac = (sv - s_start_trans) / trans_len
            e_luar[i] = -en_val + frac * (e_desain + en_val)
            e_dalam[i] = min(-en_val, -e_luar[i])
        elif s_full_start <= sv <= s_full_end:
            e_luar[i] = e_desain
            e_dalam[i] = -e_desain
        elif s_full_end < sv < s_end_trans:
            frac = (s_end_trans - sv) / trans_len
            e_luar[i] = -en_val + frac * (e_desain + en_val)
            e_dalam[i] = min(-en_val, -e_luar[i])
            
    sta_nodes = sta_0 + key_nodes
    fig_sup = go.Figure()
    fig_sup.add_trace(go.Scatter(x=sta_nodes, y=e_luar, mode='lines', name='Tepi Perkerasan Luar', line=dict(color='#C2185B', width=3), fill='tozeroy', fillcolor='rgba(194, 24, 91, 0.08)'))
    fig_sup.add_trace(go.Scatter(x=sta_nodes, y=e_dalam, mode='lines', name='Tepi Perkerasan Dalam', line=dict(color='#1565C0', width=3), fill='tozeroy', fillcolor='rgba(21, 101, 192, 0.06)'))
    fig_sup.add_hline(y=0, line_width=1.2, line_color="#424242")
    fig_sup.add_trace(go.Scatter(x=[sta_0 + s_zero_1, sta_0 + s_zero_2], y=[0.0, 0.0], mode='markers', name='Kemiringan Datar (0%)', marker=dict(size=8, color='#111111', symbol='circle')))
    
    y_top_limit = max(e_desain + 3.5, 13.0)
    y_bot_limit = min(-e_desain - 3.5, -13.0)
    for ks, klbl in zip(krit_s, krit_lbl):
        sta_k = sta_0 + ks
        fig_sup.add_vline(x=sta_k, line_width=1, line_dash="dot", line_color="#757575")
        fig_sup.add_annotation(x=sta_k, y=y_top_limit - 1.2, text=f"<b>{klbl}</b><br>{sta_str(sta_k)}", showarrow=False, font=dict(color="#3A0056", size=10.5), bgcolor="rgba(255,255,255,0.92)", bordercolor="#F48FB1", borderwidth=1, borderpad=3)
        
    mid_full_sta = sta_0 + (s_full_start + s_full_end) / 2
    fig_sup.add_annotation(x=mid_full_sta, y=e_desain + 0.9, text=f"<b>+e = {e_desain:.2f}%</b>", showarrow=False, font=dict(color="#AD1457", size=11.5), bgcolor="rgba(255,255,255,0.92)", borderpad=2)
    fig_sup.add_annotation(x=mid_full_sta, y=-e_desain - 1.1, text=f"<b>-e = -{e_desain:.2f}%</b>", showarrow=False, font=dict(color="#1565C0", size=11.5), bgcolor="rgba(255,255,255,0.92)", borderpad=2)
    
    tick_step = 50.0
    x_tick_vals = np.arange(math.ceil(sta_nodes[0] / tick_step) * tick_step, math.floor(sta_nodes[-1] / tick_step) * tick_step + 1, tick_step)
    x_tick_texts = [sta_str(v, d=0) for v in x_tick_vals]
    fig_sup.update_layout(
        title=dict(text="Diagram Superelevasi (Rotasi Sumbu Jalan)", font=dict(color='#3A0056', size=16), x=0.5, xanchor='center'),
        xaxis=dict(title="Stasioning (STA)", tickmode='array', tickvals=x_tick_vals, ticktext=x_tick_texts, gridcolor='#E0E0E0', showline=True, linecolor='#333333', mirror=True),
        yaxis=dict(title="Kemiringan Melintang (%)", range=[y_bot_limit, y_top_limit], gridcolor='#E0E0E0', showline=True, linecolor='#333333', mirror=True),
        legend=dict(orientation="h", yanchor="bottom", y=0.02, xanchor="center", x=0.5, bgcolor="rgba(255, 255, 255, 0.95)", bordercolor="#F48FB1", borderwidth=1, font=dict(color="#3A0056", size=11)),
        plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF", margin=dict(l=60, r=40, t=60, b=50), height=530
    )
    st.plotly_chart(fig_sup, use_container_width=True)

# ----------------------------------------------------------------------------
# TAB 3: POTONGAN MELINTANG INTERAKTIF (TILT CROSS-SECTION)
# ----------------------------------------------------------------------------
with tab3:
    st.markdown("<h3 style='color: #3A0056;'>📐 Potongan Melintang Interaktif & Sudut Kemiringan Kendaraan</h3>", unsafe_allow_html=True)
    st.caption("Geser posisi stasioning untuk mengamati perubahan kemiringan melintang jalan dan sudut banking mobil secara langsung!")

    sel_sta = st.slider("Pilih Posisi Stasioning (STA):", float(r["sta_awal"] - 50), float(r["sta_akhir"] + 50), float(r["sta_awal"] + r["Ls"]), 5.0)
    
    s_curr = sel_sta - r["sta_awal"]
    if s_curr <= s_start_trans or s_curr >= s_end_trans:
        cur_eluar, cur_edalam = -en_val, -en_val
        zone_name = "Bagian Lurus (Normal Crown)"
    elif s_start_trans < s_curr < s_full_start:
        frac = (s_curr - s_start_trans) / trans_len
        cur_eluar = -en_val + frac * (e_desain + en_val)
        cur_edalam = min(-en_val, -cur_eluar)
        zone_name = "Lengkung Peralihan (Runoff / Runout)"
    elif s_full_start <= s_curr <= s_full_end:
        cur_eluar = e_desain
        cur_edalam = -e_desain
        zone_name = "Lengkung Penuh (Full Superelevation)"
    else:
        frac = (s_end_trans - s_curr) / trans_len
        cur_eluar = -en_val + frac * (e_desain + en_val)
        cur_edalam = min(-en_val, -cur_eluar)
        zone_name = "Transisi Keluar Lengkung Peralihan"

    col_cs1, col_cs2 = st.columns([3, 1])
    with col_cs2:
        st.markdown(f"""
        <div class="tulip-card">
            <h4 style="margin: 0 0 10px 0; font-size: 15px; color: #880E4F;">STATUS POSISI STA</h4>
            <p style="margin: 2px 0;"><b>STA:</b> {sta_str(sel_sta)}</p>
            <p style="margin: 2px 0; font-size: 12px; color: #AD1457;"><b>Zona:</b> {zone_name}</p>
            <hr style="border: 0.5px solid #F8BBD0; margin: 10px 0;">
            <p style="margin: 2px 0; color: #C2185B;"><b>Tepi Luar:</b> {cur_eluar:+.2f}%</p>
            <p style="margin: 2px 0; color: #1565C0;"><b>Tepi Dalam:</b> {cur_edalam:+.2f}%</p>
            <p style="margin: 2px 0; color: #2E7D32;"><b>Lebar Jalan B:</b> {B:.1f} m</p>
        </div>
        """, unsafe_allow_html=True)

    with col_cs1:
        fig_cs = go.Figure()
        fig_cs.add_hline(y=0, line_dash="dash", line_color="#9E9E9E")
        
        half_b = B / 2.0
        y_left = (cur_eluar / 100.0) * half_b if arah == "Kanan" else (cur_edalam / 100.0) * half_b
        y_right = (cur_edalam / 100.0) * half_b if arah == "Kanan" else (cur_eluar / 100.0) * half_b
        
        fig_cs.add_trace(go.Scatter(
            x=[-half_b, 0, half_b], y=[y_left, 0, y_right],
            mode='lines+markers', name='Permukaan Aspal',
            line=dict(color='#AD1457', width=4),
            marker=dict(size=8, color='#C2185B')
        ))
        
        fig_cs.add_trace(go.Scatter(
            x=[-half_b, 0, half_b, half_b, -half_b],
            y=[y_left, 0, y_right, y_right - 0.35, y_left - 0.35],
            fill='toself', fillcolor='rgba(244, 143, 177, 0.25)',
            line=dict(color='rgba(0,0,0,0)'), name='Lapis Pondasi Agregat', hoverinfo='skip'
        ))

        car_x = -half_b / 2 if arah == "Kanan" else half_b / 2
        car_y = (y_left / 2) if arah == "Kanan" else (y_right / 2)
        fig_cs.add_annotation(
            x=car_x, y=car_y + 0.32, text="🚘", font=dict(size=36),
            showarrow=False
        )
        
        fig_cs.update_layout(
            title=dict(text=f"Potongan Melintang Jalan di STA {sta_str(sel_sta)}", font=dict(color='#3A0056', size=15)),
            xaxis=dict(title="Offset Melintang (m)", range=[-half_b - 2, half_b + 2], gridcolor='#E0E0E0'),
            yaxis=dict(title="Elevasi Relatif (m)", range=[-1.1, 1.1], gridcolor='#E0E0E0'),
            height=400, plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF", margin=dict(l=40, r=40, t=50, b=40)
        )
        st.plotly_chart(fig_cs, use_container_width=True)

# ----------------------------------------------------------------------------
# TAB 4: PROFIL VERTIKAL & GALIAN/TIMBUNAN
# ----------------------------------------------------------------------------
with tab4:
    st.markdown("<h3 style='color: #3A0056;'>⛰️ Profil Vertikal & Neraca Galian-Timbunan (Earthwork)</h3>", unsafe_allow_html=True)
    st.caption("Perhitungan lengkung vertikal cembung/cekung serta evaluasi galian dan timbunan terhadap tanah asli.")

    col_pv1, col_pv2 = st.columns(2)
    with col_pv1:
        st.markdown("<h5 style='color: #880E4F;'>📌 Tabel Titik Perpotongan Vertikal (PPV / PVI)</h5>", unsafe_allow_html=True)
        pvi_input = st.data_editor(
            st.session_state.pvi_df,
            num_rows="dynamic", key="pvi_editor", hide_index=True, use_container_width=True
        )
    with col_pv2:
        st.markdown("<h5 style='color: #880E4F;'>🌾 Tabel Elevasi Tanah Asli</h5>", unsafe_allow_html=True)
        tanah_input = st.data_editor(
            st.session_state.tanah_df,
            num_rows="dynamic", key="tanah_editor", hide_index=True, use_container_width=True
        )

    pvi_clean = pvi_input.dropna()
    tanah_clean = tanah_input.dropna().sort_values(tanah_input.columns[0])
    
    if len(pvi_clean) >= 2:
        P = hitung_profil(pvi_clean, V)
        fig_vert = go.Figure()
        
        fig_vert.add_trace(go.Scatter(
            x=tanah_clean.iloc[:, 0], y=tanah_clean.iloc[:, 1],
            mode='lines+markers', name='Tanah Asli',
            line=dict(color='#8D6E63', width=2, dash='dot'),
            marker=dict(size=5, color='#6D4C41')
        ))
        
        fig_vert.add_trace(go.Scatter(
            x=P["x"], y=P["y"], mode='lines', name='Elevasi Rencana Jalan',
            line=dict(color='#AD1457', width=3.2)
        ))
        
        fig_vert.add_trace(go.Scatter(
            x=P["sta"], y=P["el"], mode='markers', name='Titik PPV (PVI)',
            marker=dict(size=10, color='#D81B60', symbol='triangle-up')
        ))

        for k in P["kurva"]:
            fig_vert.add_annotation(
                x=k["sta"], y=k["el"],
                text=f"<b>{k['jenis']}</b><br>Lv={k['Lv']:.0f}m | Ev={k['Ev']:.2f}m",
                showarrow=True, arrowhead=2, arrowcolor="#D81B60",
                ay=-32 if k["A"] < 0 else 32,
                bgcolor="rgba(255,255,255,0.92)", font=dict(color="#3A0056", size=9.5),
                bordercolor="#F48FB1", borderwidth=1
            )

        fig_vert.add_vrect(
            x0=r["sta_awal"], x1=r["sta_akhir"],
            fillcolor="rgba(244, 143, 177, 0.15)", layer="below", line_width=1,
            line_dash="dash", line_color="rgba(194, 24, 91, 0.4)",
            annotation_text="Zona Tikungan Horizontal", annotation_position="top left",
            annotation_font=dict(color="#880E4F", size=10)
        )

        fig_vert.update_layout(
            title=dict(text="Profil Memanjang Trase Jalan (Alignment Vertikal)", font=dict(color='#3A0056', size=16)),
            xaxis=dict(title="Stasioning (STA)", gridcolor='#E0E0E0'),
            yaxis=dict(title="Elevasi (m DPL)", gridcolor='#E0E0E0'),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5, bgcolor="rgba(255, 255, 255, 0.95)", bordercolor="#F48FB1", borderwidth=1),
            plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF", height=460, margin=dict(l=40, r=40, t=50, b=40)
        )
        st.plotly_chart(fig_vert, use_container_width=True)

        xs_ew = np.arange(math.ceil(P["x"][0] / 50) * 50, P["x"][-1] + 0.1, 50)
        rn_ew = np.interp(xs_ew, P["x"], P["y"])
        ta_ew = np.interp(xs_ew, tanah_clean.iloc[:, 0], tanah_clean.iloc[:, 1])
        diff_ew = rn_ew - ta_ew

        df_ew = pd.DataFrame({
            "STA": [sta_str(v, 0) for v in xs_ew],
            "Elevasi Rencana (m)": np.round(rn_ew, 2),
            "Elevasi Tanah (m)": np.round(ta_ew, 2),
            "Tinggi Cut / Fill (m)": np.round(diff_ew, 2),
            "Keterangan": np.where(diff_ew > 0, "🟢 Timbunan (Fill)", "🔴 Galian (Cut)")
        })

        col_ew1, col_ew2 = st.columns([2, 1])
        with col_ew1:
            st.dataframe(df_ew, hide_index=True, use_container_width=True)
        with col_ew2:
            n_cut = np.sum(diff_ew < 0)
            n_fill = np.sum(diff_ew > 0)
            st.markdown(f"""
            <div class="tulip-card">
                <h4 style="margin: 0 0 10px 0; color: #880E4F;">RINGKASAN EARTHWORK</h4>
                <p><b>Jarak Pandang Henti (Jh):</b> {P['Jh']:.1f} m</p>
                <p><b>Titik Galian:</b> {n_cut} stasioning</p>
                <p><b>Titik Timbunan:</b> {n_fill} stasioning</p>
                <p><b>Kelandaian Maks:</b> {np.max(np.abs(P['g'])):.2f}%</p>
            </div>
            """, unsafe_allow_html=True)
    else:
        st.info("Isi minimal 2 titik PPV dengan nilai STA yang berbeda.")

# ----------------------------------------------------------------------------
# TAB 5: SIMULASI 3D MOBIL JALAN
# ----------------------------------------------------------------------------
with tab5:
    st.markdown("<h3 style='color: #3A0056;'>🚘 Visualisasi 3D Jalan & Simulator Mobil Jalan Otomatis</h3>", unsafe_allow_html=True)
    
    v_safe = r["v_safe"]
    if v_sim <= v_safe:
        car_color = "#1B5E20"
        status_box = f"""
        <div class="status-card status-oke">
            🟢 <b>KONDISI HIJAU (OKE):</b> Mobil Melaju Sangat Aman!<br>
            Kecepatan simulasi ({v_sim} km/jam) berada di bawah batas kecepatan aman tikungan ({v_safe:.1f} km/jam).
        </div>
        """
    elif v_sim <= v_safe + 10:
        car_color = "#E65100"
        status_box = f"""
        <div class="status-card status-sedang">
            🟡 <b>KONDISI KUNING (SEDANG / WASPADA):</b> Mobil Mendekati Batas Aman!<br>
            Kecepatan simulasi ({v_sim} km/jam) sedikit melampaui V_aman ({v_safe:.1f} km/jam). Harap kurangi kecepatan.
        </div>
        """
    else:
        car_color = "#B71C1C"
        status_box = f"""
        <div class="status-card status-bahaya">
            🔴 <b>KONDISI MERAH (TIDAK OKE / BAHAYA):</b> Over-Speed!<br>
            Kecepatan simulasi ({v_sim} km/jam) melampaui batas aman ({v_safe:.1f} km/jam). Berisiko selip atau kecelakaan.
        </div>
        """
    st.markdown(status_box, unsafe_allow_html=True)

    step_frame = max(1, len(s_vals) // 60)
    frame_indices = list(range(0, len(s_vals), step_frame))
    if frame_indices[-1] != len(s_vals) - 1:
        frame_indices.append(len(s_vals) - 1)
    frame_duration = max(30, int(4000 / max(v_sim, 10)))

    frames, slider_steps = [], []
    for idx in frame_indices:
        sta_lbl = f"STA {sta_str(r['sta_awal'] + s_vals[idx])}"
        frames.append(go.Frame(
            data=[go.Scatter3d(
                x=[x_horiz[idx]], y=[y_horiz[idx]], z=[z_elev[idx] + 1.2],
                mode='markers+text', marker=dict(size=12, color=car_color, symbol='diamond'),
                text=[f"🚘 {sta_lbl}"], textposition="top center", textfont=dict(color="#111111", size=12)
            )],
            traces=[0], name=str(idx)
        ))
        slider_steps.append({
            "args": [[str(idx)], {"frame": {"duration": 0, "redraw": True}, "mode": "immediate"}],
            "label": f"{s_vals[idx]:.0f}m", "method": "animate"
        })

    fig_3d = go.Figure(
        data=[
            go.Scatter3d(x=[x_horiz[0]], y=[y_horiz[0]], z=[z_elev[0] + 1.2], mode='markers+text', name='Mobil Simulasi', marker=dict(size=12, color=car_color, symbol='diamond'), text=[f"🚘 STA {sta_str(r['sta_awal'] + s_vals[0])}"], textposition="top center", textfont=dict(color="#111111", size=12)),
            go.Scatter3d(x=x_horiz, y=y_horiz, z=z_elev, mode='lines', name='As Jalan 3D', line=dict(color='#3A0056', width=6)),
            go.Scatter3d(x=xl, y=yl, z=z_elev, mode='lines', name='Tepi Kiri 3D', line=dict(color='#C2185B', width=3)),
            go.Scatter3d(x=xr, y=yr, z=z_elev, mode='lines', name='Tepi Kanan 3D', line=dict(color='#C2185B', width=3)),
        ],
        frames=frames
    )
    fig_3d.update_layout(
        title=dict(text="Visualisasi 3D Trase Jalan & Simulator Mobil (Klik ▶️ Jalan di Kiri Bawah Grafik)", font=dict(color='#3A0056', size=16)),
        template="plotly_white",
        scene=dict(xaxis_title="X (m)", yaxis_title="Y (m)", zaxis_title="Elevasi Z (m)", aspectmode='data'),
        legend=dict(bgcolor="rgba(255,255,255,0.95)", bordercolor="#F48FB1", borderwidth=1, font=dict(color="#3A0056")),
        updatemenus=[{
            "type": "buttons", "showactive": True, "x": 0.02, "y": 0.02, "xanchor": "left", "yanchor": "bottom",
            "bgcolor": "#FCE4EC", "bordercolor": "#AD1457", "font": {"color": "#880E4F", "size": 13},
            "buttons": [
                {"label": "▶️ Jalan", "method": "animate", "args": [None, {"frame": {"duration": frame_duration, "redraw": True}, "fromcurrent": True, "transition": {"duration": 0}}]},
                {"label": "⏸️ Pause", "method": "animate", "args": [[None], {"frame": {"duration": 0, "redraw": False}, "mode": "immediate", "transition": {"duration": 0}}]}
            ]
        }],
        sliders=[{
            "active": 0, "yanchor": "top", "xanchor": "left",
            "currentvalue": {"font": {"size": 13, "color": "#3A0056"}, "prefix": "Posisi Relatif: ", "visible": True, "xanchor": "right"},
            "pad": {"b": 10, "t": 40}, "len": 0.82, "x": 0.16, "y": 0.02, "steps": slider_steps
        }],
        margin=dict(l=0, r=0, b=40, t=50), height=620, paper_bgcolor="#FFFFFF", font=dict(color="#111111")
    )
    st.plotly_chart(fig_3d, use_container_width=True)

# ----------------------------------------------------------------------------
# TAB 6: GAMBAR KERJA WORD (.DOCX) DENGAN ETIKET KANAN
# ----------------------------------------------------------------------------
with tab6:
    st.markdown("<h3 style='color: #3A0056;'>📐 Hasil Gambar Kerja Skalatis (Hitam Putih + Etiket Kanan) — Format Word (.DOCX)</h3>", unsafe_allow_html=True)
    st.info("📄 Dokumen Word dicetak pada ukuran **A3 Landscape (420 × 297 mm)** hitam-putih dalam **2 lembar gambar kerja skalatis** lengkap dengan garis dimensi nilai perhitungan dan **Etiket di sisi kanan**.")

    with st.expander("✏️ Isi / Edit Data Etiket Gambar Kerja (Title Block Kanan)", expanded=False):
        col_et1, col_et2, col_et3 = st.columns(3)
        with col_et1:
            et_logo = st.text_input("LOGO / Instansi", "LABORATORIUM JALAN RAYA")
            et_kegiatan = st.text_input("KEGIATAN", "PERENCANAAN GEOMETRIK JALAN")
            et_pekerjaan = st.text_input("PEKERJAAN", f"DESAIN TIKUNGAN {r['jenis'].upper()}")
        with col_et2:
            et_lokasi = st.text_input("LOKASI", f"STA {sta_str(r['sta_awal'])} s/d {sta_str(r['sta_akhir'])}")
            et_diperiksa = st.text_input("DI PERIKSA OLEH :", "Dosen Pembimbing / Engineer")
            et_digambar = st.text_input("DI GAMBAR OLEH :", "Perencana Geometrik")
        with col_et3:
            et_tanggal = st.text_input("TANGGAL", date.today().strftime("%d-%m-%Y"))
            et_catatan = st.text_area("CATATAN (Kosongkan untuk otomatis dari hasil hitung)", "", height=110)

    meta_etiket = dict(
        logo=et_logo,
        catatan=et_catatan,
        kegiatan=et_kegiatan,
        pekerjaan=et_pekerjaan,
        lokasi=et_lokasi,
        diperiksa=et_diperiksa,
        digambar=et_digambar,
        tanggal=et_tanggal
    )

    prev1_buf, sk1 = render_cad_denah_bw(r, B, arah)
    prev2_buf, sk2 = render_cad_superelevasi_bw(r)

    col_p1, col_p2 = st.columns(2)
    with col_p1:
        st.markdown(f"**Lembar 01 — Denah Tikungan Skalatis (Skala 1 : {sk1:,})**")
        st.image(prev1_buf, use_container_width=True)
    with col_p2:
        st.markdown(f"**Lembar 02 — Diagram Superelevasi Skalatis (Skala 1 : {sk2:,})**")
        st.image(prev2_buf, use_container_width=True)

    docx_bytes, s_out1, s_out2 = generate_word_2_sheets_landscape(r, B, arah, meta_etiket)
    st.download_button(
        f"📥 Unduh File Word Gambar Kerja Skalatis (.DOCX) — 2 Lembar A3 Landscape + Etiket Kanan",
        docx_bytes,
        file_name=f"Gambar_Kerja_Geometri_Jalan_Skala_1_{s_out1}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )

# ----------------------------------------------------------------------------
# TAB 7: TABEL HITUNGAN & EXCEL REPORT (BM 2021)
# ----------------------------------------------------------------------------
with tab7:
    st.markdown("<h3 style='color: #3A0056;'>📋 Tabel Hasil Perhitungan Tikungan & Laporan Excel</h3>", unsafe_allow_html=True)
    
    a, b = ("TC", "CT") if fc else ("TS", "ST")
    rows = [
        ("Jenis Tikungan", r["jenis"], "Acuan Kriteria R_fc"),
        ("Kecepatan Rencana (Vr)", f"{r['V']:.0f}", "km/jam"),
        ("Sudut Tikungan (Δ)", f"{r['delta']:.2f}", "°"),
        ("Jari-Jari Minimum (Rmin)", f"{r['Rmin']:.2f}", "m"),
        ("Jari-Jari Rencana (Rc)", f"{r['Rc']:.2f}", "m"),
        ("Superelevasi Maksimum (e_max)", f"{r['emax']*100:.1f}", "%"),
        ("Superelevasi Desain (e)", f"{r['e']*100:.2f}", "%"),
        ("Panjang Transisi (Ls dipakai)", f"{r['Ls_sup']:.2f}", "m"),
        ("Panjang Tangen (Ts)", f"{r['T']:.3f}", "m"),
        ("Jarak Eksternal (Es)", f"{r['E']:.3f}", "m"),
        ("Koordinat Spiral (Xc / Xs)", f"{r['Xs']:.3f}", "m"),
        ("Koordinat Spiral (yc / Ys)", f"{r['Ys']:.3f}", "m"),
        ("Pergeseran Tangen (p)", f"{r['p']:.3f}", "m"),
        ("Absis Pusat Lingkaran (k)", f"{r['k']:.3f}", "m"),
        ("Panjang Total Tikungan (Ltot)", f"{r['Ltot']:.3f}", "m"),
        ("Kecepatan Aman Maksimum (V_safe)", f"{r['v_safe']:.2f}", "km/jam"),
        (f"STA {a} (Awal Tikungan)", sta_str(r["sta_awal"]), "m"),
        ("STA PI", sta_str(r["sta_pi"]), "m"),
        (f"STA {b} (Akhir Tikungan)", sta_str(r["sta_akhir"]), "m"),
    ]
    
    st.dataframe(pd.DataFrame(rows, columns=["Parameter Geometrik", "Nilai Calculated", "Satuan / Keterangan"]), hide_index=True, use_container_width=True)

    excel_bytes = generate_excel_bm2021(r, P_global)
    st.download_button(
        "📊 Unduh Laporan Excel Ber-Formula Otomatis (.XLSX)",
        excel_bytes, file_name="Perhitungan_Geometri_Jalan_BM2021.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
