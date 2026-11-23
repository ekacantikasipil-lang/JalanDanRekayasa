"""
==============================================================================
🌸 GEOMETRI JALAN INTERAKTIF - SOFT PINK & LILAC EDITION
Sistem Terpadu Perencanaan Geometrik Jalan Raya & Simulasi Trase Realistis
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
# 1. KONFIGURASI HALAMAN STREAMLIT & TEMA SOFT PINK & LILAC (CERAH & ANTI-MONOTON)
# ----------------------------------------------------------------------------
st.set_page_config(
    page_title="Geometri Jalan Interaktif - Soft Pink & Lilac Edition",
    layout="wide",
    page_icon="🌸",
    initial_sidebar_state="expanded"
)

# Inisialisasi status opening screen (Tanpa Password!)
if "started" not in st.session_state:
    st.session_state.started = False

# CSS Kustom: Soft Pastel Pink & Lilac (Super Cerah, Glowing, Estetis & 3D Cards)
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600;700&display=swap');
    
    /* 1. Global Background: Terang, Cerah, Gradasi Pastel Pink & Lavender */
    .stApp {
        background: linear-gradient(135deg, #FFF6F9 0%, #FDF2F8 30%, #F5EEFA 70%, #FFF9FC 100%);
        color: #38004C;
        font-family: 'Plus Jakarta Sans', sans-serif;
    }
    
    /* 2. Sidebar Soft Pastel Lilac */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #FFF0F6 0%, #FAF0FA 100%) !important;
        border-right: 2px solid #F8BBD0;
        box-shadow: 4px 0 24px rgba(233, 30, 99, 0.06);
    }
    section[data-testid="stSidebar"] h1, 
    section[data-testid="stSidebar"] h2, 
    section[data-testid="stSidebar"] h3 {
        color: #4A154B !important;
        font-weight: 800 !important;
    }
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] span {
        color: #5C1D61 !important;
        font-weight: 600 !important;
    }

    /* 3. Kartu 3D Soft Glossy */
    .card-3d {
        background: #FFFFFF;
        border-radius: 20px;
        border: 2px solid #FCE4EC;
        padding: 24px;
        box-shadow: 0 12px 30px rgba(233, 30, 99, 0.08), 0 4px 12px rgba(156, 39, 176, 0.05);
        margin-bottom: 22px;
        transition: transform 0.25s ease, box-shadow 0.25s ease;
    }
    .card-3d:hover {
        transform: translateY(-3px);
        box-shadow: 0 16px 36px rgba(233, 30, 99, 0.14), 0 6px 16px rgba(156, 39, 176, 0.08);
    }

    /* 4. Ikon 3D Bergaya Glossy */
    .icon-badge-3d {
        width: 48px;
        height: 48px;
        border-radius: 14px;
        background: linear-gradient(135deg, #FFF0F5 0%, #F3E5F5 100%);
        border: 1.5px solid #F8BBD0;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        font-size: 24px;
        box-shadow: 0 6px 14px rgba(233, 30, 99, 0.12), inset 0 2px 4px rgba(255, 255, 255, 0.9);
    }

    /* 5. Header Banner Cerah & Elegan */
    .hero-bright {
        background: linear-gradient(135deg, #FFFFFF 0%, #FFF5F9 50%, #FAF0FA 100%);
        border: 2px solid #F8BBD0;
        border-radius: 22px;
        padding: 22px 28px;
        margin-bottom: 24px;
        box-shadow: 0 12px 32px rgba(233, 30, 99, 0.1);
        display: flex;
        justify-content: space-between;
        align-items: center;
        flex-wrap: wrap;
        gap: 16px;
    }
    .hero-bright h1 {
        color: #4A154B;
        font-size: 26px;
        font-weight: 800;
        margin: 0;
        display: flex;
        align-items: center;
        gap: 10px;
    }
    .hero-bright p {
        color: #8E24AA;
        font-size: 13.5px;
        margin: 4px 0 0 0;
        font-weight: 600;
    }
    .pill-tag {
        background: #FFFFFF;
        border: 1.5px solid #F48FB1;
        color: #9C27B0;
        padding: 6px 14px;
        border-radius: 20px;
        font-size: 12.5px;
        font-weight: 700;
        box-shadow: 0 3px 8px rgba(233, 30, 99, 0.08);
        display: inline-flex;
        align-items: center;
        gap: 6px;
    }
    .pill-tag-active {
        background: linear-gradient(135deg, #EC407A 0%, #AB47BC 100%);
        color: #FFFFFF !important;
        border: none;
        box-shadow: 0 4px 15px rgba(236, 64, 122, 0.4);
    }

    /* 6. Grid Metrik KPI 3D Cerah */
    .kpi-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(185px, 1fr));
        gap: 16px;
        margin-bottom: 24px;
    }
    .kpi-card-3d {
        background: #FFFFFF;
        border-radius: 18px;
        padding: 18px 20px;
        border: 1.5px solid #FCE4EC;
        box-shadow: 0 8px 22px rgba(233, 30, 99, 0.07);
        position: relative;
        overflow: hidden;
        transition: transform 0.25s ease, box-shadow 0.25s ease;
    }
    .kpi-card-3d:hover {
        transform: translateY(-4px);
        box-shadow: 0 14px 28px rgba(233, 30, 99, 0.15);
    }
    .kpi-card-3d::before {
        content: '';
        position: absolute;
        top: 0; left: 0; right: 0;
        height: 5px;
        background: linear-gradient(90deg, #F06292, #BA68C8);
    }
    .kpi-title {
        font-size: 11px;
        font-weight: 800;
        color: #7B1FA2;
        text-transform: uppercase;
        letter-spacing: 0.6px;
        margin-bottom: 4px;
    }
    .kpi-val {
        font-size: 24px;
        font-weight: 800;
        color: #38004C;
        font-family: 'JetBrains Mono', monospace;
        line-height: 1.1;
    }
    .kpi-sub {
        font-size: 11.5px;
        font-weight: 600;
        color: #D81B60;
        margin-top: 6px;
    }

    /* 7. Tab Cerah Pink & Lilac */
    .stTabs [data-baseweb="tab-list"] {
        background-color: #FFFFFF !important;
        border-radius: 18px;
        padding: 8px;
        border: 2px solid #F8BBD0;
        box-shadow: 0 6px 20px rgba(233, 30, 99, 0.08);
        gap: 8px;
        margin-bottom: 20px;
    }
    .stTabs [data-baseweb="tab"],
    .stTabs button[role="tab"] {
        color: #8E24AA !important;
        font-weight: 700 !important;
        font-size: 14.5px !important;
        border-radius: 12px;
        padding: 10px 20px;
        background-color: transparent !important;
        border: none !important;
        transition: all 0.2s ease;
    }
    .stTabs [data-baseweb="tab"]:hover,
    .stTabs button[role="tab"]:hover {
        background-color: #FFF0F5 !important;
        color: #D81B60 !important;
    }
    .stTabs [aria-selected="true"],
    .stTabs button[aria-selected="true"] {
        background: linear-gradient(135deg, #EC407A 0%, #AB47BC 100%) !important;
        box-shadow: 0px 6px 18px rgba(236, 64, 122, 0.35) !important;
    }
    .stTabs [aria-selected="true"] p,
    .stTabs [aria-selected="true"] span,
    .stTabs button[aria-selected="true"] p {
        color: #FFFFFF !important;
        font-weight: 800 !important;
    }

    /* 8. Tombol 3D Soft Pink-Ungu */
    .stButton>button {
        background: linear-gradient(135deg, #EC407A 0%, #BA68C8 100%);
        color: white !important;
        border: none;
        border-radius: 14px;
        padding: 12px 26px;
        font-weight: 800;
        font-size: 14.5px;
        box-shadow: 0 6px 18px rgba(236, 64, 122, 0.3);
        transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1);
    }
    .stButton>button:hover {
        background: linear-gradient(135deg, #D81B60 0%, #8E24AA 100%);
        transform: translateY(-2px);
        box-shadow: 0 10px 25px rgba(216, 27, 96, 0.4);
    }

    /* 9. Kotak Input Cerah */
    div[data-baseweb="input"] {
        background-color: #FFFFFF !important;
        border-radius: 12px;
        border: 2px solid #F8BBD0 !important;
    }
    div[data-baseweb="input"]:focus-within {
        border-color: #AB47BC !important;
        box-shadow: 0 0 0 3px rgba(171, 71, 188, 0.2) !important;
    }

    /* 10. Opening Splash Banner */
    .opening-card {
        max-width: 800px;
        margin: 40px auto;
        background: #FFFFFF;
        border-radius: 28px;
        border: 2.5px solid #F8BBD0;
        padding: 44px 38px;
        box-shadow: 0 20px 60px rgba(233, 30, 99, 0.12), 0 8px 20px rgba(156, 39, 176, 0.08);
        text-align: center;
    }
</style>
""", unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# 2. HALAMAN PEMBUKA / OPENING SCREEN (TANPA PASSWORD!)
# ----------------------------------------------------------------------------
if not st.session_state.started:
    st.markdown("""
    <div class="opening-card">
        <div style="font-size: 64px; margin-bottom: 12px; filter: drop-shadow(0 8px 16px rgba(233, 30, 99, 0.25));">
            🌸 🛣️ 🚘
        </div>
        <div class="pill-tag" style="margin-bottom: 16px;">
            SEMESTER 5 APLIKOM • REKAYASA JALAN RAYA
        </div>
        <h1 style="color: #4A154B; font-weight: 800; font-size: 32px; margin: 0 0 10px 0;">
            Geometri Jalan Interaktif
        </h1>
        <p style="color: #8E24AA; font-size: 16px; margin: 0 auto 30px auto; max-width: 580px; font-weight: 500;">
            Aplikasi Desain Tikungan & Peta Realistis Standar Bina Marga (TPGJAK 1997 / RAE 2021) dengan Animasi Pergerakan Mobil Nyata di Trase Jalan.
        </p>
        
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(210px, 1fr)); gap: 16px; margin-bottom: 34px; text-align: left;">
            <div style="background: #FFF5F9; border: 1.5px solid #FCE4EC; border-radius: 16px; padding: 16px;">
                <div style="font-size: 26px; margin-bottom: 6px;">🗺️</div>
                <b style="color: #4A154B; font-size: 14px;">Peta & Trase Real</b>
                <p style="color: #880E4F; font-size: 12px; margin: 4px 0 0 0;">Animasi mobil berbelok di atas aspal dengan marka jalan nyata.</p>
            </div>
            <div style="background: #FAF2FA; border: 1.5px solid #F3E5F5; border-radius: 16px; padding: 16px;">
                <div style="font-size: 26px; margin-bottom: 6px;">📈</div>
                <b style="color: #4A154B; font-size: 14px;">Superelevasi Presisi</b>
                <p style="color: #7B1FA2; font-size: 12px; margin: 4px 0 0 0;">Diagram pencapaian kemiringan tepi luar & dalam jalan.</p>
            </div>
            <div style="background: #FFF5F9; border: 1.5px solid #FCE4EC; border-radius: 16px; padding: 16px;">
                <div style="font-size: 26px; margin-bottom: 6px;">📑</div>
                <b style="color: #4A154B; font-size: 14px;">Gambar CAD A3 Word</b>
                <p style="color: #880E4F; font-size: 12px; margin: 4px 0 0 0;">Siap ekspor 2 lembar A3 lengkap dengan Etiket 60 mm.</p>
            </div>
            <div style="background: #FAF2FA; border: 1.5px solid #F3E5F5; border-radius: 16px; padding: 16px;">
                <div style="font-size: 26px; margin-bottom: 6px;">📊</div>
                <b style="color: #4A154B; font-size: 14px;">Excel 17-Rumus</b>
                <p style="color: #7B1FA2; font-size: 12px; margin: 4px 0 0 0;">Laporan otomatis ber-formula hidup standar akademik.</p>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    col_c1, col_c2, col_c3 = st.columns([1, 1.4, 1])
    with col_c2:
        if st.button("🌸 Buka Aplikasi & Mulai Eksplorasi 🚀", use_container_width=True):
            st.session_state.started = True
            st.rerun()

    st.stop()

# ----------------------------------------------------------------------------
# 3. ENGINE MATEMATIS GEOMETRIK JALAN (TPGJAK 1997 / RAE 2021)
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

# Helper untuk geometri mobil & sorot lampu depan
def get_car_geometry(xc, yc, heading_rad, L=4.6, W=2.0):
    cos_h = math.cos(heading_rad)
    sin_h = math.sin(heading_rad)
    
    # Titik bodi mobil
    corners_local = [
        (-L*0.48, -W*0.48), (L*0.35, -W*0.48), (L*0.50, -W*0.35),
        (L*0.50, W*0.35), (L*0.35, W*0.48), (-L*0.48, W*0.48),
        (-L*0.48, -W*0.48)
    ]
    bx = [xc + px*cos_h - py*sin_h for px, py in corners_local]
    by = [yc + px*sin_h + py*cos_h for px, py in corners_local]
    
    # Kaca kabin mobil
    cabin_local = [
        (-L*0.16, -W*0.36), (L*0.18, -W*0.36), (L*0.25, -W*0.30),
        (L*0.25, W*0.30), (L*0.18, W*0.36), (-L*0.16, W*0.36),
        (-L*0.16, -W*0.36)
    ]
    cx = [xc + px*cos_h - py*sin_h for px, py in cabin_local]
    cy = [yc + px*sin_h + py*cos_h for px, py in cabin_local]
    
    # Sorot lampu depan (Headlights Beam)
    beam_len = 14.0
    beam_spread = 3.6
    beam_local = [
        (L*0.50, -W*0.35), (L*0.50 + beam_len, -beam_spread),
        (L*0.50 + beam_len, beam_spread), (L*0.50, W*0.35),
        (L*0.50, -W*0.35)
    ]
    beam_x = [xc + px*cos_h - py*sin_h for px, py in beam_local]
    beam_y = [yc + px*sin_h + py*cos_h for px, py in beam_local]
    
    return bx, by, cx, cy, beam_x, beam_y

# ----------------------------------------------------------------------------
# 4. SIDEBAR INPUT & KONTROL PARAMETER
# ----------------------------------------------------------------------------
with st.sidebar:
    st.markdown("<h2 style='margin-bottom: 2px;'>⚙️ Parameter Rencana</h2>", unsafe_allow_html=True)
    st.caption("Standar Acuan: TPGJAK 1997 / RAE 2021")
    
    if st.button("🏠 Halaman Opening", use_container_width=True):
        st.session_state.started = False
        st.rerun()

    st.markdown("---")
    st.subheader("1. Kriteria Rencana")
    V = st.number_input("Kecepatan Rencana Vr (km/jam)", 20, 120, 60, 10)
    v_sim = st.number_input("Kecepatan Simulasi Mobil (km/jam)", 10, 150, 60, 5)
    delta = st.number_input("Sudut Tikungan Δ (°)", 1.0, 179.0, 45.0, 0.5)
    emax = st.number_input("e maks (%)", 4.0, 12.0, 10.0, 0.5) / 100
    en = st.number_input("Kemiringan Normal en (%)", 1.0, 4.0, 2.0, 0.5) / 100
    
    Rmin = r_min(V, emax)
    st.caption(f"ℹ️ *R minimum acuan TPGJAK = {Rmin:.2f} m*")
    
    mode = st.radio("Jari-Jari Tikungan Rc", ["Otomatis (Dibulatkan)", "Manual Input"])
    if mode.startswith("Oto"):
        Rc = float(math.ceil(Rmin / 10) * 10)
    else:
        Rc = st.number_input("Rc (m)", 10.0, 5000.0, 150.0, 5.0)
        
    sta_pi = st.number_input("STA PI (m)", 0.0, 100000.0, 300.0, 10.0)
    arah = st.radio("Arah Tikungan", ["Kanan", "Kiri"], horizontal=True)
    
    st.markdown("---")
    st.subheader("2. Lebar Perkerasan & Bahu")
    B = st.number_input("Lebar Perkerasan Aspal B (m)", 3.0, 30.0, 7.0, 0.5)
    Bahu = st.number_input("Lebar Bahu Jalan (m)", 0.5, 5.0, 1.5, 0.25)

# Kalkulasi Trase Jalan Utama
r = hitung_tikungan(V, delta, emax, en, Rc, sta_pi)
sg = 1 if arah == "Kiri" else -1
Ls, Lc = r["Ls"], r["Lc"]
dl = math.radians(r["delta"])
Ltot = r["Ltot"]
ext = math.ceil(max(140, 1.5 * r["T"]) / 10) * 10
s_vals = np.unique(np.concatenate([np.arange(-ext, Ltot + ext, 0.5), [0, Ls, Ls + Lc, Ltot]]))
th_vals = heading(s_vals, Ls, Lc, Rc, dl)
dS = np.diff(s_vals)
x_horiz = np.concatenate([[0], np.cumsum(dS * (np.cos(th_vals[:-1]) + np.cos(th_vals[1:])) / 2)])
y_horiz = np.concatenate([[0], np.cumsum(dS * (np.sin(th_vals[:-1]) + np.sin(th_vals[1:])) / 2)])
i0 = np.argmin(np.abs(s_vals))
x_horiz, y_horiz = x_horiz - x_horiz[i0], (y_horiz - y_horiz[i0]) * sg
ths_ = th_vals * sg
nx, ny = -np.sin(ths_), np.cos(ths_)

# Data Profil Vertikal
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
# 5. ENGINE GAMBAR CAD WORD HITAM-PUTIH (A3 LANDSCAPE + ETIKET KANAN 60MM)
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

    mirror = 1.0 if arah == "Kanan" else -1.0
    d_rad = math.radians(delta_deg)
    half_d = d_rad / 2.0

    R_eff = Rc if fc else (Rc + p_val)
    x_O, y_O = 0.0, 0.0
    x_PI, y_PI = 0.0, R_eff / math.cos(half_d)

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

    all_x = np.concatenate([curve_x, [p_tan_start[0], p_tan_end[0], p_PI[0], p_O[0]]])
    all_y = np.concatenate([curve_y, [p_tan_start[1], p_tan_end[1], p_PI[1], p_O[1]]])
    x_min, x_max = np.min(all_x), np.max(all_x)
    y_min, y_max = np.min(all_y), np.max(all_y)

    margin_x = 0.16 * max(x_max - x_min, 30.0)
    margin_y = 0.16 * max(y_max - y_min, 30.0)
    world_w_m = (x_max - x_min) + 2 * margin_x
    world_h_m = (y_max - y_min) + 2 * margin_y

    s_w = (world_w_m * 1000.0) / draw_w_mm
    s_h = (world_h_m * 1000.0) / draw_h_mm
    s_req = max(s_w, s_h)

    std_scales = [100, 200, 250, 500, 750, 1000, 1250, 1500, 2000, 2500, 5000, 10000]
    scale_val = std_scales[-1]
    for sc in std_scales:
        if sc >= s_req:
            scale_val = sc
            break

    fig_w_in = draw_w_mm / 25.4
    fig_h_in = draw_h_mm / 25.4
    fig, ax = plt.subplots(figsize=(fig_w_in, fig_h_in), dpi=300)
    fig.patch.set_facecolor('white')
    ax.set_facecolor('white')

    plot_w_m = (draw_w_mm / 1000.0) * scale_val
    plot_h_m = (draw_h_mm / 1000.0) * scale_val
    cx = 0.5 * (x_min + x_max)
    cy = 0.5 * (y_min + y_max)
    ax.set_xlim(cx - plot_w_m / 2.0, cx + plot_w_m / 2.0)
    ax.set_ylim(cy - plot_h_m / 2.0, cy + plot_h_m / 2.0)
    ax.set_aspect('equal', adjustable='box')
    ax.axis('off')

    # Trase As Jalan
    ax.plot([p_tan_start[0], p_TS[0]], [p_tan_start[1], p_TS[1]], color='#000000', linewidth=1.4, zorder=5)
    ax.plot(curve_x, curve_y, color='#000000', linewidth=1.5, zorder=5)
    ax.plot([p_ST[0], p_tan_end[0]], [p_ST[1], p_tan_end[1]], color='#000000', linewidth=1.4, zorder=5)

    # Garis Tangen Putus-putus
    ax.plot([p_TS[0], p_PI[0]], [p_TS[1], p_PI[1]], color='#000000', linewidth=0.75, linestyle='--', zorder=3)
    ax.plot([p_ST[0], p_PI[0]], [p_ST[1], p_PI[1]], color='#000000', linewidth=0.75, linestyle='--', zorder=3)

    # Garis Hubung Pusat O
    ax.plot([p_O[0], p_PI[0]], [p_O[1], p_PI[1]], color='#000000', linewidth=0.75, zorder=3)
    if not fc:
        ax.plot([p_O[0], p_SC[0]], [p_O[1], p_SC[1]], color='#000000', linewidth=0.75, zorder=3)
        ax.plot([p_O[0], p_CS[0]], [p_O[1], p_CS[1]], color='#000000', linewidth=0.75, zorder=3)
    else:
        ax.plot([p_O[0], p_TS[0]], [p_O[1], p_TS[1]], color='#000000', linewidth=0.75, zorder=3)
        ax.plot([p_O[0], p_ST[0]], [p_O[1], p_ST[1]], color='#000000', linewidth=0.75, zorder=3)

    nodes_x = [p_TS[0], p_ST[0], p_SCS[0]]
    nodes_y = [p_TS[1], p_ST[1], p_SCS[1]]
    if not fc:
        nodes_x.extend([p_SC[0], p_CS[0]])
        nodes_y.extend([p_SC[1], p_CS[1]])
    ax.scatter(nodes_x, nodes_y, facecolor='white', edgecolor='#000000', s=20, linewidths=0.9, zorder=7)

    lbl_a, lbl_b = ("TC", "CT") if fc else ("TS", "ST")
    ax.text(p_PI[0], p_PI[1] + 0.025 * world_h_m, f"PI (STA {sta_str(r['sta_pi'])})", fontsize=8, fontweight='bold', ha='center', va='bottom', color='#000000')
    ax.text(p_TS[0] - mirror * 0.025 * world_w_m, p_TS[1] - 0.015 * world_h_m, f"{lbl_a}\n({sta_str(r['sta_awal'])})", fontsize=7.5, fontweight='bold', ha='center', va='top', color='#000000')
    ax.text(p_ST[0] + mirror * 0.025 * world_w_m, p_ST[1] - 0.015 * world_h_m, f"{lbl_b}\n({sta_str(r['sta_akhir'])})", fontsize=7.5, fontweight='bold', ha='center', va='top', color='#000000')

    # Dimensi Ts dan Es
    draw_dim_line(ax, p_TS, p_PI, f"Ts = {Ts:.2f} m", offset_vec=mirror * 0.035 * world_w_m * n_in, fontsize=7)
    draw_dim_line(ax, p_PI, p_ST, f"Ts = {Ts:.2f} m", offset_vec=mirror * 0.035 * world_w_m * n_out, fontsize=7)
    draw_dim_line(ax, p_PI, p_SCS, f"Es = {Es:.2f} m", offset_vec=(mirror * 0.025 * world_w_m, 0), fontsize=7)

    # Batang Skala Grafis
    bar_phys_mm = 50.0
    bar_real_m = (bar_phys_mm / 1000.0) * scale_val
    sb_x0 = cx + plot_w_m / 2.0 - bar_real_m - 0.04 * plot_w_m
    sb_y0 = cy - plot_h_m / 2.0 + 0.05 * plot_h_m
    sb_h = 0.015 * plot_h_m
    for seg_i in range(4):
        seg_w = bar_real_m / 4.0
        fc_seg = '#000000' if seg_i % 2 == 0 else '#FFFFFF'
        rect = patches.Rectangle((sb_x0 + seg_i * seg_w, sb_y0), seg_w, sb_h, linewidth=0.7, edgecolor='#000000', facecolor=fc_seg, zorder=8)
        ax.add_patch(rect)
    ax.text(sb_x0, sb_y0 - 0.015 * plot_h_m, "0", fontsize=6, ha='center', va='top', color='#000000')
    ax.text(sb_x0 + bar_real_m, sb_y0 - 0.015 * plot_h_m, f"{bar_real_m:g} m", fontsize=6, ha='center', va='top', color='#000000')
    ax.text(sb_x0 + bar_real_m / 2.0, sb_y0 + sb_h + 0.01 * plot_h_m, f"SKALA 1 : {scale_val:,}", fontsize=6.5, fontweight='bold', ha='center', va='bottom', color='#000000')

    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=300)
    plt.close(fig)
    buf.seek(0)
    return buf, scale_val

def render_cad_superelevasi_bw(r, draw_w_mm=330.0, draw_h_mm=277.0):
    e_des = r["e"] * 100.0
    en_val = r["en"] * 100.0
    ls_val = r["Ls_sup"]
    lc_val = r["Lc"]
    ltot_val = r["Ltot"]
    fc = r["jenis"].startswith("Full")

    if fc:
        s_start = -(2.0 / 3.0) * ls_val
        s_full1 = (1.0 / 3.0) * ls_val
        s_full2 = ltot_val - (1.0 / 3.0) * ls_val
        s_end = ltot_val + (2.0 / 3.0) * ls_val
    else:
        s_start = 0.0
        s_full1 = ls_val
        s_full2 = ltot_val - ls_val
        s_end = ltot_val

    world_w_m = (s_end - s_start) * 1.35
    s_h_calc = (world_w_m * 1000.0) / draw_w_mm
    std_scales = [100, 200, 250, 500, 750, 1000, 1250, 1500, 2000, 2500, 5000]
    scale_h = std_scales[-1]
    for sc in std_scales:
        if sc >= s_h_calc:
            scale_h = sc
            break

    fig_w_in = draw_w_mm / 25.4
    fig_h_in = draw_h_mm / 25.4
    fig, ax = plt.subplots(figsize=(fig_w_in, fig_h_in), dpi=300)
    fig.patch.set_facecolor('white')
    ax.set_facecolor('white')

    mid_s = 0.5 * (s_start + s_end)
    plot_w_m = (draw_w_mm / 1000.0) * scale_h
    x_min_lim = mid_s - plot_w_m / 2.0
    x_max_lim = mid_s + plot_w_m / 2.0
    y_lim = max(e_des + 4.5, 14.0)

    ax.set_xlim(x_min_lim, x_max_lim)
    ax.set_ylim(-y_lim, y_lim)
    ax.axis('off')

    # Sumbu Datar 0%
    ax.plot([x_min_lim, x_max_lim], [0, 0], color='#000000', linewidth=0.75, linestyle='-.')
    ax.text(x_min_lim + 0.02 * world_w_m, 0.4, "Sumbu Jalan (0%)", fontsize=6.5, color='#000000')

    # Profil Superelevasi
    x_luar = [s_start - 15, s_start, s_full1, s_full2, s_end, s_end + 15]
    y_luar = [-en_val, -en_val, e_des, e_des, -en_val, -en_val]
    ax.plot(x_luar, y_luar, color='#000000', linewidth=1.5)

    x_dalam = [s_start - 15, s_start, s_full1, s_full2, s_end, s_end + 15]
    y_dalam = [-en_val, -en_val, -e_des, -e_des, -en_val, -en_val]
    ax.plot(x_dalam, y_dalam, color='#000000', linewidth=1.3)

    mid_lc = 0.5 * (s_full1 + s_full2)
    ax.text(mid_lc, e_des + 0.035 * y_lim, f"e = {e_des:.2f} %", fontsize=7.5, fontweight='bold', ha='center', va='bottom', color='#000000')
    ax.text(mid_lc, -e_des - 0.035 * y_lim, f"-e = -{e_des:.2f} %", fontsize=7.5, fontweight='bold', ha='center', va='top', color='#000000')

    # Dimensi Panjang Lengkung Bawah
    y_dim_b2 = -y_lim * 0.72
    draw_dim_line(ax, (s_start, y_dim_b2), (s_full1, y_dim_b2), f"LS = {ls_val:.2f} m", text_offset=(0, 0.035 * y_lim), fontsize=7)
    if lc_val > 0:
        draw_dim_line(ax, (s_full1, y_dim_b2), (s_full2, y_dim_b2), f"LC = {lc_val:.2f} m", text_offset=(0, 0.035 * y_lim), fontsize=7)
    draw_dim_line(ax, (s_full2, y_dim_b2), (s_end, y_dim_b2), f"LS = {ls_val:.2f} m", text_offset=(0, 0.035 * y_lim), fontsize=7)

    # Skala Batang Grafis
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
        p.paragraph_format.line_spacing = 1.05
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
    draw_h_mm = inner_h_mm

    buf1, scale1 = render_cad_denah_bw(r, B, arah, draw_w_mm=draw_w_mm, draw_h_mm=draw_h_mm)
    buf2, scale2 = render_cad_superelevasi_bw(r, draw_w_mm=draw_w_mm, draw_h_mm=draw_h_mm)

    doc = Document()
    sheets_spec = [
        (buf1, "DENAH GEOMETRIK TIKUNGAN", f"1 : {scale1:,}", "01", "02"),
        (buf2, "DIAGRAM SUPERELEVASI", f"1 : {scale2:,}", "02", "02")
    ]

    for s_idx, (img_stream, jdl_gbr, skl_txt, no_lbr, jml_lbr) in enumerate(sheets_spec):
        section = doc.sections[0] if s_idx == 0 else doc.add_section()
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width = Mm(page_w_mm)
        section.page_height = Mm(page_h_mm)
        section.left_margin = Mm(margin_left_mm)
        section.right_margin = Mm(margin_other_mm)
        section.top_margin = Mm(margin_other_mm)
        section.bottom_margin = Mm(margin_other_mm)

        p_empty = section.header.paragraphs[0]
        p_empty.text = ""
        p_empty.paragraph_format.space_before = Pt(0)
        p_empty.paragraph_format.space_after = Pt(0)

        tbl = doc.add_table(rows=1, cols=2)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False

        row = tbl.rows[0]
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:trHeight {nsdecls("w")} w:val="{int(inner_h_mm * 56.7)}" w:hRule="exact"/>'))

        c_draw = row.cells[0]
        c_draw.width = Mm(draw_w_mm)
        c_etiket = row.cells[1]
        c_etiket.width = Mm(etiket_w_mm)

        for c_i, c_obj in enumerate([c_draw, c_etiket]):
            tcPr = c_obj._tc.get_or_add_tcPr()
            b_left = "14" if c_i == 0 else "0"
            b_right = "14"
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

        catatan_final = meta["catatan"] if meta["catatan"].strip() else (
            f"1. TIKUNGAN {r['jenis'].upper()}\n"
            f"2. Vr = {r['V']:.0f} km/jam; Rmin = {r['Rmin']:.2f} m\n"
            f"3. Rc = {r['Rc']:.2f} m; e = {r['e']*100:.2f}%\n"
            f"4. Ls = {r['Ls_sup']:.2f} m; Lc = {r['Lc']:.2f} m\n"
            f"5. Ts = {r['T']:.3f} m; Es = {r['E']:.3f} m\n"
            f"6. Ltotal = {r['Ltot']:.3f} m\n"
            f"7. V_safe = {r['v_safe']:.1f} km/jam"
        )

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

# ----------------------------------------------------------------------------
# 6. ENGINE LAPORAN EXCEL BER-FORMULA LIVE (17-STEP BM STANDAR)
# ----------------------------------------------------------------------------
def generate_excel_bm2021(r, pvi_data):
    wb = openpyxl.Workbook()
    font_title = Font(name="Calibri", size=13, bold=True, color="3A0056")
    font_sec = Font(name="Calibri", size=11, bold=True, color="880E4F")
    font_reg = Font(name="Calibri", size=10, color="2D3142")
    font_bold = Font(name="Calibri", size=10, bold=True, color="2D3142")
    font_italic = Font(name="Calibri", size=9, italic=True, color="555555")

    fill_cream = PatternFill(start_color="FFF3E0", end_color="FFF3E0", fill_type="solid")
    fill_yellow = PatternFill(start_color="FFF9C4", end_color="FFF9C4", fill_type="solid")
    fill_grey = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")

    thin = Side(style="thin", color="BDBDBD")
    med = Side(style="medium", color="424242")
    border_cell = Border(left=thin, right=thin, top=thin, bottom=thin)
    border_box = Border(left=med, right=med, top=med, bottom=med)
    align_c = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def apply_border_range(ws, cell_range, border):
        for row in ws[cell_range]:
            for cell in row:
                cell.border = border

    ws1 = wb.active
    ws1.title = "Tikungan Horisontal"
    col_widths = {"A": 6, "B": 12, "C": 12, "D": 11, "E": 11, "F": 11, "G": 5, "H": 12, "I": 12}
    for col_l, w in col_widths.items():
        ws1.column_dimensions[col_l].width = w

    ws1.merge_cells("A1:I1")
    ws1["A1"] = f"PERHITUNGAN TIKUNGAN 1 ({r['jenis'].upper()})"
    ws1["A1"].font = font_title
    ws1["A1"].fill = fill_cream
    ws1["A1"].alignment = align_c
    ws1.row_dimensions[1].height = 26
    apply_border_range(ws1, "A1:I1", border_box)

    inputs_h = [
        (2, "Kecepatan Rencana (Vr)", r["V"], '0.0" km/jam"', fill_yellow, "=D2/3.6", '0.0000" m/s"'),
        (3, "Superelevasi maksimum (e_max)", r["emax"], "0.0%", fill_yellow, "=D3", "0.000"),
        (4, "Δ (sudut tikungan)", r["delta"], '0.00"°"', None, "=RADIANS(D4)", '0.0000" rad"'),
        (5, "T (waktu tempuh asumsi)", 3.0, '0.0" s"', fill_yellow, "=D5", '0.0" s"'),
        (6, "B (lebar perkerasan)", B, '0.0" m"', fill_yellow, "=D6", '0.0" m"'),
        (7, "en (kemiringan normal)", r["en"], "0.0%", fill_grey, "=D7", "0.000"),
        (8, "C (perubahan percepatan sentrifugal)", 0.4, '0.0" m/s³"', fill_grey, "=D8", '0.0" m/s³"'),
        (9, "re (tingkat pencapaian kelandaian)", 0.035 if r["V"] <= 70 else 0.025, "0.0%", None, "=D9", "0.000"),
    ]
    for row_idx, label, val_in, fmt_in, fill_in, formula_h, fmt_h in inputs_h:
        ws1.row_dimensions[row_idx].height = 22
        ws1.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=3)
        ws1.merge_cells(start_row=row_idx, start_column=4, end_row=row_idx, end_column=6)
        ws1.merge_cells(start_row=row_idx, start_column=8, end_row=row_idx, end_column=9)
        c_lbl = ws1.cell(row=row_idx, column=1, value=label)
        c_lbl.font = font_reg
        c_lbl.alignment = align_c
        c_val = ws1.cell(row=row_idx, column=4, value=val_in)
        c_val.font = font_reg
        c_val.alignment = align_c
        c_val.number_format = fmt_in
        if fill_in:
            for c_i in range(4, 7):
                ws1.cell(row=row_idx, column=c_i).fill = fill_in
        c_eq = ws1.cell(row=row_idx, column=7, value="=")
        c_eq.font = font_bold
        c_eq.alignment = align_c
        c_res = ws1.cell(row=row_idx, column=8, value=formula_h)
        c_res.font = font_reg
        c_res.alignment = align_c
        c_res.number_format = fmt_h
        apply_border_range(ws1, f"A{row_idx}:I{row_idx}", border_cell)

    def add_step_block(ws, start_r, step_no, title_txt, rows_spec):
        end_r = start_r + len(rows_spec)
        ws.merge_cells(start_row=start_r, start_column=1, end_row=end_r, end_column=1)
        c_no = ws.cell(row=start_r, column=1, value=step_no)
        c_no.font = font_bold
        c_no.alignment = align_c
        ws.merge_cells(start_row=start_r, start_column=2, end_row=start_r, end_column=9)
        c_tit = ws.cell(row=start_r, column=2, value=title_txt)
        c_tit.font = font_sec
        c_tit.alignment = align_c
        ws.row_dimensions[start_r].height = 22
        for c_i in range(2, 10):
            ws.cell(row=start_r, column=c_i).fill = fill_cream
        for idx_sub, (sub_hdr, rumus_txt, formula_val, num_fmt, highlight_res) in enumerate(rows_spec):
            cur_r = start_r + 1 + idx_sub
            ws.row_dimensions[cur_r].height = 24
            if sub_hdr:
                ws.merge_cells(start_row=cur_r, start_column=2, end_row=cur_r, end_column=9)
                c_sh = ws.cell(row=cur_r, column=2, value=rumus_txt)
                c_sh.font = font_reg
                c_sh.alignment = align_c
            else:
                ws.merge_cells(start_row=cur_r, start_column=2, end_row=cur_r, end_column=6)
                c_rm = ws.cell(row=cur_r, column=2, value=rumus_txt)
                c_rm.font = font_italic
                c_rm.alignment = align_c
                c_eq = ws.cell(row=cur_r, column=7, value="=")
                c_eq.font = font_bold
                c_eq.alignment = align_c
                ws.merge_cells(start_row=cur_r, start_column=8, end_row=cur_r, end_column=9)
                c_vl = ws.cell(row=cur_r, column=8, value=formula_val)
                c_vl.font = font_bold if highlight_res else font_reg
                c_vl.alignment = align_c
                if num_fmt:
                    c_vl.number_format = num_fmt
                if highlight_res:
                    ws.cell(row=cur_r, column=8).fill = fill_grey
                    ws.cell(row=cur_r, column=9).fill = fill_grey
        apply_border_range(ws, f"A{start_r}:I{end_r}", border_box)
        for r_i in range(start_r, end_r + 1):
            for c_i in range(1, 10):
                cell = ws.cell(row=r_i, column=c_i)
                b_top = med if r_i == start_r else thin
                b_bot = med if r_i == end_r else thin
                b_left = med if c_i in (1, 2) else (thin if c_i in (7, 8) else None)
                b_right = med if c_i in (1, 9) else (thin if c_i in (6, 7) else None)
                cell.border = Border(top=b_top, bottom=b_bot, left=b_left, right=b_right)
        return end_r + 1

    curr = 10
    curr = add_step_block(ws1, curr, 1, "FMAX (Koefisien Gesek Melintang Maksimum)", [
        (False, "f_max = -0.00065·Vr + 0.192 (Vr < 80) atau -0.00125·Vr + 0.24", "=IF(D2<80, -0.00065*D2+0.192, -0.00125*D2+0.24)", "0.0000", False)
    ])
    curr = add_step_block(ws1, curr, 2, "MENGHITUNG Rmin", [
        (False, "Rmin = Vr² / (127 × (e_max + f_max))", "=D2^2/(127*(H3+H11))", '0.0000" m"', False)
    ])
    curr = add_step_block(ws1, curr, 3, "MENENTUKAN JARI-JARI RENCANA (Rc)", [
        (False, "Rc rencana ≥ Rmin, diambil Rc rencana", r["Rc"], '0.00" m"', True)
    ])
    curr = add_step_block(ws1, curr, 4, "MENGHITUNG SUPERELEVASI DESAIN (e)", [
        (False, "Dmax = 181913.53 × (e_max + f_max) / Vr²", "=181913.53*(H3+H11)/(D2^2)", "0.0000", False),
        (False, "D = 1432.39 / Rc", "=1432.39/H15", "0.0000", False),
        (False, "e = MAX(en, -e_max·(D/Dmax)² + 2·e_max·(D/Dmax))", "=MAX(H7, -H3*(H18/H17)^2 + 2*H3*(H18/H17))", "0.00%", False),
    ])
    curr = add_step_block(ws1, curr, 5, "MENGHITUNG Ls min (Lengkung Peralihan Minimum)", [
        (True, "Berdasarkan waktu tempuh maksimum lengkung peralihan (3 detik)", None, None, False),
        (False, "Ls = (Vr / 3.6) × T", "=(D2/3.6)*D5", "0.0000", False),
        (True, "Berdasarkan antisipasi gaya sentrifugal (Rumus Shortt)", None, None, False),
        (False, "Ls = (0.022·Vr³)/(Rc·C) - (2.727·Vr·e)/C", "=(0.022*(D2^3))/(H15*D8)-(2.727*D2*H19)/D8", "0.0000", False),
        (True, "Berdasarkan tingkat pencapaian perubahan kelandaian", None, None, False),
        (False, "Ls = ((e + en) × Vr) / (3.6 × re)", "=((H19+H7)*D2)/(3.6*H9)", "0.0000", False),
    ])
    curr = add_step_block(ws1, curr, 6, "MENENTUKAN Ls RENCANA (Dibulatkan kelipatan 5 m)", [
        (False, "Ls rencana ≥ MAX(Ls1, Ls2, Ls3), diambil Ls rencana", "=CEILING(MAX(H22,H24,H26), 5)", '0.00" m"', True)
    ])
    curr = add_step_block(ws1, curr, 7, "Menghitung θs (deflection angle tiap spiral)", [
        (False, "θs = (Ls × 90) / (π × Rc)", "=(H28*90)/(PI()*H15)", '0.0000"°"', False)
    ])
    curr = add_step_block(ws1, curr, 8, "Menghitung θc (sudut lingkaran / central angle busur)", [
        (False, "θc = Δ - 2·θs", "=IF(H28=0, D4, D4-2*H30)", '0.0000"°"', False)
    ])
    curr = add_step_block(ws1, curr, 9, "Menghitung Panjang Busur Lingkaran (Lc)", [
        (False, "Lc = (θc / 360) × 2π·Rc", "=(H32/360)*2*PI()*H15", '0.0000" m"', False)
    ])
    curr = add_step_block(ws1, curr, 10, "Menentukan L total", [
        (False, "Ltotal = Lc + 2·Ls", "=H34+2*H28", '0.0000" m"', False)
    ])
    curr = add_step_block(ws1, curr, 11, "Menentukan Xs", [
        (False, "Xs = Ls × (1 - Ls² / (40·Rc²))", "=H28*(1-(H28^2)/(40*(H15^2)))", '0.0000" m"', False)
    ])
    curr = add_step_block(ws1, curr, 12, "Menentukan Ys", [
        (False, "Ys = Ls² / (6·Rc)", "=(H28^2)/(6*H15)", '0.0000" m"', False)
    ])
    curr = add_step_block(ws1, curr, 13, "Menentukan P (Pergeseran Tangen)", [
        (False, "P = Ys - Rc × (1 - cos(θs))", "=H40-H15*(1-COS(RADIANS(H30)))", '0.0000" m"', False)
    ])
    curr = add_step_block(ws1, curr, 14, "Menentukan K", [
        (False, "K = Xs - Rc × sin(θs)", "=H38-H15*SIN(RADIANS(H30))", '0.0000" m"', False)
    ])
    curr = add_step_block(ws1, curr, 15, "Menentukan Es (Jarak Eksternal)", [
        (False, "Es = (Rc + P) / cos(Δ/2) - Rc", "=(H15+H42)/COS(RADIANS(D4/2))-H15", '0.0000" m"', False)
    ])
    curr = add_step_block(ws1, curr, 16, "Menentukan Ts (Panjang Tangen)", [
        (False, "Ts = (Rc + P) × tan(Δ/2) + K", "=(H15+H42)*TAN(RADIANS(D4/2))+H44", '0.0000" m"', False)
    ])
    curr = add_step_block(ws1, curr, 17, "Kontrol Kecepatan Aman & Stasioning", [
        (False, "V_aman = √(127 × Rc × (e + f_max))", "=SQRT(127*H15*(H19+H11))", '0.00" km/jam"', False),
        (False, "STA PI (Input Desain)", r["sta_pi"], '0.00" m"', True),
        (False, "STA Awal Tikungan (TS / TC) = STA PI - Ts", "=H51-H48", '0.00" m"', False),
        (False, "STA Akhir Tikungan (ST / CT) = STA Awal + Ltotal", "=H52+H36", '0.00" m"', False),
    ])

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ----------------------------------------------------------------------------
# 7. DASHBOARD UTAMA DENGAN HEADER & KPI CARDS BERWARNA CERAH
# ----------------------------------------------------------------------------
v_safe = r["v_safe"]
is_safe = v_safe >= V
safety_diff = v_safe - V

# Header Bar Cerah
st.markdown(f"""
<div class="hero-bright">
    <div>
        <h1>🌸 Sistem Geometri Jalan Interaktif</h1>
        <p>Perencanaan Geometrik Jalan Raya & Simulasi Trase Peta Realistis (TPGJAK 1997 / RAE 2021)</p>
    </div>
    <div style="display: flex; gap: 8px; flex-wrap: wrap;">
        <span class="pill-tag pill-tag-active">✨ {r['jenis']}</span>
        <span class="pill-tag">Vr: {V} km/h</span>
        <span class="pill-tag">Δ: {delta:.1f}°</span>
        <span class="pill-tag" style="color: {'#2E7D32' if is_safe else '#C2185B'};">
            {'🛡️ Aman' if is_safe else '⚠️ Evaluasi'} (V_safe: {v_safe:.1f} km/h)
        </span>
    </div>
</div>
""", unsafe_allow_html=True)

# 6 Kartu KPI 3D Cerah
st.markdown(f"""
<div class="kpi-grid">
    <div class="kpi-card-3d">
        <div class="kpi-title">Jenis Tikungan</div>
        <div class="kpi-val" style="font-size: 19px; color: #D81B60;">{r['jenis'].split()[0]}</div>
        <div class="kpi-sub">{r['jenis']}</div>
    </div>
    <div class="kpi-card-3d">
        <div class="kpi-title">Kecepatan Aman</div>
        <div class="kpi-val" style="color: #00897B;">{v_safe:.1f} <span style="font-size: 13px;">km/h</span></div>
        <div class="kpi-sub" style="color: #00897B;">Margin: {'+' if safety_diff>=0 else ''}{safety_diff:.1f} km/h</div>
    </div>
    <div class="kpi-card-3d">
        <div class="kpi-title">Jari-Jari (Rc)</div>
        <div class="kpi-val" style="color: #7B1FA2;">{Rc:.0f} <span style="font-size: 13px;">m</span></div>
        <div class="kpi-sub">Rmin = {Rmin:.1f} m ({(Rc/Rmin):.2f}x)</div>
    </div>
    <div class="kpi-card-3d">
        <div class="kpi-title">Superelevasi (e)</div>
        <div class="kpi-val" style="color: #E64A19;">{r['e']*100:.2f} <span style="font-size: 13px;">%</span></div>
        <div class="kpi-sub">e_maks = {emax*100:.1f}%</div>
    </div>
    <div class="kpi-card-3d">
        <div class="kpi-title">Panjang Tikungan</div>
        <div class="kpi-val" style="color: #F57F17;">{r['Ltot']:.1f} <span style="font-size: 13px;">m</span></div>
        <div class="kpi-sub">Ls = {r['Ls_sup']:.0f}m | Lc = {r['Lc']:.1f}m</div>
    </div>
    <div class="kpi-card-3d">
        <div class="kpi-title">Rentang STA</div>
        <div class="kpi-val" style="font-size: 16px; color: #1565C0;">{sta_str(r['sta_awal'], d=0)}</div>
        <div class="kpi-sub">s/d {sta_str(r['sta_akhir'], d=0)}</div>
    </div>
</div>
""", unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# 8. TABS UTAMA APLIKASI (PROFIL MELINTANG DIHAPUS SESUAI PERMINTAAN USER)
# ----------------------------------------------------------------------------
tab_map, tab_cad2d, tab_sup, tab_prof, tab_doc, tab_xls = st.tabs([
    "🗺️ Peta & Simulasi Realistis (Mobil Jalan)",
    "📍 Denah Teknis 2D (CAD)",
    "📈 Diagram Superelevasi",
    "⛰️ Profil Vertikal & Galian/Timbunan",
    "📄 Gambar Kerja Word (A3 + Etiket)",
    "📊 Tabel Hitungan & Excel 17-Rumus"
])

# ----------------------------------------------------------------------------
# TAB 1: PETA & SIMULASI REALISTIS (JALAN TIKUNGAN NYATA DENGAN MOBIL LEWAT)
# ----------------------------------------------------------------------------
with tab_map:
    st.markdown("""
    <div style="background: #FFFFFF; border-radius: 16px; border: 1.5px solid #F8BBD0; padding: 16px 20px; margin-bottom: 18px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap;">
        <div>
            <h4 style="margin: 0; color: #4A154B; font-size: 16px;">🛣️ Peta Trase Geometrik Nyata & Simulator Kendaraan Berbelok</h4>
            <p style="margin: 3px 0 0 0; color: #8E24AA; font-size: 12.5px;">Aspal gelap dengan marka jalan putih, bahu jalan kerikil, pemandangan pohon hijau, dan sorot lampu mobil.</p>
        </div>
        <div style="font-size: 13px; font-weight: 700; color: #C2185B;">
            ▶️ Klik tombol <b>'Mulai Jalan'</b> di bawah grafik untuk memutar animasi!
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Koordinat batas perkerasan jalan
    xl, yl = x_horiz + (B / 2) * nx, y_horiz + (B / 2) * ny
    xr, yr = x_horiz - (B / 2) * nx, y_horiz - (B / 2) * ny
    
    # Koordinat bahu jalan
    x_sh_l, y_sh_l = x_horiz + (B / 2 + Bahu) * nx, y_horiz + (B / 2 + Bahu) * ny
    x_sh_r, y_sh_r = x_horiz - (B / 2 + Bahu) * nx, y_horiz - (B / 2 + Bahu) * ny

    # Koordinat Marka Jalan Putus-putus (Dashed Centerline Strips)
    dash_step = 6.0
    dash_len = 3.0
    dash_x, dash_y = [], []
    for s_curr in np.arange(s_vals[0], s_vals[-1], dash_step):
        s_end_d = min(s_curr + dash_len, s_vals[-1])
        idx_c1 = np.argmin(np.abs(s_vals - s_curr))
        idx_c2 = np.argmin(np.abs(s_vals - s_end_d))
        dash_x.extend([x_horiz[idx_c1], x_horiz[idx_c2], None])
        dash_y.extend([y_horiz[idx_c1], y_horiz[idx_c2], None])

    # Pepohonan / Vegetasi di Pinggir Jalan (Realistic Landscape)
    tree_offset = B / 2 + Bahu + 4.5
    tree_step = 25.0
    tree_x, tree_y = [], []
    for s_t in np.arange(s_vals[0] + 10, s_vals[-1] - 10, tree_step):
        idx_t = np.argmin(np.abs(s_vals - s_t))
        tree_x.append(x_horiz[idx_t] + tree_offset * nx[idx_t])
        tree_y.append(y_horiz[idx_t] + tree_offset * ny[idx_t])
        tree_x.append(x_horiz[idx_t] - tree_offset * nx[idx_t])
        tree_y.append(y_horiz[idx_t] - tree_offset * ny[idx_t])

    # Bangun Frame Animasi Kendaraan Nyata
    step_anim = max(1, len(s_vals) // 55)
    anim_indices = list(range(0, len(s_vals), step_anim))
    if anim_indices[-1] != len(s_vals) - 1:
        anim_indices.append(len(s_vals) - 1)
    
    anim_dur = max(35, int(4200 / max(v_sim, 10)))
    frames = []
    slider_steps = []

    for f_idx in anim_indices:
        xc_now = x_horiz[f_idx]
        yc_now = y_horiz[f_idx]
        th_now = ths_[f_idx]
        s_now = s_vals[f_idx]

        bx, by, cx, cy, beam_x, beam_y = get_car_geometry(xc_now, yc_now, th_now)
        
        cur_sta_str = sta_str(r['sta_awal'] + s_now, d=0)

        frame_data = [
            # Sorot Lampu Depan (Headlights)
            go.Scatter(
                x=beam_x, y=beam_y, fill='toself',
                fillcolor='rgba(255, 235, 59, 0.35)',
                line=dict(color='rgba(255, 214, 0, 0.6)', width=1),
                mode='lines', showlegend=False, hoverinfo='skip'
            ),
            # Bodi Mobil (Sports Red / Fuchsia)
            go.Scatter(
                x=bx, y=by, fill='toself',
                fillcolor='#E91E63',
                line=dict(color='#880E4F', width=2),
                mode='lines', showlegend=False, hoverinfo='skip'
            ),
            # Kaca Kabin Mobil
            go.Scatter(
                x=cx, y=cy, fill='toself',
                fillcolor='#212121',
                line=dict(color='#000000', width=1.2),
                mode='lines', showlegend=False, hoverinfo='skip'
            ),
            # Label STA Mobil
            go.Scatter(
                x=[xc_now], y=[yc_now],
                mode='text', text=[f"<b>🚘 {cur_sta_str}</b>"],
                textposition="top center",
                textfont=dict(color="#4A154B", size=12),
                showlegend=False, hoverinfo='skip'
            )
        ]

        frames.append(go.Frame(data=frame_data, traces=[5, 6, 7, 8], name=str(f_idx)))
        slider_steps.append({
            "args": [[str(f_idx)], {"frame": {"duration": 0, "redraw": True}, "mode": "immediate"}],
            "label": f"{s_now:.0f}m", "method": "animate"
        })

    # Inisialisasi Mobil Awal
    bx0, by0, cx0, cy0, beam_x0, beam_y0 = get_car_geometry(x_horiz[0], y_horiz[0], ths_[0])

    fig_map = go.Figure()

    # 1. Bahu Jalan (Shoulder)
    fig_map.add_trace(go.Scatter(
        x=np.concatenate([x_sh_l, x_sh_r[::-1]]),
        y=np.concatenate([y_sh_l, y_sh_r[::-1]]),
        fill='toself', fillcolor='#CFD8DC',
        line=dict(color='rgba(0,0,0,0)'),
        name='Bahu Jalan Kerikil', hoverinfo='skip'
    ))

    # 2. Aspal Perkerasan Jalan Gelap Realistis
    fig_map.add_trace(go.Scatter(
        x=np.concatenate([xl, xr[::-1]]),
        y=np.concatenate([yl, yr[::-1]]),
        fill='toself', fillcolor='#2E3440',
        line=dict(color='#FFFFFF', width=2),
        name='Aspal Jalan Raya', hoverinfo='skip'
    ))

    # 3. Marka Tengah Putus-putus (Dashed Line)
    fig_map.add_trace(go.Scatter(
        x=dash_x, y=dash_y, mode='lines',
        line=dict(color='#FFFFFF', width=2),
        name='Marka Tengah Jalan', hoverinfo='skip'
    ))

    # 4. Pepohonan Pinggir Jalan (Realistic Landscape)
    fig_map.add_trace(go.Scatter(
        x=tree_x, y=tree_y, mode='markers',
        marker=dict(size=14, color='#43A047', symbol='circle', line=dict(color='#1B5E20', width=1.5)),
        name='Vegetasi / Pohon', hoverinfo='skip'
    ))

    # 5. Patok Stasioning Tiap 50m
    sta_min_p = math.ceil((r["sta_awal"] + s_vals[0]) / 50.0) * 50.0
    sta_max_p = math.floor((r["sta_awal"] + s_vals[-1]) / 50.0) * 50.0
    post_stas = np.arange(sta_min_p, sta_max_p + 1, 50.0)
    post_x, post_y, post_lbl = [], [], []
    for st_v in post_stas:
        s_rel = st_v - r["sta_awal"]
        if s_vals[0] <= s_rel <= s_vals[-1]:
            idx_p = np.argmin(np.abs(s_vals - s_rel))
            post_x.append(x_horiz[idx_p] + (B/2 + Bahu + 1.2) * nx[idx_p])
            post_y.append(y_horiz[idx_p] + (B/2 + Bahu + 1.2) * ny[idx_p])
            post_lbl.append(sta_str(st_v, d=0))

    fig_map.add_trace(go.Scatter(
        x=post_x, y=post_y, mode='markers+text',
        marker=dict(size=7, color='#E53935', symbol='square'),
        text=post_lbl, textposition="top right", textfont=dict(color="#37474F", size=9),
        name='Patok STA (Tiap 50m)'
    ))

    # 6. Sorot Lampu Mobil (Trace index 5)
    fig_map.add_trace(go.Scatter(
        x=beam_x0, y=beam_y0, fill='toself', fillcolor='rgba(255, 235, 59, 0.35)',
        line=dict(color='rgba(255, 214, 0, 0.6)', width=1), mode='lines', showlegend=False, hoverinfo='skip'
    ))

    # 7. Bodi Mobil Nyata (Trace index 6)
    fig_map.add_trace(go.Scatter(
        x=bx0, y=by0, fill='toself', fillcolor='#E91E63',
        line=dict(color='#880E4F', width=2), mode='lines', name='Mobil Nyata', hoverinfo='skip'
    ))

    # 8. Kaca Kabin (Trace index 7)
    fig_map.add_trace(go.Scatter(
        x=cx0, y=cy0, fill='toself', fillcolor='#212121',
        line=dict(color='#000000', width=1.2), mode='lines', showlegend=False, hoverinfo='skip'
    ))

    # 9. Teks Posisi Mobil (Trace index 8)
    fig_map.add_trace(go.Scatter(
        x=[x_horiz[0]], y=[y_horiz[0]], mode='text',
        text=[f"<b>🚘 {sta_str(r['sta_awal'] + s_vals[0], d=0)}</b>"],
        textposition="top center", textfont=dict(color="#4A154B", size=12),
        showlegend=False, hoverinfo='skip'
    ))

    fig_map.frames = frames

    fig_map.update_layout(
        title=dict(text=f"Simulasi Trase Peta Realistis - Tikungan {r['jenis']} (Arah {arah})", font=dict(color='#4A154B', size=17), x=0.5),
        xaxis=dict(title="Koordinat X (m)", gridcolor='rgba(200, 230, 201, 0.4)', showline=True, linecolor='#81C784', zeroline=False),
        yaxis=dict(title="Koordinat Y (m)", gridcolor='rgba(200, 230, 201, 0.4)', showline=True, linecolor='#81C784', scaleanchor="x", scaleratio=1, zeroline=False),
        plot_bgcolor="#E8F5E9",
        paper_bgcolor="#FFFFFF",
        height=620,
        margin=dict(l=40, r=40, t=50, b=50),
        legend=dict(bgcolor="rgba(255,255,255,0.92)", bordercolor="#F8BBD0", borderwidth=1, font=dict(color="#4A154B", size=11), x=0.01, y=0.99),
        updatemenus=[{
            "type": "buttons", "showactive": True, "x": 0.02, "y": 0.03, "xanchor": "left", "yanchor": "bottom",
            "bgcolor": "#FFF0F5", "bordercolor": "#EC407A", "font": {"color": "#880E4F", "size": 13, "weight": "bold"},
            "buttons": [
                {"label": "▶️ Mulai Jalan", "method": "animate", "args": [None, {"frame": {"duration": anim_dur, "redraw": True}, "fromcurrent": True, "transition": {"duration": 0}}]},
                {"label": "⏸️ Pause", "method": "animate", "args": [[None], {"frame": {"duration": 0, "redraw": False}, "mode": "immediate", "transition": {"duration": 0}}]}
            ]
        }],
        sliders=[{
            "active": 0, "yanchor": "top", "xanchor": "left",
            "currentvalue": {"font": {"size": 13, "color": "#4A154B"}, "prefix": "Posisi Relatif: ", "visible": True, "xanchor": "right"},
            "pad": {"b": 10, "t": 40}, "len": 0.78, "x": 0.19, "y": 0.03, "steps": slider_steps
        }]
    )
    st.plotly_chart(fig_map, use_container_width=True)

# ----------------------------------------------------------------------------
# TAB 2: DENAH TEKNIS 2D (CAD TEKNIS RESMI)
# ----------------------------------------------------------------------------
with tab_cad2d:
    col_d1, col_d2 = st.columns([3.2, 1])
    with col_d2:
        st.markdown("""
        <div style="background: #FFFFFF; border-radius: 16px; border: 1.5px solid #F8BBD0; padding: 18px; box-shadow: 0 6px 18px rgba(233, 30, 99, 0.06);">
            <h4 style="color: #4A154B; margin: 0 0 10px 0; font-size: 15px;">⚙️ Layer Gambar CAD</h4>
        """, unsafe_allow_html=True)
        show_pave = st.checkbox("Badan Perkerasan Aspal", True)
        show_shou = st.checkbox("Bahu Jalan", True)
        show_tang = st.checkbox("Garis Tangen & PI", True)
        show_rad = st.checkbox("Jari-jari Rc & Pusat O", True)
        show_tick = st.checkbox("Garis Ticks STA Tiap 50m", True)
        st.markdown("</div>", unsafe_allow_html=True)

    with col_d1:
        fig_2d = go.Figure()
        if show_shou:
            fig_2d.add_trace(go.Scatter(
                x=np.concatenate([x_sh_l, x_sh_r[::-1]]),
                y=np.concatenate([y_sh_l, y_sh_r[::-1]]),
                fill='toself', fillcolor='rgba(244, 143, 177, 0.18)',
                line=dict(color='rgba(194, 24, 91, 0.35)', dash='dot', width=1),
                name='Bahu Jalan', hoverinfo='skip'
            ))

        if show_pave:
            fig_2d.add_trace(go.Scatter(
                x=np.concatenate([xl, xr[::-1]]),
                y=np.concatenate([yl, yr[::-1]]),
                fill='toself', fillcolor='rgba(245, 238, 250, 0.7)',
                line=dict(color='#AB47BC', width=1.5), name='Badan Perkerasan', hoverinfo='skip'
            ))

        fig_2d.add_trace(go.Scatter(
            x=x_horiz, y=y_horiz, mode='lines', name='As Jalan (Centerline)',
            line=dict(color='#4A154B', width=3.5),
            hovertemplate='<b>STA:</b> %{customdata}<br><b>X:</b> %{x:.2f} m<br><b>Y:</b> %{y:.2f} m<extra></extra>',
            customdata=[sta_str(r['sta_awal'] + sv) for sv in s_vals]
        ))

        idx_start = np.argmin(np.abs(s_vals - 0))
        idx_end = np.argmin(np.abs(s_vals - r["Ltot"]))
        x_ts, y_ts = x_horiz[idx_start], y_horiz[idx_start]
        x_st, y_st = x_horiz[idx_end], y_horiz[idx_end]
        x_pi = x_ts + r["T"] * np.cos(ths_[idx_start])
        y_pi = y_ts + r["T"] * np.sin(ths_[idx_start])

        if show_tang:
            fig_2d.add_trace(go.Scatter(
                x=[x_ts, x_pi, x_st], y=[y_ts, y_pi, y_st],
                mode='lines', name='Garis Tangen', line=dict(color='#0284C7', width=1.8, dash='dash')
            ))
            fig_2d.add_trace(go.Scatter(
                x=[x_pi], y=[y_pi], mode='markers', name='Titik PI',
                marker=dict(size=11, color='#0284C7', symbol='diamond', line=dict(width=1.5, color='#FFFFFF'))
            ))
            fig_2d.add_annotation(
                x=x_pi, y=y_pi,
                text=f"<b>PI</b><br>Δ = {r['delta']:.1f}°<br>STA {sta_str(r['sta_pi'])}",
                showarrow=False, xshift=12, yshift=32 if arah == "Kanan" else -32, xanchor="left",
                font=dict(color="#0284C7", size=11), bgcolor="rgba(255,255,255,0.92)",
                bordercolor="#BAE6FD", borderwidth=1, borderpad=4
            )

        mid_idx = idx_start if r["jenis"].startswith("Full") else np.argmin(np.abs(s_vals - (r["Ltot"] / 2)))
        x_center = x_horiz[mid_idx] + sg * r["Rc"] * nx[mid_idx]
        y_center = y_horiz[mid_idx] + sg * r["Rc"] * ny[mid_idx]

        if show_rad:
            fig_2d.add_trace(go.Scatter(
                x=[x_horiz[mid_idx], x_center], y=[y_horiz[mid_idx], y_center],
                mode='lines', name=f'Jari-jari Rc = {r["Rc"]:.0f} m',
                line=dict(color='#00897B', width=1.5, dash='dot')
            ))
            fig_2d.add_trace(go.Scatter(
                x=[x_center], y=[y_center], mode='markers', name='Pusat Lengkung (O)',
                marker=dict(size=12, color='#00897B', symbol='cross-thin', line=dict(width=2.5, color='#00897B'))
            ))

        if show_tick:
            for st_val in post_stas:
                s_rel = st_val - r["sta_awal"]
                if s_vals[0] <= s_rel <= s_vals[-1]:
                    idx_t = np.argmin(np.abs(s_vals - s_rel))
                    xt0, yt0 = x_horiz[idx_t], y_horiz[idx_t]
                    nxt, nyt = nx[idx_t], ny[idx_t]
                    t_half = max(B * 0.85, 5.5)
                    fig_2d.add_trace(go.Scatter(
                        x=[xt0 - t_half * nxt, xt0 + t_half * nxt],
                        y=[yt0 - t_half * nyt, yt0 + t_half * nyt],
                        mode='lines', showlegend=False, hoverinfo='skip', line=dict(color='#9E9E9E', width=1)
                    ))
                    fig_2d.add_annotation(
                        x=xt0 - sg * (t_half + 2.5) * nxt,
                        y=yt0 - sg * (t_half + 2.5) * nyt,
                        text=sta_str(st_val, d=0), showarrow=False,
                        font=dict(color="#4A5568", size=9.5), bgcolor="rgba(255,255,255,0.85)", borderpad=1
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
            x=pts_x, y=pts_y, mode='markers', name='Titik Utama Lengkung',
            marker=dict(size=10, color='#EC407A', symbol='circle', line=dict(width=2, color='#FFFFFF'))
        ))
        for px, py, plbl in zip(pts_x, pts_y, pts_lbls):
            fig_2d.add_annotation(
                x=px, y=py, text=plbl, showarrow=False,
                yshift=-26 if arah == "Kanan" else 26,
                font=dict(color="#880E4F", size=10.5), bgcolor="rgba(255,255,255,0.92)",
                bordercolor="#F48FB1", borderwidth=1, borderpad=3
            )

        fig_2d.update_layout(
            title=dict(text=f"Denah Trase Geometrik - {r['jenis']}", font=dict(color='#4A154B', size=16), x=0.5),
            xaxis=dict(title="X (m)", gridcolor='rgba(244, 143, 177, 0.25)', showline=True, linecolor='#AB47BC', mirror=True),
            yaxis=dict(title="Y (m)", gridcolor='rgba(244, 143, 177, 0.25)', showline=True, linecolor='#AB47BC', mirror=True, scaleanchor="x", scaleratio=1),
            legend=dict(bgcolor="rgba(255, 255, 255, 0.92)", bordercolor="#F48FB1", borderwidth=1, font=dict(color="#4A154B", size=11), x=0.99, y=0.99, xanchor="right", yanchor="top"),
            plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF", margin=dict(l=50, r=30, t=50, b=40), height=580
        )
        st.plotly_chart(fig_2d, use_container_width=True)

# ----------------------------------------------------------------------------
# TAB 3: DIAGRAM SUPERELEVASI INTERAKTIF
# ----------------------------------------------------------------------------
with tab_sup:
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
    fig_sup.add_trace(go.Scatter(x=sta_nodes, y=e_luar, mode='lines', name='Tepi Luar (+e)', line=dict(color='#EC407A', width=3), fill='tozeroy', fillcolor='rgba(236, 64, 122, 0.08)'))
    fig_sup.add_trace(go.Scatter(x=sta_nodes, y=e_dalam, mode='lines', name='Tepi Dalam (-e)', line=dict(color='#7B1FA2', width=3), fill='tozeroy', fillcolor='rgba(123, 31, 162, 0.06)'))
    fig_sup.add_hline(y=0, line_width=1.2, line_color="#424242")
    fig_sup.add_trace(go.Scatter(x=[sta_0 + s_zero_1, sta_0 + s_zero_2], y=[0.0, 0.0], mode='markers', name='Kemiringan Datar 0%', marker=dict(size=8, color='#4A154B', symbol='circle')))
    
    y_top_limit = max(e_desain + 3.5, 13.0)
    y_bot_limit = min(-e_desain - 3.5, -13.0)
    for ks, klbl in zip(krit_s, krit_lbl):
        sta_k = sta_0 + ks
        fig_sup.add_vline(x=sta_k, line_width=1, line_dash="dot", line_color="#F48FB1")
        fig_sup.add_annotation(x=sta_k, y=y_top_limit - 1.2, text=f"<b>{klbl}</b><br>{sta_str(sta_k)}", showarrow=False, font=dict(color="#4A154B", size=10), bgcolor="rgba(255,255,255,0.92)", bordercolor="#F48FB1", borderwidth=1, borderpad=3)
        
    mid_full_sta = sta_0 + (s_full_start + s_full_end) / 2
    fig_sup.add_annotation(x=mid_full_sta, y=e_desain + 0.9, text=f"<b>+e = {e_desain:.2f}%</b>", showarrow=False, font=dict(color="#D81B60", size=11.5), bgcolor="rgba(255,255,255,0.92)", borderpad=2)
    fig_sup.add_annotation(x=mid_full_sta, y=-e_desain - 1.1, text=f"<b>-e = -{e_desain:.2f}%</b>", showarrow=False, font=dict(color="#7B1FA2", size=11.5), bgcolor="rgba(255,255,255,0.92)", borderpad=2)
    
    tick_step = 50.0
    x_tick_vals = np.arange(math.ceil(sta_nodes[0] / tick_step) * tick_step, math.floor(sta_nodes[-1] / tick_step) * tick_step + 1, tick_step)
    x_tick_texts = [sta_str(v, d=0) for v in x_tick_vals]
    fig_sup.update_layout(
        title=dict(text="Diagram Superelevasi (Rotasi Sumbu Jalan)", font=dict(color='#4A154B', size=16), x=0.5),
        xaxis=dict(title="Stasioning (STA)", tickmode='array', tickvals=x_tick_vals, ticktext=x_tick_texts, gridcolor='rgba(244, 143, 177, 0.25)', showline=True, linecolor='#AB47BC', mirror=True),
        yaxis=dict(title="Kemiringan Melintang (%)", range=[y_bot_limit, y_top_limit], gridcolor='rgba(244, 143, 177, 0.25)', showline=True, linecolor='#AB47BC', mirror=True),
        legend=dict(orientation="h", yanchor="bottom", y=0.02, xanchor="center", x=0.5, bgcolor="rgba(255, 255, 255, 0.92)", bordercolor="#F48FB1", borderwidth=1, font=dict(color="#4A154B", size=11)),
        plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF", margin=dict(l=50, r=30, t=50, b=40), height=530
    )
    st.plotly_chart(fig_sup, use_container_width=True)

# ----------------------------------------------------------------------------
# TAB 4: PROFIL VERTIKAL & GALIAN / TIMBUNAN
# ----------------------------------------------------------------------------
with tab_prof:
    col_pv1, col_pv2 = st.columns([1, 1])
    with col_pv1:
        st.markdown("<p style='font-weight: 700; color: #7B1FA2;'>Tabel Titik Perpotongan Vertikal (PPV):</p>", unsafe_allow_html=True)
        st.session_state.pvi_df = st.data_editor(st.session_state.pvi_df, num_rows="dynamic", use_container_width=True)
    with col_pv2:
        st.markdown("<p style='font-weight: 700; color: #7B1FA2;'>Tabel Elevasi Tanah Asli:</p>", unsafe_allow_html=True)
        st.session_state.tanah_df = st.data_editor(st.session_state.tanah_df, num_rows="dynamic", use_container_width=True)

    P = hitung_profil(st.session_state.pvi_df, V)
    if P is not None:
        fig_prof = go.Figure()
        t_df = st.session_state.tanah_df.dropna().sort_values(st.session_state.tanah_df.columns[0])
        fig_prof.add_trace(go.Scatter(x=t_df.iloc[:, 0], y=t_df.iloc[:, 1], mode='lines+markers', name='Elevasi Tanah Asli', line=dict(color='#8D6E63', width=2, dash='dot')))
        fig_prof.add_trace(go.Scatter(x=P["x"], y=P["y"], mode='lines', name='Elevasi Rencana Jalan', line=dict(color='#D81B60', width=3)))
        
        fig_prof.add_trace(go.Scatter(x=P["sta"], y=P["el"], mode='markers', name='Titik PPV', marker=dict(size=9, color='#4A154B', symbol='diamond')))
        for k in P["kurva"]:
            fig_prof.add_annotation(
                x=k["sta"], y=k["el"] + 1.2,
                text=f"<b>{k['jenis']}</b><br>Lv={k['Lv']:.0f}m",
                showarrow=False, font=dict(color="#7B1FA2", size=10), bgcolor="rgba(255,255,255,0.9)", borderpad=2
            )

        fig_prof.update_layout(
            title=dict(text="Profil Memanjang Trase Jalan (Kelandaian & Lengkung Vertikal)", font=dict(color='#4A154B', size=16), x=0.5),
            xaxis=dict(title="Stasioning (STA)", gridcolor='rgba(244, 143, 177, 0.25)'),
            yaxis=dict(title="Elevasi (m)", gridcolor='rgba(244, 143, 177, 0.25)'),
            plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF", height=480,
            margin=dict(l=40, r=30, t=50, b=40)
        )
        st.plotly_chart(fig_prof, use_container_width=True)

        xs = np.arange(math.ceil(P["x"][0] / 50) * 50, P["x"][-1] + 0.1, 50)
        rn = np.interp(xs, P["x"], P["y"])
        ta = np.interp(xs, t_df.iloc[:, 0], t_df.iloc[:, 1])
        diff_h = rn - ta
        df_cutfill = pd.DataFrame({
            "STA": [sta_str(v, 0) for v in xs],
            "Elevasi Rencana (m)": rn.round(2),
            "Elevasi Tanah (m)": ta.round(2),
            "Selisih H (m)": diff_h.round(2),
            "Keterangan": np.where(diff_h > 0, "Timbunan (Fill)", "Galian (Cut)")
        })
        st.dataframe(df_cutfill, hide_index=True, use_container_width=True)
    else:
        st.info("💡 Masukkan minimal 2 titik PPV untuk menampilkan profil vertikal.")

# ----------------------------------------------------------------------------
# TAB 5: GAMBAR KERJA WORD (A3 + ETIKET KANAN 60MM)
# ----------------------------------------------------------------------------
with tab_doc:
    st.markdown("<h3 style='color: #4A154B;'>📐 Gambar Kerja Word (.DOCX) Skalatis Hitam Putih + Etiket Kanan</h3>", unsafe_allow_html=True)
    st.info("📄 File Word dicetak pada ukuran **A3 Landscape (420 × 297 mm)** standar gambar teknik sipil dengan **Etiket resmi selebar 60 mm di sisi kanan**.")

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
            et_catatan = st.text_area("CATATAN (Kosongkan untuk otomatis dari hasil hitung)", "", height=100)

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
        f"📥 Unduh Dokumen Word Gambar Kerja Skalatis (.DOCX) — 2 Lembar A3 Landscape + Etiket Kanan",
        docx_bytes,
        file_name=f"Gambar_Kerja_Geometri_Jalan_Skala_1_{s_out1}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )

# ----------------------------------------------------------------------------
# TAB 6: TABEL RINGKASAN & EKSPOR EXCEL BER-FORMULA 17 LANGKAH
# ----------------------------------------------------------------------------
with tab_xls:
    st.markdown("<h3 style='color: #4A154B;'>📋 Tabel Ringkasan Parameter & Laporan Excel Formula 17 Langkah</h3>", unsafe_allow_html=True)
    
    fc = r["jenis"].startswith("Full")
    a, b = ("TC", "CT") if fc else ("TS", "ST")
    rows = [
        ("Jenis Tikungan", r["jenis"], "Acuan Kriteria R_fc TPGJAK 1997"),
        ("Kecepatan Rencana (Vr)", f"{r['V']:.0f}", "km/jam"),
        ("Sudut Tikungan (Δ)", f"{r['delta']:.2f}", "° (derajat)"),
        ("Jari-Jari Minimum (Rmin)", f"{r['Rmin']:.2f}", "m"),
        ("Jari-Jari Rencana (Rc)", f"{r['Rc']:.2f}", "m"),
        ("Superelevasi Maksimum (e_max)", f"{r['emax']*100:.1f}", "%"),
        ("Superelevasi Desain (e)", f"{r['e']*100:.2f}", "%"),
        ("Panjang Lengkung Peralihan (Ls)", f"{r['Ls_sup']:.2f}", "m"),
        ("Panjang Busur Lingkaran (Lc)", f"{r['Lc']:.2f}", "m"),
        ("Panjang Total Tikungan (Ltot)", f"{r['Ltot']:.3f}", "m"),
        ("Panjang Tangen (Ts)", f"{r['T']:.3f}", "m"),
        ("Jarak Eksternal (Es)", f"{r['E']:.3f}", "m"),
        ("Pergeseran Tangen (p)", f"{r['p']:.3f}", "m"),
        ("Absis Pusat Lingkaran (k)", f"{r['k']:.3f}", "m"),
        ("Kecepatan Aman Maksimum (V_safe)", f"{r['v_safe']:.2f}", "km/jam"),
        (f"STA {a} (Awal Tikungan)", sta_str(r["sta_awal"]), "m"),
        ("STA PI", sta_str(r["sta_pi"]), "m"),
        (f"STA {b} (Akhir Tikungan)", sta_str(r["sta_akhir"]), "m"),
    ]
    
    st.dataframe(pd.DataFrame(rows, columns=["Parameter Geometrik", "Nilai Terhitung", "Keterangan / Satuan"]), hide_index=True, use_container_width=True)

    excel_bytes = generate_excel_bm2021(r, P_global)
    st.download_button(
        "📊 Unduh Laporan Excel Ber-Formula Otomatis (.XLSX) — 17 Langkah Lengkap",
        excel_bytes, file_name="Perhitungan_Geometri_Jalan_BM2021.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
