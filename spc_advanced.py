import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import numpy as np
from scipy import stats
import os

# ============================================
# AYARLAR (CLOUD UYUMLU)
# ============================================
SABIT_DOSYA_ADI = "SPC 15.12.2025.xlsx"
LOGO_DOSYA_ADI = "logo.png"
SIRKET_ISMI = "MALHOTRA KABLO"

# ============================================
# TOLERANS DEĞERLERİ (DEĞİŞTİRİLEBİLİR)
# ============================================
# Her kesit için direnç ve birim ağırlık tolerans limitleri
# Değerleri buradan değiştirebilirsiniz veya sidebar'dan manuel giriş yapabilirsiniz

TOLERANS_DEGERLERI = {
    'BNT07C0.20': {'direnc_atl': 83.95, 'direnc_utl': 84.8, 'birim_agirlik_atl': 1.82, 'birim_agirlik_utl': 1.81},
    'BNH07X0.26': {'direnc_atl': 49.7, 'direnc_utl': 50.2, 'birim_agirlik_atl': 3.08, 'birim_agirlik_utl': 3.05},
    'BNH19X0.16': {'direnc_atl': 48.3, 'direnc_utl': 48.8, 'birim_agirlik_atl': 3.17, 'birim_agirlik_utl': 3.14},
    'BNF07X0.25': {'direnc_atl': 53.85, 'direnc_utl': 54.4, 'birim_agirlik_atl': 2.84, 'birim_agirlik_utl': 2.82},
    'BNF19X0.18': {'direnc_atl': 36.73, 'direnc_utl': 37.1, 'birim_agirlik_atl': 4.17, 'birim_agirlik_utl': 4.13},
    'BNH07X0.32': {'direnc_atl': 32.37, 'direnc_utl': 32.7, 'birim_agirlik_atl': 4.73, 'birim_agirlik_utl': 4.68},
    'BNH19X0.19': {'direnc_atl': 34.25, 'direnc_utl': 34.6, 'birim_agirlik_atl': 4.47, 'birim_agirlik_utl': 4.43},
    'BNF19X0.22': {'direnc_atl': 24.45, 'direnc_utl': 24.7, 'birim_agirlik_atl': 6.26, 'birim_agirlik_utl': 6.2},
    'BNH19X0.24': {'direnc_atl': 21.48, 'direnc_utl': 21.7, 'birim_agirlik_atl': 7.13, 'birim_agirlik_utl': 7.06},
    'BNF19X0.26': {'direnc_atl': 18.31, 'direnc_utl': 18.5, 'birim_agirlik_atl': 8.36, 'birim_agirlik_utl': 8.28},
    'BNH19X0.29': {'direnc_atl': 14.75, 'direnc_utl': 14.9, 'birim_agirlik_atl': 10.38, 'birim_agirlik_utl': 10.28},
    'BNH37X0.21': {'direnc_atl': 14.45, 'direnc_utl': 14.6, 'birim_agirlik_atl': 10.6, 'birim_agirlik_utl': 10.64},
    'BNF19X0.32': {'direnc_atl': 12.55, 'direnc_utl': 12.8, 'birim_agirlik_atl': 12.2, 'birim_agirlik_utl': 11.96},
    'BNH37X0.26': {'direnc_atl': 9.41, 'direnc_utl': 9.6, 'birim_agirlik_atl': 16.27, 'birim_agirlik_utl': 15.95},
    'BNF19X0.38': {'direnc_atl': 9.36, 'direnc_utl': 9.55, 'birim_agirlik_atl': 16.36, 'birim_agirlik_utl': 16.04},
    'BNF37X0.30': {'direnc_atl': 7.55, 'direnc_utl': 7.7, 'birim_agirlik_atl': 20.28, 'birim_agirlik_utl': 19.89},
    'BNF50X0.25': {'direnc_atl': 7.55, 'direnc_utl': 7.7, 'birim_agirlik_atl': 20.28, 'birim_agirlik_utl': 19.89},
    'BNH41X0.32': {'direnc_atl': 5.55, 'direnc_utl': 5.7, 'birim_agirlik_atl': 27.59, 'birim_agirlik_utl': 26.87},
    'BNF56X0.30': {'direnc_atl': 4.75, 'direnc_utl': 4.99, 'birim_agirlik_atl': 32.24, 'birim_agirlik_utl': 30.69},
    'BNF84X0.30': {'direnc_atl': 3.19, 'direnc_utl': 3.35, 'birim_agirlik_atl': 48.01, 'birim_agirlik_utl': 45.72},
    'BNH50X0.45': {'direnc_atl': 2.36, 'direnc_utl': 2.48, 'birim_agirlik_atl': 64.89, 'birim_agirlik_utl': 61.75},
    'BNF80X0.40': {'direnc_atl': 1.84, 'direnc_utl': 1.94, 'birim_agirlik_atl': 83.23, 'birim_agirlik_utl': 78.94},
    'BNF147X0.31': {'direnc_atl': 1.84, 'direnc_utl': 1.94, 'birim_agirlik_atl': 83.23, 'birim_agirlik_utl': 78.94},
    'BNF126X0.40': {'direnc_atl': 1.18, 'direnc_utl': 1.24, 'birim_agirlik_atl': 129.79, 'birim_agirlik_utl': 123.51},
    'BNC196X0.41': {'direnc_atl': 0.75, 'direnc_utl': 0.79, 'birim_agirlik_atl': 204.2, 'birim_agirlik_utl': 193.86},
}

# ============================================
# PARAMETRE TİPLERİ (Tolerans Eşleştirmesi İçin)
# ============================================
# Hangi parametrelerin direnç, hangilerinin birim ağırlık olduğunu belirtir
PARAMETRE_TIPLERI = {
    'Tartılan Birim Ağırlık (g/m)': 'birim_agirlik',
    'Başlangıç Birim Ağırlık (g/m)': 'birim_agirlik',
    'Başlangıç Direnç (Ω)': 'direnc',
    'Bitiş Direnç (Ω)': 'direnc',
    'Başlangıç CR': None,  # Tolerans yok
    'Bitiş CR': None,  # Tolerans yok
    'Başlangıç - Bitiş CR Farkı': None,  # Tolerans yok
}

# Sayfa ayarları
st.set_page_config(
    page_title=f"SPC - {SIRKET_ISMI}", 
    layout="wide", 
    page_icon="📊",
    initial_sidebar_state="expanded"
)

# ============================================
# RENK PALETİ (Renk körü dostu)
# ============================================
COLORS = {
    'primary': '#2563eb',       # Mavi
    'success': '#16a34a',       # Yeşil
    'warning': '#d97706',       # Turuncu
    'danger': '#dc2626',        # Kırmızı
    'purple': '#7c3aed',        # Mor
    'gray': '#6b7280',          # Gri
    'light_blue': '#93c5fd',
    'light_green': '#86efac',
    'light_red': '#fca5a5',
    'background': '#f8fafc'
}

# ============================================
# SÜTUN EŞLEŞTİRME (İSTEDİĞİNİZ 7 PARAMETRE)
# ============================================
COL_DATE = 'TARİH'
COL_GROUP = 'KESİT'
COL_MACHINE = 'MAKİNE'

PARAM_MAP = {
    'Tartılan Birim Ağırlık (g/m)': {
        'sutun': 'TARTILAN BİRİM AĞIRLIK',
        'aciklama': 'Numunenin tartılan gerçek birim ağırlığı.',
        'birim': 'g/m',
        'icon': '⚖️'
    },
    'Başlangıç Birim Ağırlık (g/m)': {
        'sutun': 'BAŞLANGIÇ BİRİM AĞIRLIK',
        'aciklama': 'Üretim başlangıcında ölçülen birim ağırlık.',
        'birim': 'g/m',
        'icon': '⚖️'
    },
    'Başlangıç Direnç (Ω)': {
        'sutun': 'KALİTE BAŞLANGIÇ ÖLÇÜLEN DİRENÇ',
        'aciklama': 'Üretim başlangıcında ölçülen direnç değeri.',
        'birim': 'Ω',
        'icon': '🔌'
    },
    'Bitiş Direnç (Ω)': {
        'sutun': 'KALİTE BİTİŞ ÖLÇÜLEN DİRENÇ',
        'aciklama': 'Üretim bitişinde ölçülen direnç değeri.',
        'birim': 'Ω',
        'icon': '🔌'
    },
    'Başlangıç CR': {
        'sutun': 'BAŞLANGIÇ CR',
        'aciklama': 'Üretim başlangıç iletken direnç (CR) değeri.',
        'birim': '-',
        'icon': '📈'
    },
    'Bitiş CR': {
        'sutun': 'BİTİŞ CR',
        'aciklama': 'Üretim bitiş iletken direnç (CR) değeri.',
        'birim': '-',
        'icon': '📉'
    },
    'Başlangıç - Bitiş CR Farkı': {
        'sutun': 'BAŞLANGIÇ - BİTİŞ CR',
        'aciklama': 'Başlangıç ve bitiş CR değerleri arasındaki fark.',
        'birim': '-',
        'icon': '📊'
    }
}

# ============================================
# CSS STİLLERİ (ORİJİNAL DETAYLI TASARIM)
# ============================================
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: 700;
        color: #1e3a5f;
        margin-bottom: 0.5rem;
        padding-bottom: 0.5rem;
        border-bottom: 3px solid #2563eb;
    }
    
    .section-header {
        font-size: 1.4rem;
        font-weight: 600;
        color: #1e3a5f;
        margin: 1.5rem 0 1rem 0;
        padding: 0.5rem 0;
        border-left: 4px solid #2563eb;
        padding-left: 1rem;
    }
    
    .metric-card {
        background: white;
        padding: 1.25rem;
        border-radius: 10px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        text-align: center;
        margin: 0.5rem 0;
        border-left: 4px solid #2563eb;
    }
    
    .metric-card.excellent { border-left-color: #16a34a; background: linear-gradient(to right, #f0fdf4, white); }
    .metric-card.good { border-left-color: #22c55e; }
    .metric-card.warning { border-left-color: #d97706; background: linear-gradient(to right, #fffbeb, white); }
    .metric-card.danger { border-left-color: #dc2626; background: linear-gradient(to right, #fef2f2, white); }
    
    .metric-value {
        font-size: 1.8rem;
        font-weight: bold;
        color: #1e3a5f;
        margin: 0.3rem 0;
    }
    
    .metric-label {
        font-size: 0.9rem;
        color: #64748b;
        font-weight: 500;
    }
    
    .metric-desc {
        font-size: 0.8rem;
        color: #94a3b8;
        margin-top: 0.3rem;
    }
    
    .info-box {
        background: #f1f5f9;
        padding: 1rem 1.25rem;
        border-radius: 8px;
        border-left: 4px solid #2563eb;
        margin: 0.75rem 0;
        font-size: 0.9rem;
    }
    
    .tolerance-box {
        background: #fef3c7;
        padding: 0.75rem 1rem;
        border-radius: 8px;
        border-left: 4px solid #d97706;
        margin: 0.5rem 0;
        font-size: 0.85rem;
    }
    
    .alert-box {
        padding: 1rem;
        border-radius: 8px;
        margin: 0.75rem 0;
    }
    
    .alert-success { background: #f0fdf4; border: 1px solid #86efac; color: #166534; }
    .alert-warning { background: #fffbeb; border: 1px solid #fcd34d; color: #92400e; }
    .alert-danger { background: #fef2f2; border: 1px solid #fca5a5; color: #991b1b; }
    
    .rule-violation {
        background: #fef2f2;
        border: 1px solid #fca5a5;
        border-radius: 6px;
        padding: 0.5rem 0.75rem;
        margin: 0.25rem 0;
        font-size: 0.85rem;
        color: #991b1b;
    }
    
    .rule-ok {
        background: #f0fdf4;
        border: 1px solid #86efac;
        border-radius: 6px;
        padding: 0.5rem 0.75rem;
        margin: 0.25rem 0;
        font-size: 0.85rem;
        color: #166534;
    }
</style>
""", unsafe_allow_html=True)

# ============================================
# TOLERANS YARDIMCI FONKSİYONLARI
# ============================================

def get_tolerans_for_kesit(kesit, parametre_adi):
    """Seçilen kesit ve parametre için tolerans değerlerini döndürür"""
    param_tipi = PARAMETRE_TIPLERI.get(parametre_adi)
    
    if param_tipi is None:
        # CR parametreleri için tolerans yok
        return None, None
    
    if kesit not in TOLERANS_DEGERLERI:
        return None, None
    
    tolerans = TOLERANS_DEGERLERI[kesit]
    
    if param_tipi == 'direnc':
        return tolerans['direnc_atl'], tolerans['direnc_utl']
    elif param_tipi == 'birim_agirlik':
        return tolerans['birim_agirlik_atl'], tolerans['birim_agirlik_utl']
    
    return None, None

def get_all_kesit_codes():
    """Tüm tanımlı kesit kodlarını döndürür"""
    return list(TOLERANS_DEGERLERI.keys())

# ============================================
# GELİŞMİŞ YARDIMCI FONKSİYONLAR
# ============================================

def calculate_spc_metrics(data, usl=None, lsl=None):
    """Tüm SPC metriklerini hesapla - KAPSAMLI VERSİYON"""
    n = len(data)
    if n < 2: return None
    
    mean = data.mean()
    std_sample = data.std(ddof=1)
    
    # Moving Range
    mr = data.diff().abs().dropna()
    mr_mean = mr.mean() if len(mr) > 0 else 0
    sigma_within = mr_mean / 1.128 if mr_mean > 0 else std_sample
    
    ucl = mean + 3 * sigma_within
    lcl = mean - 3 * sigma_within
    mr_ucl = 3.267 * mr_mean if mr_mean > 0 else 0
    
    # Sigma bantları
    sigma_bands = {
        '1sigma_upper': mean + sigma_within,
        '1sigma_lower': mean - sigma_within,
        '2sigma_upper': mean + 2 * sigma_within,
        '2sigma_lower': mean - 2 * sigma_within,
    }
    
    # Yeterlilik
    cp, cpk, cpu, cpl = None, None, None, None
    pp, ppk, ppu, ppl = None, None, None, None
    ppm, ppm_upper, ppm_lower = None, None, None
    sigma_level = None
    
    if usl is not None and lsl is not None and sigma_within > 0:
        cp = (usl - lsl) / (6 * sigma_within)
        cpu = (usl - mean) / (3 * sigma_within)
        cpl = (mean - lsl) / (3 * sigma_within)
        cpk = min(cpu, cpl)
        
        if std_sample > 0:
            pp = (usl - lsl) / (6 * std_sample)
            ppu = (usl - mean) / (3 * std_sample)
            ppl = (mean - lsl) / (3 * std_sample)
            ppk = min(ppu, ppl)
        
        z_upper = (usl - mean) / sigma_within
        z_lower = (mean - lsl) / sigma_within
        ppm = ((1 - stats.norm.cdf(z_upper)) + stats.norm.cdf(-z_lower)) * 1e6
        
        if cpk > 0: sigma_level = cpk * 3
    
    # Normallik
    normality_stat, normality_p = None, None
    if 3 <= n <= 5000:
        try: _, normality_p = stats.shapiro(data)
        except: pass
    
    return {
        'mean': mean, 'std': std_sample, 'sigma_within': sigma_within,
        'ucl': ucl, 'lcl': lcl, 'mr_mean': mr_mean, 'mr_ucl': mr_ucl,
        'sigma_bands': sigma_bands,
        'cp': cp, 'cpk': cpk, 'pp': pp, 'ppk': ppk, 'ppm': ppm,
        'sigma_level': sigma_level, 'normality_p': normality_p,
        'skewness': stats.skew(data), 'kurtosis': stats.kurtosis(data),
        'n': n, 'min': data.min(), 'max': data.max(), 'median': data.median(),
        'iqr': data.quantile(0.75) - data.quantile(0.25)
    }

def check_western_electric_rules(data, mean, sigma):
    """Western Electric Kuralları"""
    violations = []
    n = len(data)
    if n < 8: return violations, []
    
    # Kural 1: 3 sigma dışı
    r1 = [i for i, v in enumerate(data) if abs(v - mean) > 3 * sigma]
    if r1: violations.append({'rule': 1, 'desc': '3σ dışında nokta', 'points': r1, 'severity': 'high'})
    
    # Kural 2: 3 noktadan 2'si 2 sigma dışı
    r2 = []
    for i in range(2, n):
        w = data.iloc[i-2:i+1]
        if ((w - mean) > 2*sigma).sum() >= 2 or ((mean - w) > 2*sigma).sum() >= 2: r2.append(i)
    if r2: violations.append({'rule': 2, 'desc': '3 noktadan 2si 2σ bölgesinde', 'points': r2, 'severity': 'medium'})
    
    # Kural 3: 5 noktadan 4'ü 1 sigma dışı
    r3 = []
    for i in range(4, n):
        w = data.iloc[i-4:i+1]
        if ((w - mean) > sigma).sum() >= 4 or ((mean - w) > sigma).sum() >= 4: r3.append(i)
    if r3: violations.append({'rule': 3, 'desc': '5 noktadan 4ü 1σ bölgesinde', 'points': r3, 'severity': 'medium'})

    # Kural 4: 8 nokta aynı tarafta
    r4 = []
    for i in range(7, n):
        w = data.iloc[i-7:i+1]
        if all(w > mean) or all(w < mean): r4.append(i)
    if r4: violations.append({'rule': 4, 'desc': '8 ardışık nokta aynı tarafta', 'points': r4, 'severity': 'medium'})
    
    # Kural 5: 6 nokta trend
    r5 = []
    for i in range(5, n):
        w = data.iloc[i-5:i+1].values
        d = np.diff(w)
        if all(d > 0) or all(d < 0): r5.append(i)
    if r5: violations.append({'rule': 5, 'desc': '6 ardışık nokta trend', 'points': r5, 'severity': 'medium'})

    all_points = set()
    for v in violations: all_points.update(v['points'])
    return violations, list(all_points)

def get_cpk_info(cpk):
    if cpk is None: return {'color': COLORS['gray'], 'status': 'Belirsiz', 'icon': '❓', 'class': '', 'desc': 'Tolerans girilmedi', 'action': 'Limitleri girin'}
    if cpk >= 1.67: return {'color': COLORS['success'], 'status': 'Mükemmel', 'icon': '🌟', 'class': 'excellent', 'desc': 'Dünya standardı', 'action': 'Sürdürün'}
    if cpk >= 1.33: return {'color': '#22c55e', 'status': 'İyi', 'icon': '✅', 'class': 'good', 'desc': 'Yeterli', 'action': 'İzleyin'}
    if cpk >= 1.00: return {'color': COLORS['warning'], 'status': 'Kabul Edilebilir', 'icon': '⚠️', 'class': 'warning', 'desc': 'Sınırda', 'action': 'İyileştirin'}
    return {'color': COLORS['danger'], 'status': 'Yetersiz', 'icon': '❌', 'class': 'danger', 'desc': 'Riskli Süreç', 'action': 'Acil müdahale'}

def create_gauge_chart(value, title):
    if value is None: value = 0
    fig = go.Figure(go.Indicator(
        mode="gauge+number+delta", value=value,
        title={'text': title},
        delta={'reference': 1.33},
        gauge={'axis': {'range': [0, 2.5]},
               'bar': {'color': get_cpk_info(value)['color']},
               'steps': [{'range': [0, 1], 'color': '#fecaca'}, {'range': [1, 1.33], 'color': '#fed7aa'}, {'range': [1.33, 2.5], 'color': '#bbf7d0'}],
               'threshold': {'line': {'color': "black", 'width': 4}, 'thickness': 0.75, 'value': 1.33}}))
    fig.update_layout(height=280, margin=dict(l=30, r=30, t=50, b=30))
    return fig

def create_capability_histogram(data, mean, sigma, usl=None, lsl=None, title="Yetenek Histogramı"):
    fig = go.Figure()
    fig.add_trace(go.Histogram(x=data, nbinsx=35, name='Dağılım', marker_color=COLORS['primary'], opacity=0.7))
    
    x_range = np.linspace(data.min() - sigma, data.max() + sigma, 200)
    y_norm = stats.norm.pdf(x_range, mean, sigma)
    scale = len(data) * (data.max() - data.min()) / 35
    fig.add_trace(go.Scatter(x=x_range, y=y_norm * scale, mode='lines', name='Normal Dağılım', line=dict(color=COLORS['danger'], width=3)))
    
    fig.add_vline(x=mean, line_color=COLORS['success'], line_width=3, annotation_text="Ort.")
    if usl: fig.add_vline(x=usl, line_color=COLORS['warning'], line_dash='dash', annotation_text="USL")
    if lsl: fig.add_vline(x=lsl, line_color=COLORS['warning'], line_dash='dash', annotation_text="LSL")
    
    fig.update_layout(title=title, height=400, template="plotly_white")
    return fig

def create_control_chart(data, x_axis, metrics, usl=None, lsl=None, title="I-MR Kontrol Grafiği", 
                         violation_points=None, param_name="", unit=""):
    """Gelişmiş kontrol grafiği (Sigma Bantlı)"""
    fig = go.Figure()
    mean, sigma, ucl, lcl = metrics['mean'], metrics['sigma_within'], metrics['ucl'], metrics['lcl']
    
    # Sigma bantları
    fig.add_hrect(y0=mean-sigma, y1=mean+sigma, fillcolor="rgba(34, 197, 94, 0.12)", line_width=0)
    fig.add_hrect(y0=mean-2*sigma, y1=mean-sigma, fillcolor="rgba(234, 179, 8, 0.08)", line_width=0)
    fig.add_hrect(y0=mean+sigma, y1=mean+2*sigma, fillcolor="rgba(234, 179, 8, 0.08)", line_width=0)
    fig.add_hrect(y0=mean-3*sigma, y1=mean-2*sigma, fillcolor="rgba(239, 68, 68, 0.06)", line_width=0)
    fig.add_hrect(y0=mean+2*sigma, y1=mean+3*sigma, fillcolor="rgba(239, 68, 68, 0.06)", line_width=0)
    
    fig.add_trace(go.Scatter(x=x_axis, y=data, mode='lines+markers', name='Ölçüm', line=dict(color=COLORS['primary'])))
    
    fig.add_hline(y=mean, line_color=COLORS['success'], line_width=3, annotation_text="Ort.")
    fig.add_hline(y=ucl, line_color=COLORS['danger'], line_dash='dash', annotation_text="UCL")
    fig.add_hline(y=lcl, line_color=COLORS['danger'], line_dash='dash', annotation_text="LCL")
    
    if usl: fig.add_hline(y=usl, line_color=COLORS['warning'], line_dash='dot', annotation_text="USL")
    if lsl: fig.add_hline(y=lsl, line_color=COLORS['warning'], line_dash='dot', annotation_text="LSL")
    
    # Kontrol dışı noktalar
    out_mask = (data > ucl) | (data < lcl)
    out_points = data[out_mask]
    if len(out_points) > 0:
        # HATA DÜZELTMESİ: .iloc yerine güvenli indeksleme
        if hasattr(x_axis, 'loc'):
            out_x = x_axis.loc[out_points.index]
        else:
            out_x = [x_axis[i] for i in range(len(x_axis)) if i in out_points.index]
            
        fig.add_trace(go.Scatter(x=out_x, y=out_points, mode='markers', name='Kontrol Dışı', 
                                marker=dict(color=COLORS['danger'], size=12, symbol='x')))
    
    # Kural ihlalleri
    if violation_points:
        # HATA DÜZELTMESİ: violation_points tamsayı indekslerdir, x_axis'e göre güvenli erişim
        if hasattr(x_axis, 'iloc'):
            viol_x = x_axis.iloc[violation_points]
        else:
            viol_x = [x_axis[i] for i in violation_points]
            
        fig.add_trace(go.Scatter(x=viol_x, y=data.iloc[violation_points], mode='markers', name='Kural İhlali',
                                marker=dict(color=COLORS['purple'], size=10, symbol='diamond')))
    
    fig.update_layout(title=title, height=550, template="plotly_white")
    return fig

def create_mr_chart(data, x_axis, mr_mean, mr_ucl, title="Moving Range (MR) Grafiği"):
    """Moving Range Grafiği (DÜZELTİLMİŞ)"""
    mr = data.diff().abs()
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=x_axis, y=mr, mode='lines+markers', name='MR', line=dict(color=COLORS['purple'])))
    fig.add_hline(y=mr_mean, line_color=COLORS['success'], annotation_text="MR Ort")
    fig.add_hline(y=mr_ucl, line_color=COLORS['danger'], line_dash='dash', annotation_text="UCL")
    
    # HATA DÜZELTMESİ BURADA YAPILDI
    mr_out = mr[mr > mr_ucl]
    if len(mr_out) > 0:
        if hasattr(x_axis, 'loc'):
            mr_out_x = x_axis.loc[mr_out.index]
        else:
            mr_out_x = [x_axis[i] for i in range(len(x_axis)) if i in mr_out.index]
            
        fig.add_trace(go.Scatter(x=mr_out_x, y=mr_out, mode='markers', name='MR Kontrol Dışı',
                                marker=dict(color=COLORS['danger'], size=10, symbol='x')))
    
    fig.update_layout(title=title, height=350, template="plotly_white")
    return fig

def count_out_of_limits(data, upper=None, lower=None):
    count = 0
    if upper is not None: count += (data > upper).sum()
    if lower is not None: count += (data < lower).sum()
    return int(count)

def format_number(val):
    if val is None: return "-"
    return f"{val:,.4f}" if abs(val) < 1000 else f"{val:,.2f}"

def load_data(file_path):
    try:
        if file_path.endswith('.csv'): df = pd.read_csv(file_path)
        else: df = pd.read_excel(file_path)
        return df, None
    except Exception as e: return None, str(e)

# ============================================
# UYGULAMA AKIŞI
# ============================================
if os.path.exists(LOGO_DOSYA_ADI):
    st.sidebar.image(LOGO_DOSYA_ADI, width=180)

st.sidebar.title("Ayarlar")

# Veri Yükleme
df = None
if os.path.exists(SABIT_DOSYA_ADI):
    df, error = load_data(SABIT_DOSYA_ADI)
    if error: st.sidebar.error(f"Sabit dosya hatası: {error}")
else:
    uploaded = st.sidebar.file_uploader("Veri Dosyası", type=['xlsx', 'csv'])
    if uploaded:
        if uploaded.name.endswith('.csv'): df = pd.read_csv(uploaded)
        else: df = pd.read_excel(uploaded)

if df is None:
    st.markdown('<p class="main-header">📊 Gelişmiş SPC Analiz Sistemi</p>', unsafe_allow_html=True)
    st.info(f"📁 Lütfen veri dosyası yükleyin veya **{os.path.basename(SABIT_DOSYA_ADI)}** dosyasını ekleyin.")
    st.stop()

# Sütun isimlerini temizle (newline, fazla boşluk vb.)
df.columns = df.columns.str.replace('\n', ' ').str.replace('  ', ' ').str.strip()

# Tarih
if COL_DATE in df.columns:
    df[COL_DATE] = pd.to_datetime(df[COL_DATE], errors='coerce')

# Orijinal veriyi sakla (haftalık analiz için - tarih dönüşümünden sonra)
df_original = df.copy()

# Filtreler
st.sidebar.subheader("Filtreler")
if COL_DATE in df.columns:
    from datetime import datetime, timedelta
    
    # Yıl ve Hafta numarası sütunları ekle (ISO hafta numarası)
    df['_YIL'] = df[COL_DATE].dt.isocalendar().year
    df['_HAFTA'] = df[COL_DATE].dt.isocalendar().week
    df['_YIL_HAFTA'] = df['_YIL'].astype(str) + '-W' + df['_HAFTA'].astype(str).str.zfill(2)
    
    # Mevcut hafta kombinasyonlarını al ve sırala (küçükten büyüğe)
    hafta_listesi = df[['_YIL', '_HAFTA', '_YIL_HAFTA']].drop_duplicates().sort_values(
        by=['_YIL', '_HAFTA'], ascending=[True, True]
    ).reset_index(drop=True)
    
    # Hafta seçeneklerini oluştur (41.HAFTA, 42.HAFTA formatında - gerçek hafta numaraları)
    hafta_secenekleri = ['Tümü'] + [f"{int(row['_HAFTA'])}.HAFTA" for _, row in hafta_listesi.iterrows()] + ['Manuel Tarih Seçimi']
    hafta_kodlari = ['Tümü'] + hafta_listesi['_YIL_HAFTA'].tolist() + ['Manuel']
    
    # Hafta seçimi
    secili_hafta_idx = st.sidebar.selectbox(
        "📅 Hafta Seçimi", 
        range(len(hafta_secenekleri)),
        format_func=lambda x: hafta_secenekleri[x]
    )
    
    if hafta_kodlari[secili_hafta_idx] == 'Manuel':
        # Manuel tarih seçimi
        min_d, max_d = df[COL_DATE].min().date(), df[COL_DATE].max().date()
        dr = st.sidebar.date_input("📆 Tarih Aralığı", value=(min_d, max_d))
        if len(dr) == 2:
            df = df[(df[COL_DATE].dt.date >= dr[0]) & (df[COL_DATE].dt.date <= dr[1])]
            st.sidebar.caption(f"📆 {dr[0].strftime('%d.%m.%Y')} - {dr[1].strftime('%d.%m.%Y')}")
    elif secili_hafta_idx > 0:  # "Tümü" değilse
        secili_hafta_kodu = hafta_kodlari[secili_hafta_idx]
        df = df[df['_YIL_HAFTA'] == secili_hafta_kodu]
        
        # Seçilen haftanın tarih aralığını göster
        hafta_baslangic = df[COL_DATE].min()
        hafta_bitis = df[COL_DATE].max()
        st.sidebar.caption(f"📆 {hafta_baslangic.strftime('%d.%m.%Y')} - {hafta_bitis.strftime('%d.%m.%Y')}")
    
    # Yardımcı sütunları temizle
    df = df.drop(columns=['_YIL', '_HAFTA', '_YIL_HAFTA'])

# Kesit seçimi - Tolerans için önemli
sec_kesit = 'Tümü'
if COL_GROUP in df.columns:
    kesitler = ['Tümü'] + sorted(df[COL_GROUP].dropna().unique().tolist())
    sec_kesit = st.sidebar.selectbox("Kesit", kesitler)
    if sec_kesit != 'Tümü': df = df[df[COL_GROUP] == sec_kesit]

if COL_MACHINE in df.columns:
    makineler = ['Tümü'] + sorted(df[COL_MACHINE].dropna().astype(str).unique().tolist())
    sec_makine = st.sidebar.selectbox("Makine", makineler)
    if sec_makine != 'Tümü': df = df[df[COL_MACHINE].astype(str) == sec_makine]

st.sidebar.markdown(f"**Kayıt:** {len(df)}")

# Ana Ekran
st.markdown(f'<p class="main-header">📊 SPC Analiz - {SIRKET_ISMI}</p>', unsafe_allow_html=True)

# Parametre
mevcut_params = {k: v for k, v in PARAM_MAP.items() if v['sutun'] in df.columns}
if not mevcut_params:
    st.error("Veri setinde analiz edilecek uygun sütun bulunamadı.")
    st.stop()

col_sel, col_info = st.columns([1, 2])
with col_sel:
    secili_key = st.selectbox("Parametre Seçiniz", list(mevcut_params.keys()))
param = mevcut_params[secili_key]

with col_info:
    st.markdown(f'<div class="info-box"><b>{param["icon"]} {secili_key}</b><br>{param["aciklama"]}</div>', unsafe_allow_html=True)

# ============================================
# TOLERANS DEĞERLERİ (OTOMATİK + MANUEL)
# ============================================
st.markdown("### 🎯 Tolerans Limitleri")

# Seçilen kesit için otomatik tolerans değerlerini al
auto_lsl, auto_usl = None, None
param_tipi = PARAMETRE_TIPLERI.get(secili_key)

if sec_kesit != 'Tümü' and sec_kesit in TOLERANS_DEGERLERI:
    auto_lsl, auto_usl = get_tolerans_for_kesit(sec_kesit, secili_key)

# Tolerans ayarlama modu
tol_col1, tol_col2, tol_col3 = st.columns([1, 1, 1])

with tol_col1:
    tolerans_modu = st.radio(
        "Tolerans Kaynağı",
        ["Otomatik (Kesit'e göre)", "Manuel Giriş"],
        index=0 if (auto_lsl is not None and auto_usl is not None) else 1,
        horizontal=True
    )

# Bilgi kutusu
if param_tipi is None:
    st.markdown('<div class="tolerance-box">ℹ️ <b>CR parametreleri</b> için tolerans limiti tanımlanmamıştır.</div>', unsafe_allow_html=True)
elif sec_kesit == 'Tümü':
    st.markdown('<div class="tolerance-box">⚠️ <b>Tümü</b> seçiliyken otomatik tolerans kullanılamaz. Lütfen kesit seçin veya manuel giriş yapın.</div>', unsafe_allow_html=True)
elif sec_kesit not in TOLERANS_DEGERLERI:
    st.markdown(f'<div class="tolerance-box">⚠️ <b>{sec_kesit}</b> kesiti için tolerans tanımlı değil. Manuel giriş yapabilirsiniz.</div>', unsafe_allow_html=True)

# Tolerans değerleri
c1, c2 = st.columns(2)

if tolerans_modu == "Otomatik (Kesit'e göre)" and auto_lsl is not None and auto_usl is not None:
    # Otomatik değerler göster ama değiştirilebilir
    with c1:
        lsl = st.number_input("LSL (Alt Limit)", value=float(auto_lsl), format="%.4f", 
                              help="Otomatik değer yüklendi - değiştirebilirsiniz")
    with c2:
        usl = st.number_input("USL (Üst Limit)", value=float(auto_usl), format="%.4f",
                              help="Otomatik değer yüklendi - değiştirebilirsiniz")
    
    # Otomatik yüklendiğini göster
    st.success(f"✅ **{sec_kesit}** kesiti için tolerans değerleri otomatik yüklendi: LSL={auto_lsl}, USL={auto_usl}")
else:
    # Manuel giriş
    with c1:
        lsl = st.number_input("LSL (Alt Limit)", value=None, format="%.4f",
                              help="Alt spesifikasyon limitini girin")
    with c2:
        usl = st.number_input("USL (Üst Limit)", value=None, format="%.4f",
                              help="Üst spesifikasyon limitini girin")

# Tolerans tablosu göster butonu
with st.expander("📋 Tanımlı Tolerans Değerlerini Görüntüle"):
    tol_df_data = []
    for kod, degerler in TOLERANS_DEGERLERI.items():
        tol_df_data.append({
            'Kesit Kodu': kod,
            'Direnç ATL (Ω)': degerler['direnc_atl'],
            'Direnç UTL (Ω)': degerler['direnc_utl'],
            'B.Ağırlık ATL (g/m)': degerler['birim_agirlik_atl'],
            'B.Ağırlık UTL (g/m)': degerler['birim_agirlik_utl'],
        })
    tol_df = pd.DataFrame(tol_df_data)
    st.dataframe(tol_df, use_container_width=True, height=400)
    st.info("💡 Bu değerleri değiştirmek için programın başındaki **TOLERANS_DEGERLERI** dictionary'sini düzenleyin.")

# Veri Hazırlama
if COL_DATE in df.columns: df = df.sort_values(COL_DATE)
data = df[param['sutun']].dropna()

if len(data) < 2:
    st.warning("Yeterli veri yok.")
    st.stop()

# X Ekseni (Hata kaynağı burasıydı, şimdi güvenli)
if COL_DATE in df.columns:
    x_axis = df.loc[data.index, COL_DATE]
else:
    x_axis = list(range(len(data)))

# Analiz
metrics = calculate_spc_metrics(data, usl, lsl)
violations, v_points = check_western_electric_rules(data, metrics['mean'], metrics['sigma_within'])

# Sonuçlar
st.markdown("### 📊 Analiz Sonuçları")
k1, k2, k3, k4, k5 = st.columns(5)
k1.markdown(f'<div class="metric-card"><div class="metric-label">Ortalama</div><div class="metric-value">{format_number(metrics["mean"])}</div></div>', unsafe_allow_html=True)
k2.markdown(f'<div class="metric-card"><div class="metric-label">Std Sapma</div><div class="metric-value">{format_number(metrics["sigma_within"])}</div></div>', unsafe_allow_html=True)
k3.markdown(f'<div class="metric-card warning"><div class="metric-label">UCL</div><div class="metric-value">{format_number(metrics["ucl"])}</div></div>', unsafe_allow_html=True)
k4.markdown(f'<div class="metric-card warning"><div class="metric-label">LCL</div><div class="metric-value">{format_number(metrics["lcl"])}</div></div>', unsafe_allow_html=True)

out_c = count_out_of_limits(data, metrics['ucl'], metrics['lcl'])
cls = "danger" if out_c > 0 else "excellent"
k5.markdown(f'<div class="metric-card {cls}"><div class="metric-label">Kontrol Dışı</div><div class="metric-value">{out_c}</div></div>', unsafe_allow_html=True)

# Yeterlilik
if metrics['cpk']:
    st.markdown("### ⭐ Süreç Yeterliliği")
    info = get_cpk_info(metrics['cpk'])
    g1, g2, g3 = st.columns([1.2, 1.5, 1.3])
    with g1: st.plotly_chart(create_gauge_chart(metrics['cpk'], "Cpk"), use_container_width=True)
    with g2:
        st.markdown(f'<div class="metric-card"><div class="metric-label">Cp</div><div class="metric-value">{metrics["cp"]:.2f}</div></div>', unsafe_allow_html=True)
        st.markdown(f'<div class="metric-card {info["class"]}"><div class="metric-label">Cpk</div><div class="metric-value">{metrics["cpk"]:.2f}</div></div>', unsafe_allow_html=True)
    with g3:
        st.markdown(f'<div class="alert-box alert-{info["class"] if info["class"]!="excellent" else "success"}"><b>{info["status"]}</b><br>{info["desc"]}<br><b>Aksiyon:</b> {info["action"]}</div>', unsafe_allow_html=True)

# Normallik
if metrics['normality_p'] is not None and metrics['normality_p'] <= 0.05:
    st.warning(f"⚠️ Veriler normal dağılmıyor (p={metrics['normality_p']:.4f}).")

# Western Electric
if violations:
    with st.expander(f"⚠️ {len(violations)} Kural İhlali Tespit Edildi", expanded=True):
        for v in violations:
            st.markdown(f'<div class="rule-violation">🔴 <b>Kural {v["rule"]}:</b> {v["desc"]} ({len(v["points"])} nokta)</div>', unsafe_allow_html=True)

# Grafikler
st.markdown("### 📈 Grafikler")
tab1, tab2, tab3 = st.tabs(["Kontrol Grafiği", "Histogram", "Moving Range"])

with tab1:
    fig_ctrl = create_control_chart(data, x_axis, metrics, usl, lsl, title=f"I-MR Grafiği: {secili_key}", violation_points=v_points)
    st.plotly_chart(fig_ctrl, use_container_width=True)

with tab2:
    fig_hist = create_capability_histogram(data, metrics['mean'], metrics['sigma_within'], usl, lsl, title=f"Dağılım: {secili_key}")
    st.plotly_chart(fig_hist, use_container_width=True)

with tab3:
    fig_mr = create_mr_chart(data, x_axis, metrics['mr_mean'], metrics['mr_ucl'])
    st.plotly_chart(fig_mr, use_container_width=True)

# Makine
if COL_MACHINE in df.columns and df[COL_MACHINE].nunique() > 1:
    st.markdown("### 🏭 Makine Karşılaştırma")
    c_tbl, c_box = st.columns([1, 2])
    stats_df = df.groupby(COL_MACHINE)[param['sutun']].agg(['count', 'mean', 'std', 'min', 'max']).round(4)
    with c_tbl: st.dataframe(stats_df, height=400)
    with c_box:
        fig_box = px.box(df, x=COL_MACHINE, y=param['sutun'], color=COL_MACHINE)
        st.plotly_chart(fig_box, use_container_width=True)

# ============================================
# HAFTA BAZLI CPK TABLOSU
# ============================================
if COL_DATE in df_original.columns and sec_kesit != 'Tümü' and usl is not None and lsl is not None:
    st.markdown("### 📊 Haftalık Cpk Analizi")
    
    # Orijinal veriyi kullan (filtrelenmemiş kesit verisi)
    df_kesit = df_original[df_original[COL_GROUP] == sec_kesit].copy()
    df_kesit['_HAFTA'] = df_kesit[COL_DATE].dt.isocalendar().week
    
    # Her hafta için Cpk hesapla
    hafta_cpk_data = []
    for hafta in sorted(df_kesit['_HAFTA'].unique()):
        hafta_df = df_kesit[df_kesit['_HAFTA'] == hafta]
        hafta_data = hafta_df[param['sutun']].dropna()
        
        if len(hafta_data) >= 2:
            hafta_metrics = calculate_spc_metrics(hafta_data, usl, lsl)
            
            if hafta_metrics:
                cpk_val = hafta_metrics['cpk']
                cp_val = hafta_metrics['cp']
                
                # Durum belirleme
                if cpk_val is None:
                    durum = "❓ Belirsiz"
                    durum_renk = "gray"
                elif cpk_val >= 1.67:
                    durum = "🌟 Mükemmel"
                    durum_renk = "#16a34a"
                elif cpk_val >= 1.33:
                    durum = "✅ İyi"
                    durum_renk = "#22c55e"
                elif cpk_val >= 1.00:
                    durum = "⚠️ Kabul Edilebilir"
                    durum_renk = "#d97706"
                else:
                    durum = "❌ Yetersiz"
                    durum_renk = "#dc2626"
                
                hafta_cpk_data.append({
                    'Hafta': f"{int(hafta)}.HAFTA",
                    'Veri Sayısı': len(hafta_data),
                    'Ortalama': round(hafta_metrics['mean'], 4),
                    'Std Sapma': round(hafta_metrics['sigma_within'], 4),
                    'Cp': round(cp_val, 3) if cp_val else None,
                    'Cpk': round(cpk_val, 3) if cpk_val else None,
                    'Durum': durum
                })
    
    if hafta_cpk_data:
        cpk_df = pd.DataFrame(hafta_cpk_data)
        
        # Tablo ve grafik yan yana
        tbl_col, chart_col = st.columns([1.2, 1.8])
        
        with tbl_col:
            st.dataframe(cpk_df, use_container_width=True, height=400)
        
        with chart_col:
            # Cpk trend grafiği
            fig_cpk_trend = go.Figure()
            
            # Cpk çizgisi
            fig_cpk_trend.add_trace(go.Scatter(
                x=cpk_df['Hafta'], 
                y=cpk_df['Cpk'], 
                mode='lines+markers+text',
                name='Cpk',
                line=dict(color=COLORS['primary'], width=3),
                marker=dict(size=10),
                text=[f"{v:.2f}" if v else "" for v in cpk_df['Cpk']],
                textposition="top center"
            ))
            
            # Referans çizgileri
            fig_cpk_trend.add_hline(y=1.67, line_color=COLORS['success'], line_dash='dash', 
                                     annotation_text="Mükemmel (1.67)")
            fig_cpk_trend.add_hline(y=1.33, line_color='#22c55e', line_dash='dash', 
                                     annotation_text="İyi (1.33)")
            fig_cpk_trend.add_hline(y=1.00, line_color=COLORS['warning'], line_dash='dash', 
                                     annotation_text="Kabul (1.00)")
            
            # Bölge renklendirme
            fig_cpk_trend.add_hrect(y0=1.67, y1=3, fillcolor="rgba(22, 163, 74, 0.1)", line_width=0)
            fig_cpk_trend.add_hrect(y0=1.33, y1=1.67, fillcolor="rgba(34, 197, 94, 0.1)", line_width=0)
            fig_cpk_trend.add_hrect(y0=1.00, y1=1.33, fillcolor="rgba(217, 119, 6, 0.1)", line_width=0)
            fig_cpk_trend.add_hrect(y0=0, y1=1.00, fillcolor="rgba(220, 38, 38, 0.1)", line_width=0)
            
            fig_cpk_trend.update_layout(
                title=f"📈 {sec_kesit} - Haftalık Cpk Trendi",
                xaxis_title="Hafta",
                yaxis_title="Cpk Değeri",
                height=400,
                template="plotly_white",
                yaxis=dict(range=[0, max(2.5, cpk_df['Cpk'].max() * 1.2 if cpk_df['Cpk'].max() else 2.5)])
            )
            
            st.plotly_chart(fig_cpk_trend, use_container_width=True)
        
        # Özet istatistikler
        valid_cpk = [x for x in cpk_df['Cpk'].tolist() if x is not None]
        if valid_cpk:
            sum_col1, sum_col2, sum_col3, sum_col4 = st.columns(4)
            sum_col1.metric("Ortalama Cpk", f"{np.mean(valid_cpk):.3f}")
            sum_col2.metric("Min Cpk", f"{min(valid_cpk):.3f}")
            sum_col3.metric("Max Cpk", f"{max(valid_cpk):.3f}")
            sum_col4.metric("Toplam Hafta", len(valid_cpk))
    else:
        st.warning("Haftalık Cpk analizi için yeterli veri bulunamadı.")

# Rehberler (Eğitim İçeriği)
st.markdown("---")
with st.expander("📚 SPC Rehberi & Formüller"):
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("""
        **Kontrol Limitleri:**
        * UCL = X̄ + 3σ
        * LCL = X̄ - 3σ
        """)
    with col2:
        st.markdown("""
        **Yeterlilik:**
        * Cp = (USL - LSL) / 6σ
        * Cpk = min(CPU, CPL)
        """)

st.caption(f"Gelişmiş SPC Sistemi v4.1 | {SIRKET_ISMI}")
