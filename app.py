import streamlit as st
import pandas as pd
import plotly.express as px
import io

# Sayfa Genişlik ve Başlık Ayarları
st.set_page_config(page_title="Konsolide E-Ticaret Yönetim Paneli v8.5", layout="wide")

st.title("🤖 Çok Kanallı E-Ticaret Konsolide Finans Paneli v8.5")
st.markdown("Değişken eşitleme hatası giderilmiştir. Tüm veriler dosyalardan anlık hesaplanır.")
st.write("---")

# SEKME SİSTEMİ
tab_yükleme, tab_rapor = st.tabs(["📥 Veri Yükleme İstasyonu", "📊 Konsolide Finans & Raporlar"])

with tab_yükleme:
    st.subheader("1. Trendyol Raporları")
    col1, col2, col3 = st.columns(3)
    with col1:
        finans_file = st.file_uploader("SiparisKayitlari ile başlayan Trendyol dosyası", type=["xlsx", "xls"], key="finans")
    with col2:
        prod_file = st.file_uploader("prod_ ile başlayan Trendyol dosyası", type=["xlsx", "xls"], key="prod")
    with col3:
        maliyet_file = st.file_uploader("Trendyol Maliyet listesi dosyası", type=["xlsx", "xls"], key="maliyet")
        
    ty_reklam = st.number_input("🔗 Varsa Trendyol Ekstra Reklam Gideri (TL):", min_value=0.0, value=0.0, step=100.0)

    st.write("---")
    st.subheader("2. Amazon Raporları")
    col_amz1, col_amz2 = st.columns(2)
    with col_amz1:
        amazon_file = st.file_uploader("Haziran Amazon veya Amazon Satış Raporu", type=["xlsx", "xls", "csv"], key="amazon_sales")
    with col_amz2:
        amazon_maliyet_file = st.file_uploader("Amazon_Haziran_Maliyet_Sablonu dosyası", type=["xlsx", "xls", "csv"], key="amazon_cost")

    st.write("---")
    baslat_btn = st.button("🚀 Tüm Pazaryerlerinin Akıllı Analizini Başlat", use_container_width=True)

# GÜVENLİ SAYI VE TARİH AYRIŞTIRICI MOTOR
def clean_number(val):
    if pd.isnull(val): return 0.0
    if isinstance(val, (int, float)): return float(val)
    if hasattr(val, 'strftime'): return 0.0
    val_str = str(val).strip()
    if '-' in val_str and len(val_str) > 7:
        parts = val_str.split('-')
        try: return float(parts[0]) + float(parts[1])/100.0
        except: return 0.0
    try:
        if ',' in val_str and '.' in val_str:
            if val_str.rfind(',') > val_str.rfind('.'): val_str = val_str.replace('.', '').replace(',', '.')
            else: val_str = val_str.replace(',', '')
        else: val_str = val_str.replace(',', '.')
        if val_str.count('.') > 1:
            parts = val_str.split('.')
            val_str = "".join(parts[:-1]) + "." + parts[-1]
        return float(val_str)
    except: return 0.0

# AMAZON NET KAZANÇ KURUŞ DÜZELTİCİ
def clean_amazon_net_kazanc(val):
    if pd.isnull(val): return 0.0
    if isinstance(val, (int, float)):
        num = float(val)
        if '.' not in str(val) and abs(num) > 100000: return num / 10000.0
        return num
    val_str = str(val).strip()
    if val_str.startswith('-') and '-' in val_str[1:]: return 0.0
    try:
        num = float(val_str.replace(',', '.'))
        if '.' not in val_str && abs(num) > 100000: return num / 10000.0
        return num
    except: return 0.0

if baslat_btn:
    st.session_state['hesaplandi'] = False
    ty_aktif, amz_aktif = False, False
    
    # Sıfırlama Sayaçları (Hata Alan Bölüm Düzeltildi)
    ty_ciro, ty_kesinti, ty_maliyet_gideri, ty_kar = 0.0, 0.0, 0.0, 0.0
    ty_sip_adet, ty_urun_adet = 0, 0
    
    amz_ciro, amz_kesinti, amz_maliyet_gideri, amz_kar = 0.0, 0.0, 0.0, 0.0
    amz_sip_adet, amz_urun_adet = 0, 0
    
    df_ty_final, df_amz_final = pd.DataFrame(), pd.DataFrame()
    
    # 🧡 TRENDYOL HESAPLAMA MOTORU
    if finans_file and prod_file and maliyet_file:
        try:
            df_finans = pd.read_excel(finans_file)
            df_prod = pd.read_excel(prod_file, skiprows=1)
            df_maliyet = pd.read_excel(maliyet_file)