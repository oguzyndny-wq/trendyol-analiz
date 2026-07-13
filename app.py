import streamlit as st
import pandas as pd
import plotly.express as px
import io

st.set_page_config(page_title="E-Ticaret Konsolide Paneli v10.4", layout="wide")
st.title("🤖 Çok Kanallı E-Ticaret Konsolide Finans Paneli v10.4")
st.markdown("Görsel geri bildirimli ve dinamik sekme uyarılı tam stabil versiyon.")
st.write("---")

# Session State Değişkenlerinin Önceden Güvenli Tanımlanması
varsayilanlar = {
    'ty_ciro': 0.0, 'ty_kesinti': 0.0, 'ty_maliyet': 0.0, 'ty_kar': 0.0, 'ty_sip_adet': 0, 'ty_urun_adet': 0,
    'amz_ciro': 0.0, 'amz_kesinti': 0.0, 'amz_maliyet': 0.0, 'amz_kar': 0.0, 'amz_sip_adet': 0, 'amz_urun_adet': 0,
    'genel_ciro': 0.0, 'genel_maliyet': 0.0, 'genel_kar': 0.0, 'genel_sip_adet': 0, 'genel_urun_adet': 0, 'genel_marj': 0.0,
    'hesaplandi': False
}
for k, v in varsayilanlar.items():
    if k not in st.session_state:
        st.session_state[k] = v

tab_yükleme, tab_rapor = st.tabs(["📥 Veri Yükleme İstasyonu", "📊 Konsolide Finans & Raporlar"])

with tab_yükleme:
    st.subheader("1. Trendyol Raporları")
    col1, col2, col3 = st.columns(3)
    with col1: finans_file = st.file_uploader("SiparisKayitlari ile başlayan dosya", type=["xlsx", "xls"], key="finans")
    with col2: prod_file = st.file_uploader("prod_ ile başlayan dosya", type=["xlsx", "xls"], key="prod")
    with col3: maliyet_file = st.file_uploader("Trendyol Maliyet listesi", type=["xlsx", "xls"], key="maliyet")
    ty_reklam = st.number_input("🔗 Varsa Trendyol Ekstra Reklam Gideri (TL):", min_value=0.0, value=0.0, step=100.0)
    st.write("---")
    st.subheader("2. Amazon Raporları")
    col_amz1, col_amz2 = st.columns(2)
    with col_amz1: amazon_file = st.file_uploader("Haziran Amazon Raporu", type=["xlsx", "xls", "csv"], key="amazon_sales")
    with col_amz2: amazon_maliyet_file = st.file_uploader("Amazon Maliyet Şablonu", type=["xlsx", "xls", "csv"], key="amazon_cost")
    st.write("---")
    
    baslat_btn = st.button("🚀 Tüm Pazaryerlerinin Akıllı Analizini Başlat", use_container_width=True)

def clean_number(val):
    if pd.isnull(val): return 0.0
    if isinstance(val, (int, float)): return float(val)
    if hasattr(val, 'strftime'): return 0.0
    val_str = str(val).strip()
    if '-' in val_str and len(val_str) > 7:
        p = val_str.split('-')
        try: return float(p[0]) + float(p[1])/100.0
        except: return 0.0
    try:
        if ',' in val_str and '.' in val_str:
            if val_str.rfind(',') > val_str.rfind('.'): val_str = val_str.replace('.', '').replace(',', '.')
            else: val_str = val_str.replace(',', '')
        else: val_str = val_str.replace(',', '.')
        if val_str.count('.') > 1:
            p = val_str.split('.')
            val_str = "".join(p[:-1]) + "." + p[-1]
        return float(val_str)
    except: return 0.0

def clean_amazon_net_kazanc(val):
    if pd.isnull(val): return 0.0
    val_str = str(val).strip()
    if '-' in val_str and len(val_str) > 7: return 0.0
    try:
        num = float(val_str.replace(',', '.'))
        return num / 10000.0 if ('.' not in val_str and abs(num) > 100000) else num
    except: return 0.0

if baslat_btn:
    st.session_state['hesaplandi'] = False
    ty_aktif, amz_aktif = False, False
    ty_ciro, ty_kesinti, ty_maliyet_gideri, ty_kar, ty_sip_adet, ty_urun_adet = 0.0, 0.0, 0.0, 0.0, 0, 0
    amz_ciro, amz_kesinti, amz_maliyet_gideri, amz_kar, amz_sip_adet, amz_urun_adet = 0.0, 0.0, 0.0, 0.0, 0, 0
    
    # 🧡 TRENDYOL MOTORU
    if finans_file and prod_file and maliyet_file:
        try:
            df_finans = pd.read_excel(finans_file)
            df_prod = pd.read_excel(prod_file, skiprows=1)
            df_maliyet = pd.read_excel(maliyet_file)
            df_maliyet.columns = [c.strip() for c in df_maliyet.columns]
            df_finans.columns = [c.strip() for c in df_finans.columns]
            df_prod.columns = [c.strip() for c in df_prod.columns]
            ty_sip_adet = int(df_prod['Sipariş Numarası'].nunique())
            
            finans_dict = {}
            for idx, row in df_finans.iterrows():
                s_no = str(row['Sipariş No']).strip()
                finans_dict[s_no] = {
                    't_adet': clean_number(row['Ürün Adedi']),
                    'kom': clean_number(row['Komisyon/Yurt Dışı Stok Destek Bedeli']),
                    'kar': clean_number(row['Gönderi Kargo Bedeli']),
                    'hiz': clean_number(row['Platform Hizmet Bedeli'])
                }
            maliyet_dict = dict(zip(df_maliyet['TRENDYOL BARKOD'].astype(str).str.strip(), df_maliyet['TOPLAM MALİYET']))
            ty_sonuc = []
            for idx, row in df_prod.iterrows():
                barkod = str(row['Barkod']).strip()
                siparis_no = str(row['Sipariş Numarası']).strip()
                statü = str(row.get('Statü', row.get('Sipariş Durumu', ''))).strip().lower()
                adet = clean_number(row['Adet'])
                satis_tutari = clean_number(row['Satış Tutarı'])
                if barkod == 'nan' or siparis_no == 'nan': continue
                if "iptal" in statü or "reddedildi" in statü: continue
                ty_urun_adet += int(adet)
                birim_maliyet = clean_number(maliyet_dict.get(barkod, 0.0))
                toplam_maliyet = birim_maliyet * adet
                is_iade = "iade" in statü
                h_ciro = 0.0 if is_iade else satis_tutari
                h_maliyet = 0.0 if is_iade else toplam_maliyet
                b_kom, b_kar, b_hiz = 0.0, 0.0, 0.0
                if siparis