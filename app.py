import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="E-Ticaret Konsolide Paneli v11.2", layout="wide")
st.title("🤖 Çok Kanallı E-Ticaret Konsolide Finans Paneli v11.2")
st.markdown("Amazon maliyet şablonundan doğrudan veri çeken, sipariş adetli ve kesin hesaplamalı kararlı sürüm.")
st.write("---")

# Session State Hazırlığı (KeyError Koruması)
varsayilanlar = {
    'ty_ciro': 0.0, 'ty_kesinti': 0.0, 'ty_maliyet': 0.0, 'ty_kar': 0.0, 'ty_sip_adet': 0, 'ty_urun_adet': 0,
    'amz_ciro': 0.0, 'amz_kesinti': 0.0, 'amz_maliyet': 0.0, 'amz_kar': 0.0, 'amz_sip_adet': 0, 'amz_urun_adet': 0,
    'genel_ciro': 0.0, 'genel_maliyet': 0.0, 'genel_kar': 0.0, 'genel_sip_adet': 0, 'genel_urun_adet': 0, 'genel_marj': 0.0,
    'hesaplandi': False
}
for k, v in varsayilanlar.items():
    if k not in st.session_state: st.session_state[k] = v

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
    with col_amz1: amazon_file = st.file_uploader("Haziran Amazon Raporu (Opsiyonel)", type=["xlsx", "xls", "csv"], key="amazon_sales")
    with col_amz2: amazon_maliyet_file = st.file_uploader("Amazon Maliyet Şablonu (Zorunlu veriler buradan okunur)", type=["xlsx", "xls", "csv"], key="amazon_cost")
    st.write("---")
    baslat_btn = st.button("🚀 Tüm Pazaryerlerinin Akıllı Analizini Başlat", use_container_width=True)

def clean_num(val):
    if pd.isnull(val): return 0.0
    if isinstance(val, (int, float)): return float(val)
    val_str = str(val).strip().replace('.', '').replace(',', '.')
    try: return float(val_str)
    except: return 0.0

if baslat_btn:
    st.session_state['hesaplandi'] = False
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
                    't_adet': clean_num(row['Ürün Adedi']),
                    'kom': clean_num(row['Komisyon/Yurt Dışı Stok Destek Bedeli']),
                    'kar': clean_num(row['Gönderi Kargo Bedeli']),
                    'hiz': clean_num(row['Platform Hizmet Bedeli'])
                }
            maliyet_dict = dict(zip(df_maliyet['TRENDYOL BARKOD'].astype(str).str.strip(), df_maliyet['TOPLAM MALİYET']))
            ty_sonuc = []
            for idx, row in df_prod.iterrows():
                barkod = str(row['Barkod']).strip()
                siparis_no = str(row['Sipariş Numarası']).strip()
                statü = str(row.get('Statü', row.get('Sipariş Durumu', ''))).strip().lower()
                adet = clean_num(row['Adet'])
                satis_tutari = clean_num(row['Satış Tutarı'])
                if barkod == 'nan' or siparis_no == 'nan' or "iptal" in statü or "reddedildi" in statü: continue
                ty_urun_adet += int(adet)
                birim_maliyet = clean_num(maliyet_dict.get(barkod, 0.0))
                toplam_maliyet = birim_maliyet * adet
                is_iade = "iade" in statü
                h_ciro = 0.0 if is_iade else satis_tutari
                h_maliyet = 0.0 if is_iade else toplam_maliyet
                
                f = finans_dict.get(siparis_no, {'t_adet': 0, 'kom': 0, 'kar': 0, 'hiz': 0})
                div = f['t_adet'] if f['t_adet'] > 0 else 1
                b_kom = f['kom'] / div * adet if f['t_adet'] > 0 else 0
                b_kar = f['kar'] / div * adet if f['t_adet'] > 0 else 0
                b_hiz = f['hiz'] / div * adet if f['t_adet'] > 0 else 0
                
                net_kar = h_ciro + b_kom + b_kar + b_hiz - h_maliyet
                ty_sonuc.append({"Ciro": h_ciro, "Kesinti": b_kom + b_kar + b_hiz, "Maliyet": h_maliyet, "Net Kâr": net_kar})
            df_ty = pd.DataFrame(ty_sonuc)
            ty_ciro, ty_kesinti, ty_maliyet_gideri = df_ty['Ciro'].sum(), abs(df_ty['Kesinti'].sum()), df_ty['Maliyet'].sum()
            ty_kar = df_ty['Net Kâr'].sum() - ty_reklam
        except Exception as e: st.error(f"Trendyol Hatası: {str(e)}")

    # 💛 AMAZON MOTORU (Hassas Düzenlenmiş Yeni Yapı)
    if amazon_maliyet_file:
        try:
            df_amz_cost = pd.read_csv(amazon_maliyet_file) if hasattr(amazon_maliyet_file, 'name') and amazon_maliyet_file.name.endswith('.csv') else pd.read_excel(amazon_maliyet_file)
            df_amz_cost.columns = [c.strip() for c in df_amz_cost.columns]
            
            # Sütun isimlerini esnek yakalama
            birim_col = 'Satilan_Net_Birim' if 'Satilan_Net_Birim' in df_amz_cost.columns else df_amz_cost.columns[1]
            satis_col = 'Brut_Satis' if 'Brut_Satis' in df_amz_cost.columns else df_amz_cost.columns[2]
            kazanc_col = 'Amazon_Net_Kazanc' if 'Amazon_Net_Kazanc' in df_amz_cost.columns else df_amz_cost.columns[3]
            maliyet_col = 'Birim Alış Maliyeti (₺)' if 'Birim Alış Maliyeti (₺)' in df_amz_cost.columns else df_amz_cost.columns[4]
            
            amz_sonuc = []
            for idx, row in df_amz_cost.iterrows():
                net_birim = clean_num(row.get(