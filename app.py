import streamlit as st
import pandas as pd
import plotly.express as px
import io

# Sayfa Genişlik ve Başlık Ayarları
st.set_page_config(page_title="Konsolide E-Ticaret Yapay Zeka Paneli v7.2", layout="wide")

st.title("🤖 Çok Kanallı (Trendyol & Amazon) Akıllı Yapay Zeka Paneli v7.2")
st.markdown("Amazon'daki bozuk kuruş, nokta ve virgül işaretleri tamamen temizlenmiştir.")
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
        
    ty_reklam = st.number_input("🔗 Varsa Trendyol Bu Aya Ait Toplam Ekstra Reklam Gideri (TL):", min_value=0.0, value=0.0, step=100.0)

    st.write("---")
    st.subheader("2. Amazon Raporları")
    col_amz1, col_amz2 = st.columns(2)
    with col_amz1:
        amazon_file = st.file_uploader("Haziran Amazon veya Amazon Satış Raporu", type=["xlsx", "xls", "csv"], key="amazon_sales")
    with col_amz2:
        amazon_maliyet_file = st.file_uploader("Amazon_Haziran_Maliyet_Sablonu dosyası", type=["xlsx", "xls", "csv"], key="amazon_cost")

    st.write("---")
    baslat_btn = st.button("🚀 Tüm Pazaryerlerinin Akıllı Analizini Başlat", use_container_width=True)

# 🧙‍♂️ GELİŞMİŞ SAYI TEMİZLEME SİHİRBAZI (Saçma Rakamların Panzehiri)
def clean_amazon_number(val):
    if pd.isnull(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    if hasattr(val, 'strftime'): 
        return 0.0
    
    try:
        val_str = str(val).strip()
        # Eğer hem nokta hem virgül varsa (Örn: 1,15.11400 veya 1.151,14 gibi)
        if ',' in val_str and '.' in val_str:
            # Son işarete bakarak kuruş ayracını bulalım
            if val_str.rfind(',') > val_str.rfind('.'):
                # Noktalar binlik ayracıdır, sil. Virgülü noktaya çevir.
                val_str = val_str.replace('.', '').replace(',', '.')
            else:
                # Virgüller binlik ayracıdır, sil.
                val_str = val_str.replace(',', '')
        else:
            # Sadece tek bir ayraç varsa kuruşa göre çevir
            val_str = val_str.replace(',', '.')
            
        # Çift nokta kalmışsa temizle
        if val_str.count('.') > 1:
            parts = val_str.split('.')
            val_str = "".join(parts[:-1]) + "." + parts[-1]
            
        return float(val_str)
    except:
        return 0.0

if baslat_btn:
    st.session_state['hesaplandi'] = False
    
    # -----------------------------
    # TRENDYOL ANALİZ MOTORU
    # -----------------------------
    ty_aktif = False
    ty_ciro, ty_kesinti, ty_hakedis, ty_kar, ty_iptal_iade = 0.0, 0.0, 0.0, 0.0, 0
    df_ty_final = pd.DataFrame()
    
    if finans_file and prod_file and maliyet_file:
        try:
            df_finans = pd.read_excel(finans_file)
            df_prod = pd.read_excel(prod_file, skiprows=1)
            df_maliyet = pd.read_excel(maliyet_file)
            
            df_maliyet.columns = [c.strip() for c in df_maliyet.columns]
            df_finans.columns = [c.strip() for c in df_finans.columns]
            df_prod.columns = [c.strip() for c in df_prod.columns]
            
            finans_dict = {}
            for idx, row in df_finans.iterrows():
                s_no = str(row['Sipariş No']).strip()
                finans_dict[s_no] = {
                    'toplam_adet': clean_amazon_number(row['Ürün Adedi']),
                    'komisyon': clean_amazon_number(row['Komisyon/Yurt Dışı Stok Destek Bedeli']),
                    'kargo': clean_amazon_number(row['Gönderi Kargo Bedeli']),
                    'hizmet': clean_amazon_number(row['Platform Hizmet Bedeli'])
                }
                
            maliyet_dict = dict(zip(df_maliyet['TRENDYOL BARKOD'].astype(str).str.strip(), df_maliyet['TOPLAM MALİYET']))
            ty_sonuc = []
            
            for idx, row in df_prod.iterrows():
                barkod = str(row['Barkod']).strip()
                siparis_no = str(row['Sipariş Numarası']).strip()
                statü = str(row.get('Statü', row.get('Sipariş Durumu', ''))).strip().lower()
                adet = clean_amazon_number(row['Adet'])
                satis_tutari = clean_amazon_number(row['Satış Tutarı'])
                
                if barkod == 'nan' or siparis_no == 'nan':
                    continue
                
                if "iptal" in statü or "iade" in statü or "reddedildi" in statü:
                    ty_iptal_iade += adet
                
                if "iptal" in statü or "reddedildi" in statü:
                    continue
                
                birim_maliyet = clean_amazon_number(maliyet_dict.get(barkod, 0.0))
                toplam_maliyet = birim_maliyet * adet
                is_iade = "iade" in statü
                
                hesaplanan_ciro = 0.0 if is_iade else satis_tutari
                hesaplanan_maliyet = 0.0 if is_iade else toplam_maliyet
                
                bolunmus_komisyon, bolunmus_kargo, bolunmus_hizmet = 0.0, 0.0, 0.0
                if siparis_no in finans_dict:
                    f_data = finans_dict[siparis_no]
                    if f_data['toplam_adet'] > 0:
                        bolunmus_komisyon = (f_data['komisyon'] / f_data['toplam_adet']) * adet
                        bolunmus_kargo = (f_data['kargo'] / f_data['toplam_adet']) * adet
                        bolunmus_hizmet = (f_data['hizmet'] / f_data['toplam_adet']) * adet
                
                net_kar = hesaplanan_ciro + bolunmus_komisyon + bolunmus_kargo + bolunmus_hizmet - hesaplanan_maliyet
                
                ty_sonuc.append({
                    "Pazaryeri": "Trendyol",
                    "Sipariş No/ASIN": siparis_no,
                    "Barkod": barkod,
                    "Ürün Adı": row['Ürün Adı'],
                    "Durum": "İade" if is_iade else "Satış",
                    "Adet": adet,
                    "Ciro": hesaplanan_ciro,
                    "Kesinti": bolunmus_komisyon + bolunmus_kargo + bolunmus_hizmet,
                    "Maliyet": hesaplanan_maliyet,
                    "Net Kâr": net_kar
                })
            
            df_ty_final = pd.DataFrame(ty_sonuc)
            ty_ciro = df_ty_final['Ciro'].sum()
            ty_kesinti = abs(df_ty_final['Kesinti'].sum())
            ty_hakedis = ty_ciro - ty_kesinti
            ty_kar = df_ty_final['Net Kâr'].sum() - ty_reklam
            ty_aktif = True
            
        except Exception as e:
            st.error(f"Trendyol Veri İşleme Hatası: {str(e)}")

    # -----------------------------
    # AMAZON ANALİZ MOTORU
    # -----------------------------
    amz_aktif = False
    amz_ciro, amz_kesinti, amz_hakedis, amz_kar, amz_iptal_iade = 0.0, 0.0, 0.0, 0.0, 0
    df_amz_final = pd.DataFrame()
    
    if amazon_file and amazon_maliyet_file:
        try:
            if hasattr(amazon_file, 'name') and amazon_file.name.endswith('.csv'):
                df_amz_sales = pd.read_csv(amazon_file)
            else:
                df_amz_sales = pd.read_excel(amazon_file)
                
            if hasattr(amazon_maliyet_file, 'name') and amazon_maliyet_file.name.endswith('.csv'):
                df_amz_cost = pd.read_csv(amazon_maliyet_file)
            else:
                df_amz_cost = pd.read_excel(amazon_maliyet_file)
                
            df_amz_sales.columns = [c.strip() for c in df_amz_sales.columns]
            df_amz_cost.columns = [c.strip() for c in df_amz_cost.columns]
            
            asin_col = 'Ana ürün ASIN\'i' if 'Ana ürün ASIN\'i' in df_amz_cost.columns else 'ASIN'
            cost_col = 'Birim Alış Maliyetleri (₺)' if 'Birim Alış Maliyetleri (₺)' in df_amz_cost.columns else df_amz_cost.columns[-2]
            amz_cost_dict = dict(zip(df_amz_cost[asin_col].astype(str).str.strip(), df_amz_cost[cost_col]))
            
            amz_sonuc = []
            for idx, row in df_amz_sales.iterrows():
                asin = str(row