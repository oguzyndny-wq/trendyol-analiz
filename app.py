import streamlit as st
import pandas as pd
import plotly.express as px
import io

# Sayfa Genişlik ve Başlık Ayarları
st.set_page_config(page_title="Konsolide E-Ticaret Yapay Zeka Paneli v7.1", layout="wide")

st.title("🤖 Çok Kanallı (Trendyol & Amazon) Akıllı Yapay Zeka Paneli v7.1")
st.markdown("Amazon veri tipi uyuşmazlıkları ve tarih dönüşüm hataları giderilmiştir.")
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
        
    ty_reklam = st.number_input("🔗 Varsa Trendyol Bu Aya Ait Toplam Ekstra Reklam Gider Mi? (TL):", min_value=0.0, value=0.0, step=100.0)

    st.write("---")
    st.subheader("2. Amazon Raporları")
    col_amz1, col_amz2 = st.columns(2)
    with col_amz1:
        amazon_file = st.file_uploader("Haziran Amazon veya Amazon Satış Raporu", type=["xlsx", "xls", "csv"], key="amazon_sales")
    with col_amz2:
        amazon_maliyet_file = st.file_uploader("Amazon_Haziran_Maliyet_Sablonu dosyası", type=["xlsx", "xls", "csv"], key="amazon_cost")

    st.write("---")
    baslat_btn = st.button("🚀 Tüm Pazaryerlerinin Akıllı Analizini Başlat", use_container_width=True)

# GÜVENLİ SAYI DÖNÜŞTÜRÜCÜ MOTOR (Hatanın Panzehiri)
def safe_float(val):
    if pd.isnull(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    # Eğer hücreye yanlışlıkla tarih objesi gelmişse 0.0 dön ki çökmesin
    if hasattr(val, 'strftime'): 
        return 0.0
    try:
        return float(str(val).strip().replace(',', '.'))
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
                    'toplam_adet': safe_float(row['Ürün Adedi']),
                    'komisyon': safe_float(row['Komisyon/Yurt Dışı Stok Destek Bedeli']),
                    'kargo': safe_float(row['Gönderi Kargo Bedeli']),
                    'hizmet': safe_float(row['Platform Hizmet Bedeli'])
                }
                
            maliyet_dict = dict(zip(df_maliyet['TRENDYOL BARKOD'].astype(str).str.strip(), df_maliyet['TOPLAM MALİYET']))
            ty_sonuc = []
            
            for idx, row in df_prod.iterrows():
                barkod = str(row['Barkod']).strip()
                siparis_no = str(row['Sipariş Numarası']).strip()
                statü = str(row.get('Statü', row.get('Sipariş Durumu', ''))).strip().lower()
                adet = safe_float(row['Adet'])
                satis_tutari = safe_float(row['Satış Tutarı'])
                
                if barkod == 'nan' or siparis_no == 'nan':
                    continue
                
                if "iptal" in statü or "iade" in statü or "reddedildi" in statü:
                    ty_iptal_iade += adet
                
                if "iptal" in statü or "reddedildi" in statü:
                    continue
                
                birim_maliyet = safe_float(maliyet_dict.get(barkod, 0.0))
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
                asin = str(row['Ana ürün ASIN\'i']).strip()
                satilan_birim = safe_float(row.get('Satılan birimler', 0))
                iade_birim = safe_float(row.get('İade edilen birimler', 0))
                net_birim = safe_float(row.get('Satılan net birim sayısı', 0))
                
                satis_tutari = safe_float(row.get('Satış', 0.0))
                net_kazanc = safe_float(row.get('Toplam Net kazanç', 0.0))
                
                amz_iptal_iade += iade_birim
                
                birim_maliyet = safe_float(amz_cost_dict.get(asin, 0.0))
                toplam_maliyet = birim_maliyet * max(0.0, net_birim)
                
                kesinti = satis_tutari - net_kazanc
                net_kar = net_kazanc - toplam_maliyet
                
                amz_sonuc.append({
                    "Pazaryeri": "Amazon",
                    "Sipariş No/ASIN": asin,
                    "Barkod": asin,
                    "Ürün Adı": f"Amazon ASIN: {asin}",
                    "Durum": "Net Satış",
                    "Adet": net_birim,
                    "Ciro": satis_tutari,
                    "Kesinti": -kesinti,
                    "Maliyet": toplam_maliyet,
                    "Net Kâr": net_kar
                })
                
            df_amz_final = pd.DataFrame(amz_sonuc)
            amz_ciro = df_amz_final['Ciro'].sum()
            amz_hakedis = df_amz_final['Ciro'].sum() + df_amz_final['Kesinti'].sum() # İçerideki hakediş hesabı
            amz_kesinti = amz_ciro - amz_hakedis
            amz_kar = df_amz_final['Net Kâr'].sum()
            amz_aktif = True
            
        except Exception as e:
            st.error(f"Amazon Veri İşleme Hatası: {str(e)}")

    # -----------------------------
    # SESSION STATE KORUMA
    # -----------------------------
    if ty_aktif or amz_aktif:
        st.session_state['ty_ciro'] = ty_ciro
        st.session_state['ty_kesinti'] = ty_kesinti
        st.session_state['ty_hakedis'] = ty_hakedis
        st.session_state['ty_kar'] = ty_kar
        st.session_state['ty_iptal_iade'] = ty_iptal_iade
        
        st.session_state['amz_ciro'] = amz_ciro
        st.session_state['amz_kesinti'] = amz_kesinti
        st.session_state['amz_hakedis'] = amz_hakedis
        st.session_state['amz_kar'] = amz_kar
        st.session_state['amz_iptal_iade'] = amz_iptal_iade
        
        st.session_state['genel_ciro'] = ty_ciro + amz_ciro
        st.session_state['genel_kesinti'] = ty_kesinti + amz_kesinti
        st.session_state['genel_hakedis'] = ty_hakedis + amz_hakedis
        st.session_state['genel_kar'] = ty_kar + amz_kar
        st.session_state['genel_iptal_iade'] = ty_iptal_iade + amz_iptal_iade
        st.session_state['genel_marj'] = (st.session_state['genel_kar'] / st.session_state['genel_ciro'] * 100) if st.session_state['genel_ciro'] > 0 else 0.0
        
        frames = []
        if not df_ty_final.empty: frames.append(df_ty_final)
        if not df_amz_final.empty: frames.append(df_amz_final)
        if frames:
            st.session_state['df_konsolide'] = pd.concat(frames, ignore_index=True)
            
        st.session_state['hesaplandi'] = True
        st.success("✅ Çok kanallı analiz başarıyla tamamlandı!")

# EKRAN ÇİZİM ALANI
if st.session_state.get('hesaplandi', False):
    with tab_rapor:
        st.subheader("👑 Genel Konsolide (Şinter Şirket Toplamı) Durum Masası")
        g1, g2, g3, g4 = st.columns(4)
        g1.metric("💰 Toplam Şirket Cirosu", f"₺{st.session_state['genel_ciro']:,.2f}")
        g2.metric("🟢 Toplam Net Kâr", f"₺{st.session_state['genel_kar']:,.2f}")
        g3.metric("📈 Genel Net Kâr Marjı", f"%{st.session_state['genel_marj']:.2f}")
        g4.metric("🚨 Toplam İptal / İade Adedi", f"{int(st.session_state['genel_iptal_iade'])} Adet")
        
        st.write("---")
        st.subheader("📊 Pazaryerlerine Göre Durum Kırılımı")
        col_ty_pan, col_amz_pan = st.columns(2)
        
        with col_ty_pan:
            st.markdown("### 🧡 Trendyol Performansı")
            st.write(f"**Net Ciro:** ₺{st.session_state['ty_ciro']:,.2f}")
            st.write(f"**Trendyol Kesintileri:** ₺{st.session_state['ty_kesinti']:,.2f}")
            st.write(f"**Hesaba Giren (Hakediş):** ₺{st.session_state['ty_hakedis']:,.2f}")
            st.write(f"**Trendyol Net Kâr:** ₺{st.session_state['ty_kar']:,.2f}")
            
        with col_amz_pan:
            st.markdown("### 💛 Amazon Performansı")
            st.write(f"**Net Ciro:** ₺{st.session_state['amz_ciro']:,.2f}")
            st.write(f"**Amazon Kesintileri:** ₺{st.session_state['amz_kesinti']:,.2f}")
            st.write(f"**Hesaba Giren (Hakediş):** ₺{st.session_state['amz_hakedis']:,.2f}")
            st.write(f"**Amazon Net Kâr:** ₺{st.session_state['amz_kar']:,.2f}")

        st.write("---")
        df_g = pd.DataFrame({
            "Pazaryeri": ["Trendyol", "Amazon"],
            "Ciro": [st.session_state['ty_ciro'], st.session_state['amz_ciro']]
        })
        st.plotly_chart(px.pie(df_g, values='Ciro', names='Pazaryeri', title="Ciro Dağılımı (%)", hole=0.3), use_container_width=True)
        
        if 'df_konsolide' in st.session_state:
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                st.session_state['df_konsolide'].to_excel(writer, index=False, sheet_name='Konsolide_Satis_Raporu')
            st.download_button(label="📥 Tüm Şirket Satışlarını Konsolide Excel Olarak İndir", data=output.getvalue(), file_name="Konsolide_E-Ticaret_Raporu.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)