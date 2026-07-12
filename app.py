import streamlit as st
import pandas as pd
import plotly.express as px
import io

# Sayfa Genişlik ve Başlık Ayarları (İnternet Sitesi Tasarımı)
st.set_page_config(page_title="Trendyol Kâr Analiz Paneli", layout="wide")

st.title("📊 Trendyol Gelişmiş Kâr Analiz Paneli")
st.markdown("Trendyol raporlarınızı yükleyin, sistem kargo ve komisyonları otomatik bölüştürerek gerçek kârlılığınızı çıkarsın.")
st.write("---")

# 1. DOSYA YÜKLEME ALANLARI (Arkadaşınızın Ekranının Aynısı)
col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("📊 Sipariş Kayıtları")
    finans_file = st.file_uploader("SiparisKayitlari ile başlayan dosya", type=["xlsx", "xls"], key="finans")

with col2:
    st.subheader("📦 Sipariş Listesi")
    prod_file = st.file_uploader("prod_ ile başlayan dosya", type=["xlsx", "xls"], key="prod")

with col3:
    st.subheader("🏷️ Ürün Alış Maliyetleri")
    maliyet_file = st.file_uploader("Maliyet listesi dosyası", type=["xlsx", "xls"], key="maliyet")

st.write("---")

# Reklam Gideri Giriş Alanı (İstediğiniz Ekstra Özellik)
reklam_gideri = st.number_input("🔗 Varsa Bu Aya Ait Toplam Reklam Giderini Girin (TL):", min_value=0.0, value=0.0, step=100.0)

# ANALİZİ BAŞLAT BUTONU
if st.button("🚀 Analizi Başlat", use_container_width=True):
    if finans_file and prod_file and maliyet_file:
        with st.spinner("Hesaplamalar yapılıyor, kargolar bölüştürülüyor..."):
            
            # Verileri Oku
            df_finans = pd.read_excel(finans_file)
            # Prod siparişlerdeki 1. satır yasal uyarısını atla, 2. satırı başlık yap
            df_prod = pd.read_excel(prod_file, skiprows=1)
            df_maliyet = pd.read_excel(maliyet_file)
            
            # Sütun İsimlerini Temizle
            df_maliyet.columns = [c.strip() for c in df_maliyet.columns]
            df_finans.columns = [c.strip() for c in df_finans.columns]
            df_prod.columns = [c.strip() for c in df_prod.columns]
            
            # Finans Verilerini Sipariş No bazında haritala
            finans_dict = {}
            for idx, row in df_finans.iterrows():
                s_no = str(row['Sipariş No']).strip()
                finans_dict[s_no] = {
                    'toplam_adet': row['Ürün Adedi'],
                    'komisyon': row['Komisyon/Yurt Dışı Stok Destek Bedeli'],
                    'kargo': row['Gönderi Kargo Bedeli'],
                    'hizmet': row['Platform Hizmet Bedeli']
                }
                
            # Maliyet Haritası
            maliyet_dict = dict(zip(df_maliyet['TRENDYOL BARKOD'].astype(str).str.strip(), df_maliyet['TOPLAM MALİYET']))
            
            # Hesaplama Tablosunu Oluştur
            sonuc_listesi = []
            
            for idx, row in df_prod.iterrows():
                barkod = str(row['Barkod']).strip()
                siparis_no = str(row['Sipariş Numarası']).strip()
                urun_adi = row['Ürün Adı']
                marka = row['Marka']
                adet = float(row['Adet']) if pd.notnull(row['Adet']) else 0
                satis_tutari = float(row['Satış Tutarı']) if pd.notnull(row['Satış Tutarı']) else 0
                
                if barkod == 'nan' or siparis_no == 'nan':
                    continue
                    
                # Eşit Bölüştürme Motoru
                bolunmus_komisyon = 0
                bolunmus_kargo = 0
                bolunmus_hizmet = 0
                
                if siparis_no in finans_dict:
                    f_data = finans_dict[siparis_no]
                    if f_data['toplam_adet'] > 0:
                        bolunmus_komisyon = (f_data['komisyon'] / f_data['toplam_adet']) * adet
                        bolunmus_kargo = (f_data['kargo'] / f_data['toplam_adet']) * adet
                        bolunmus_hizmet = (f_data['hizmet'] / f_data['toplam_adet']) * adet
                
                # Maliyet Getir
                birim_maliyet = maliyet_dict.get(barkod, 0)
                toplam_maliyet = birim_maliyet * adet
                
                # Net Kâr (Trendyol kesintileri eksi değerde geldiği için matematiksel olarak topluyoruz)
                net_kar = satis_tutari + bolunmus_komisyon + bolunmus_kargo + bolunmus_hizmet - toplamMaliyet
                
                sonuc_listesi.append({
                    "Barkod": barkod,
                    "Marka": marka,
                    "Ürün Adı": urun_adi,
                    "Satış Adedi": adet,
                    "Ciro": satis_tutari,
                    "Komisyon": bolunmus_komisyon,
                    "Kargo": bolunmus_kargo,
                    "Hizmet Bedeli": bolunmus_hizmet,
                    "Ürün Maliyeti": toplam_maliyet,
                    "Net Kâr": net_kar
                })
                
            df_sonuc = pd.DataFrame(sonuc_listesi)
            
            # 📊 DASHBOARD METRİKLERİ
            toplam_ciro = df_sonuc['Ciro'].sum()
            toplam_kargo = df_sonuc['Kargo'].sum()
            toplam_komisyon = df_sonuc['Komisyon'].sum()
            toplam_hizmet = df_sonuc['Hizmet Bedeli'].sum()
            toplam_urun_maliyeti = df_sonuc['Ürün Maliyeti'].sum()
            
            # Net kârdan girilen reklam giderini düşüyoruz
            genel_net_kar = df_sonuc['Net Kâr'].sum() - reklam_gideri
            toplam_adet = df_sonuc['Satış Adedi'].sum()
            
            # İnternet Sitesi Özet Kartları
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("💰 Toplam Ciro", f"₺{toplam_ciro:,.2f}")
            m2.metric("🟢 Gerçek Net Kâr (Reklam Dahil)", f"₺{genel_net_kar:,.2f}")
            m3.metric("📦 Toplam Satış Adedi", f"{int(toplam_adet)} Adet")
            m4.metric("🚚 Toplam Kargo Maliyeti", f"₺{abs(toplam_kargo):,.2f}")
            
            # 📈 GRAFİKLER VE MARKA ANALİZLERİ
            st.write("### 🏢 Marka Bazlı Kârlılık Analizi")
            df_marka = df_sonuc.groupby('Marka').agg({'Ciro':'sum', 'Net Kâr':'sum', 'Satış Adedi':'sum'}).reset_index()
            st.dataframe(df_marka.style.format({'Ciro': '₺{:.2f}', 'Net Kâr': '₺{:.2f}'}), use_container_width=True)
            
            fig = px.bar(df_marka, x='Marka', y='Net Kâr', title="Markaların Net Kâr Dağılımı", color='Marka')
            st.plotly_chart(fig, use_container_width=True)
            
            # Excel Olarak İndirme Butonu
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_sonuc.to_excel(writer, index=False, sheet_name='Kâr Analiz Sonucu')
            processed_data = output.getvalue()
            
            st.download_button(
                label="📥 Tüm Sonuçları Detaylı Excel Olarak İndir",
                data=processed_data,
                file_name="Trendyol_Detayli_Kar_Analizi.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
    else:
        st.error("Lütfen analiz için 3 dosyayı da eksiksiz yükleyin!")