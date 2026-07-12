import streamlit as st
import pandas as pd
import plotly.express as px
import io

# Sayfa Genişlik ve Başlık Ayarları
st.set_page_config(page_title="Trendyol Kâr Analiz Paneli v3", layout="wide")

st.title("📊 Trendyol Gelişmiş Kâr Analiz Paneli v3 (Maliyet & İade Motorlu)")
st.markdown("Trendyol raporlarınızı yükleyin; sistem iade stoklarını, eksik maliyetli barkodları otomatik yönetsin.")
st.write("---")

# 1. DOSYA YÜKLEME ALANLARI
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

# Reklam Gideri Giriş Alanı
reklam_gideri = st.number_input("🔗 Varsa Bu Aya Ait Toplam Reklam Giderini Giriş Yapın (TL):", min_value=0.0, value=0.0, step=100.0)

# ANALİZİ BAŞLAT BUTONU
if st.button("🚀 Analizi Başlat", use_container_width=True):
    if finans_file and prod_file and maliyet_file:
        with st.spinner("Gelişmiş maliyet ve iade optimizasyonu hesaplanıyor..."):
            
            # Verileri Oku
            df_finans = pd.read_excel(finans_file)
            df_prod = pd.read_excel(prod_file, skiprows=1)
            df_maliyet = pd.read_excel(maliyet_file)
            
            # Sütun İsimlerini Temizle
            df_maliyet.columns = [c.strip() for c in df_maliyet.columns]
            df_finans.columns = [c.strip() for c in df_finans.columns]
            df_prod.columns = [c.strip() for c in df_prod.columns]
            
            # Finans Verilerini Haritala
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
            
            # Hesaplama Döngüsü
            sonuc_listesi = []
            eksik_maliyetler = set()
            
            for idx, row in df_prod.iterrows():
                barkod = str(row['Barkod']).strip()
                siparis_no = str(row['Sipariş Numarası']).strip()
                urun_adi = row['Ürün Adı']
                marka = row['Marka']
                statü = str(row.get('Statü', row.get('Sipariş Durumu', ''))).strip().lower()
                
                adet = float(row['Adet']) if pd.notnull(row['Adet']) else 0
                satis_tutari = float(row['Satış Tutarı']) if pd.notnull(row['Satış Tutarı']) else 0
                
                if barkod == 'nan' or siparis_no == 'nan':
                    continue
                
                # 1. KONTROL: İPTAL DURUMU (Müşteri ürünü hiç almadıysa tamamen devre dışı bırak)
                if "iptal" in statü or "reddedildi" in statü:
                    continue
                
                # Maliyet Bilgisini Çek
                if barkod in maliyet_dict:
                    birim_maliyet = maliyet_dict[barkod]
                else:
                    birim_maliyet = 0
                    eksik_maliyetler.add((barkod, urun_adi))
                
                toplam_maliyet = birim_maliyet * adet
                
                # 2. KONTROL: İADE DURUMU (Ürün geri geldiyse ciro sıfırlanır ama maliyet yükü kaldırılır)
                is_iade = "iade" in statü
                
                if is_iade:
                    # İade geldiyse bu satırdan ciro elde etmediniz (0 TL)
                    hesaplanan_ciro = 0
                    # Ürün rafa geri döndüğü için maliyet zararınız yoktur (0 TL)
                    hesaplanan_maliyet = 0
                else:
                    hesaplanan_ciro = satis_tutari
                    hesaplanan_maliyet = toplam_maliyet
                
                # Trendyol Finansal Kesintilerini Bölüştür
                bolunmus_komisyon = 0
                bolunmus_kargo = 0
                bolunmus_hizmet = 0
                
                if siparis_no in finans_dict:
                    f_data = finans_dict[siparis_no]
                    if f_data['toplam_adet'] > 0:
                        bolunmus_komisyon = (f_data['komisyon'] / f_data['toplam_adet']) * adet
                        bolunmus_kargo = (f_data['kargo'] / f_data['toplam_adet']) * adet
                        bolunmus_hizmet = (f_data['hizmet'] / f_data['toplam_adet']) * adet
                
                # Net Kâr Formülü (İade ise ciro 0, maliyet 0 olur; sadece kargo/komisyon cezası kalır)
                net_kar = hesaplanan_ciro + bolunmus_komisyon + bolunmus_kargo + bolunmus_hizmet - hesaplanan_maliyet
                
                sonuc_listesi.append({
                    "Sipariş No": siparis_no,
                    "Barkod": barkod,
                    "Marka": marka,
                    "Ürün Adı": urun_adi,
                    "Durum": "İade Edildi" if is_iade else "Satış",
                    "Satış Adedi": adet,
                    "Net Ciro": hesaplanan_ciro,
                    "Komisyon": bolunmus_komisyon,
                    "Kargo": bolunmus_kargo,
                    "Hizmet Bedeli": bolunmus_hizmet,
                    "Ürün Maliyeti": hesaplanan_maliyet,
                    "Net Kâr": net_kar
                })
                
            df_sonuc = pd.DataFrame(sonuc_listesi)
            
            # 📊 DASHBOARD METRİKLERİ
            toplam_ciro = df_sonuc['Net Ciro'].sum()
            toplam_kargo = df_sonuc['Kargo'].sum()
            toplam_komisyon = df_sonuc['Komisyon'].sum()
            genel_net_kar = df_sonuc['Net Kâr'].sum() - reklam_gideri
            toplam_adet = df_sonuc[df_sonuc['Durum'] == 'Satış']['Satış Adedi'].sum()
            
            # Özet Kartları
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("💰 Gerçekleşen Net Ciro (İadeler Düştü)", f"₺{toplam_ciro:,.2f}")
            m2.metric("🟢 Gerçek Net Kâr (Maliyet & Reklam Optimize)", f"₺{genel_net_kar:,.2f}")
            m3.metric("📦 Toplam Teslim Edilen Ürün", f"{int(toplam_adet)} Adet")
            m4.metric("🚚 Toplam Kargo Maliyeti", f"₺{abs(toplam_kargo):,.2f}")
            
            # 🏢 Marka Analizi
            st.write("### 🏢 Marka Bazlı Kârlılık Analizi")
            df_marka = df_sonuc.groupby('Marka').agg({'Net Ciro':'sum', 'Net Kâr':'sum', 'Satış Adedi':'sum'}).reset_index()
            st.dataframe(df_marka.style.format({'Net Ciro': '₺{:.2f}', 'Net Kâr': '₺{:.2f}'}), use_container_width=True)
            
            fig = px.bar(df_marka, x='Marka', y='Net Kâr', title="Markaların Net Kâr Dağılımı", color='Marka')
            st.plotly_chart(fig, use_container_width=True)
            
            # 🔍 EKSİK MALİYET DEDEKTİFİ
            if eksik_maliyetler:
                st.write("---")
                st.error("⚠️ Maliyet Listesinde Bulunmayan ve Kârı Sapıttıran Barkodlar (Acilen listenize ekleyin):")
                df_eksik = pd.DataFrame(list(eksik_maliyetler), columns=["Barkod", "Ürün Adı"])
                st.dataframe(df_eksik, use_container_width=True)
            
            # Excel İndirme Butonu
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_sonuc.to_excel(writer, index=False, sheet_name='Kâr Analiz Sonucu v3')
            processed_data = output.getvalue()
            
            st.download_button(
                label="📥 Gelişmiş Sonuçları Excel Olarak İndir",
                data=processed_data,
                file_name="Trendyol_Kar_Analizi_v3.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
    else:
        st.error("Lütfen analiz için 3 dosyayı da eksiksiz yükleyin!")