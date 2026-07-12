import streamlit as st
import pandas as pd
import plotly.express as px
import io

# Sayfa Genişlik ve Başlık Ayarları
st.set_page_config(page_title="Trendyol Ultra Kâr Analiz Paneli v4", layout="wide")

st.title("🚀 Trendyol Ultra Kâr Analiz Paneli v4 (PRO)")
st.markdown("Zarar eden ürün alarmları, iade kargo cezaları ve detaylı gider pastası entegre edilmiştir.")
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
if st.button("🚀 Detaylı Analizi Başlat", use_container_width=True):
    if finans_file and prod_file and maliyet_file:
        with st.spinner("Şirketinizin röntgeni çekiliyor, zararlar hesaplanıyor..."):
            
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
            toplam_iade_kargo_zarari = 0
            
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
                
                # İptal Durumu Kontrolü
                if "iptal" in statü or "reddedildi" in statü:
                    continue
                
                # Maliyet Kontrolü
                if barkod in maliyet_dict:
                    birim_maliyet = maliyet_dict[barkod]
                else:
                    birim_maliyet = 0
                    eksik_maliyetler.add((barkod, urun_adi))
                
                toplam_maliyet = birim_maliyet * adet
                is_iade = "iade" in statü
                
                if is_iade:
                    hesaplanan_ciro = 0
                    hesaplanan_maliyet = 0
                else:
                    hesaplanan_ciro = satis_tutari
                    hesaplanan_maliyet = toplam_maliyet
                
                # Kesintileri Getir
                bolunmus_komisyon = 0
                bolunmus_kargo = 0
                bolunmus_hizmet = 0
                
                if siparis_no in finans_dict:
                    f_data = finans_dict[siparis_no]
                    if f_data['toplam_adet'] > 0:
                        bolunmus_komisyon = (f_data['komisyon'] / f_data['toplam_adet']) * adet
                        bolunmus_kargo = (f_data['kargo'] / f_data['toplam_adet']) * adet
                        bolunmus_hizmet = (f_data['hizmet'] / f_data['toplam_adet']) * adet
                        
                        # Eğer iade ise yediğimiz kargo bedeli bizim için saf zarardır
                        if is_iade:
                            toplam_iade_kargo_zarari += abs(bolunmus_kargo)
                
                # Net Kâr Formülü
                net_kar = hesaplanan_ciro + bolunmus_komisyon + bolunmus_kargo + bolunmus_hizmet - hesaplanan_maliyet
                
                sonuc_listesi.append({
                    "Sipariş No": siparis_no,
                    "Barkod": barkod,
                    "Marka": marka,
                    "Ürün Adı": urun_adi,
                    "Durum": "İade" if is_iade else "Satış",
                    "Satış Adedi": adet,
                    "Net Ciro": hesaplanan_ciro,
                    "Komisyon": bolunmus_komisyon,
                    "Kargo": bolunmus_kargo,
                    "Hizmet Bedeli": bolunmus_hizmet,
                    "Ürün Maliyeti": hesaplanan_maliyet,
                    "Net Kâr": net_kar
                })
                
            df_sonuc = pd.DataFrame(sonuc_listesi)
            
            # 📊 METRİKLER
            toplam_ciro = df_sonuc['Net Ciro'].sum()
            toplam_kargo = df_sonuc['Kargo'].sum()
            toplam_komisyon = df_sonuc['Komisyon'].sum()
            toplam_hizmet = df_sonuc['Hizmet Bedeli'].sum()
            toplam_maliyet_gideri = df_sonuc['Ürün Maliyeti'].sum()
            genel_net_kar = df_sonuc['Net Kâr'].sum() - reklam_gideri
            
            total_rows = len(df_sonuc)
            iade_rows = len(df_sonuc[df_sonuc['Durum'] == 'İade'])
            iade_orani = (iade_rows / total_rows) * 100 if total_rows > 0 else 0
            
            # Üst Özet Kartları
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("💰 Net Ciro (İadeler Hariç)", f"₺{toplam_ciro:,.2f}")
            m2.metric("🟢 Gerçek Net Kâr", f"₺{genel_net_kar:,.2f}")
            m3.metric("📦 Genel İade Oranı", f"%{iade_orani:.2f}")
            m4.metric("🚨 İade Kargo Zararı", f"₺{toplam_iade_kargo_zarari:,.2f}")
            
            # 🛑 ÖZELLİK 1: ZARAR EDEN ÜRÜNLER ALARMI
            df_urun_kar = df_sonuc.groupby(['Barkod', 'Marka', 'Ürün Adı']).agg({'Net Kâr':'sum', 'Satış Adedi':'sum'}).reset_index()
            df_zarar_edenler = df_urun_kar[df_urun_kar['Net Kâr'] < 0].sort_values(by='Net Kâr')
            
            st.write("---")
            if not df_zarar_edenler.empty:
                st.error(f"🛑 DİKKAT: Sattıkça Zarar Ettiren {len(df_zarar_edenler)} Adet Ürün Tespit Edildi! (Acil Zam Yapılmalı)")
                st.dataframe(df_zarar_edenler.style.format({'Net Kâr': '₺{:.2f}'}), use_container_width=True)
            else:
                st.success("✅ Tebrikler! Sattıkça zarar ettiren hiçbir ürününüz bulunmuyor.")
            
            # 💸 ÖZELLİK 3: GİDER DAĞILIM PASTA GRAFİĞİ
            st.write("---")
            st.write("### 🍕 Toplam Cironun Gider Dağılım Röntgeni")
            
            gider_data = {
                "Gider Kalemi": ["Net Kâr", "Ürün Maliyetleri", "Trendyol Komisyonu", "Kargo Giderleri", "Platform Hizmet & Reklam"],
                "Tutar": [
                    max(0, genel_net_kar),
                    toplam_maliyet_gideri,
                    abs(toplam_komisyon),
                    abs(toplam_kargo),
                    abs(toplam_hizmet) + reklam_gideri
                ]
            }
            df_gider_pasta = pd.DataFrame(gider_data)
            fig_pie = px.pie(df_gider_pasta, values='Tutar', names='Gider Kalemi', title="Cironuzun Dağılım Tablosu (%)", hole=0.4)
            st.plotly_chart(fig_pie, use_container_width=True)
            
            # 🏢 Marka Analizi
            st.write("---")
            st.write("### 🏢 Marka Bazlı Detaylı Performans")
            df_marka = df_sonuc.groupby('Marka').agg({'Net Ciro':'sum', 'Net Kâr':'sum', 'Satış Adedi':'sum'}).reset_index()
            st.dataframe(df_marka.style.format({'Net Ciro': '₺{:.2f}', 'Net Kâr': '₺{:.2f}'}), use_container_width=True)
            
            # 🔍 EKSİK MALİYET DEDEKTİFİ
            if eksik_maliyetler:
                st.write("---")
                st.warning("⚠️ Maliyet Listesinde Olmadığı İçin Kârı Yanıltan Barkodlar:")
                df_eksik = pd.DataFrame(list(eksik_maliyetler), columns=["Barkod", "Ürün Adı"])
                st.dataframe(df_eksik, use_container_width=True)
            
            # Excel İndirme
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_sonuc.to_excel(writer, index=False, sheet_name='Ultra Analiz v4')
            processed_data = output.getvalue()
            
            st.write("---")
            st.download_button(
                label="📥 Tüm Profesyonel Raporları Excel Olarak İndir",
                data=processed_data,
                file_name="Trendyol_Ultra_Analiz_v4.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
    else:
        st.error("Lütfen analiz için 3 dosyayı da eksiksiz yükleyin!")