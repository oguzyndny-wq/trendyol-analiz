import streamlit as st
import pandas as pd
import plotly.express as px
import io
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.graphics.barcode import code128

# Sayfa Genişlik ve Başlık Ayarları
st.set_page_config(page_title="Trendyol Akıllı Yapay Zeka Paneli v5", layout="wide")

st.title("🤖 Trendyol Akıllı Yapay Zeka Paneli v5 (AI PRO)")
st.markdown("Fiyat tavsiyeleri, reklam verimlilik motoru ve tek tıkla barkod etiket üretici entegre edilmiştir.")
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

# PDF Barkod Oluşturma Fonksiyonu
def generate_barcode_pdf(df):
    buffer = io.BytesIO()
    p = canvas.Canvas(buffer, pagesize=(283, 283)) # Yaklaşık 100x100mm termal etiket boyutu
    
    # Sadece satışı olan benzersiz barkodları basalım
    unique_products = df[df['Durum'] == 'Satış'][['Barkod', 'Ürün Adı', 'Marka']].drop_duplicates()
    
    for idx, row in unique_products.iterrows():
        barkod_str = str(row['Barkod']).strip()
        urun_adi = str(row['Ürün Adı'])[:40] # Sığması için kısaltalım
        marka = str(row['Marka'])
        
        # Etiket Tasarımı
        p.setFont("Helvetica-Bold", 12)
        p.drawString(20, 250, marka)
        
        p.setFont("Helvetica", 10)
        p.drawString(20, 230, urun_adi)
        
        # Barkod Çizimi (Code 128)
        try:
            bc = code128.Code128(barkod_str, barHeight=50, barWidth=1.2)
            bc.drawOn(p, 20, 100)
        except:
            p.drawString(20, 120, "[Barkod Çizilemedi]")
            
        p.setFont("Helvetica", 11)
        p.drawString(20, 70, f"Barkod: {barkod_str}")
        
        p.showPage()
        
    p.save()
    return buffer.getvalue()

# ANALİZİ BAŞLAT BUTONU
if st.button("🚀 Akıllı Yapay Zeka Analizini Başlat", use_container_width=True):
    if finans_file and prod_file and maliyet_file:
        with st.spinner("Yapay zeka motoru fiyatları optimize ediyor, etiketleri hazırlıyor..."):
            
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
                
                if "iptal" in statü or "reddedildi" in statü:
                    continue
                
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
                
                bolunmus_komisyon = 0
                bolunmus_kargo = 0
                bolunmus_hizmet = 0
                
                if siparis_no in finans_dict:
                    f_data = finans_dict[siparis_no]
                    if f_data['toplam_adet'] > 0:
                        bolunmus_komisyon = (f_data['komisyon'] / f_data['toplam_adet']) * adet
                        bolunmus_kargo = (f_data['kargo'] / f_data['toplam_adet']) * adet
                        bolunmus_hizmet = (f_data['hizmet'] / f_data['toplam_adet']) * adet
                        
                        if is_iade:
                            toplam_iade_kargo_zarari += abs(bolunmus_kargo)
                
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
            
            # METRİKLER
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
            
            # 🛑 ÖZELLİK 1: AKILLI FİYAT OPTİMİZASYON MOTORU (ZAM TAVSİYELİ)
            st.write("---")
            st.subheader("💡 Yapay Zeka Akıllı Fiyatlandırma ve Zam Tavsiyeleri")
            
            df_urun_analiz = df_sonuc.groupby(['Barkod', 'Marka', 'Ürün Adı']).agg({
                'Net Ciro': 'sum',
                'Net Kâr': 'sum',
                'Satış Adedi': 'sum',
                'Komisyon': 'sum',
                'Kargo': 'sum'
            }).reset_index()
            
            tavsiye_listesi = []
            for idx, row in df_urun_analiz.iterrows():
                if row['Satış Adedi'] > 0:
                    birim_ciro = row['Net Ciro'] / row['Satış Adedi']
                    birim_kar = row['Net Kâr'] / row['Satış Adedi']
                    birim_kesinti_orani = abs(row['Komisyon'] + row['Kargo']) / row['Net Ciro'] if row['Net Ciro'] > 0 else 0.30
                    
                    if birim_kar < 0:
                        gerekli_fiyat_artisi = abs(birim_kar) * (1 + birim_kesinti_orani)
                        tavsiye_listesi.append({
                            "Barkod": row['Barkod'],
                            "Marka": row['Marka'],
                            "Ürün Adı": row['Ürün Adı'],
                            "Mevcut Birim Kâr": f"₺{birim_kar:.2f}",
                            "Durum": "🔴 ZARAR EDİYOR",
                            "AI Tavsiyesi": f"Trendyol Satış Fiyatını En Az ₺{gerekli_fiyat_artisi:.2f} ARTIRMALISINIZ!"
                        })
                    elif birim_kar < (birim_ciro * 0.1): # Kâr marjı %10'un altındaysa uyarı ver
                        tavsiye_listesi.append({
                            "Barkod": row['Barkod'],
                            "Marka": row['Marka'],
                            "Ürün Adı": row['Ürün Adı'],
                            "Mevcut Birim Kâr": f"₺{birim_kar:.2f}",
                            "Durum": "🟡 DÜŞÜK KÂR",
                            "AI Tavsiyesi": "Kâr marjınız kritik seviyede. Rekabet elveriyorsa %5 zam yapılması önerilir."
                        })
            
            if tavsiye_listesi:
                st.dataframe(pd.DataFrame(tavsiye_listesi), use_container_width=True)
            else:
                st.success("✅ Harika! Tüm ürünlerinizin kâr marjı sağlıklı durumda, fiyata dokunmaya gerek yok.")
                
            # 🎯 ÖZELLİK 3: REKLAM VERİMLİLİK MOTORU
            if reklam_gideri > 0:
                st.write("---")
                st.subheader("🎯 Reklam Verimlilik Skoru (ROAS Analizörü)")
                roas = toplam_ciro / reklam_gideri if reklam_gideri > 0 else 0
                
                rc1, rc2 = st.columns(2)
                rc1.metric("🎯 Reklam Verimlilik Skoru (ROAS)", f"{roas:.2f}x", help="Harcanan her 1 TL reklamın size kaç TL ciro getirdiğini gösterir.")
                
                if roas >= 5:
                    rc2.success("🔥 MÜKEMMEL: Reklam performansınız harika. Bu bütçeyle reklam vermeye devam edebilir, hatta bütçeyi artırabilirsiniz!")
                elif roas >= 3:
                    rc2.warning("🟡 ORTA SEVİYE: Reklam başabaş noktasında. Reklam verdiğiniz ürünlerin kâr marjını kontrol edin.")
                else:
                    rc2.error("🚨 KRİTİK ZARAR: Reklamınız harcadığı parayı çıkaramıyor. Reklam verdiğiniz kelimeleri veya ürünleri acilen değiştirin!")

            # 💸 GİDER DAĞILIM PASTA GRAFİĞİ
            st.write("---")
            st.write("### 🍕 Toplam Cironun Gider Dağılım Röntgeni")
            gider_data = {
                "Gider Kalemi": ["Net Kâr", "Ürün Maliyetleri", "Trendyol Komisyonu", "Kargo Giderleri", "Platform Hizmet & Reklam"],
                "Tutar": [max(0, genel_net_kar), toplam_maliyet_gideri, abs(toplam_komisyon), abs(toplam_kargo), abs(toplam_hizmet) + reklam_gideri]
            }
            df_gider_pasta = pd.DataFrame(gider_data)
            fig_pie = px.pie(df_gider_pasta, values='Tutar', names='Gider Kalemi', title="Cironuzun Dağılım Tablosu (%)", hole=0.4)
            st.plotly_chart(fig_pie, use_container_width=True)
            
            # 🖨️ ÖZELLİK 4: TEK TIKLA BARKOD ETİKETİ ÜRETİCİ
            st.write("---")
            st.subheader("🖨️ Operasyon ve Paketleme Kolaylığı")
            pdf_data = generate_barcode_pdf(df_sonuc)
            st.download_button(
                label="🖨️ Satılan Ürünlerin Barkod Etiketlerini İndir (Termal Yazıcı İçin Hazır PDF)",
                data=pdf_data,
                file_name="Satilan_Urun_Barkod_Etiketleri.pdf",
                mime="application/pdf",
                use_container_width=True
            )
            
            # Excel İndirme
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_sonuc.to_excel(writer, index=False, sheet_name='AI Pro Raporu v5')
            processed_data = output.getvalue()
            
            st.download_button(
                label="📥 Tüm Sonuçları Detaylı Excel Olarak İndir",
                data=processed_data,
                file_name="Trendyol_AI_Analizi_v5.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
    else:
        st.error("Lütfen analiz için 3 dosyayı da eksiksiz yükleyin!")