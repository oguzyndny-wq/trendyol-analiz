import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="E-Ticaret Paneli v12.1", layout="wide")
st.title("🤖 Çok Kanallı E-Ticaret Konsolide Finans Paneli v12.1")
st.markdown("Eksik dosya alanları eklenmiş, kırılma korumalı stabil versiyon.")
st.write("---")

# Session State Hazırlığı (KeyError Koruması)
v_list = [
    'ty_ciro', 'ty_kesinti', 'ty_maliyet', 'ty_kar', 'ty_sip_adet', 'ty_urun_adet',
    'amz_ciro', 'amz_kesinti', 'amz_maliyet', 'amz_kar', 'amz_sip_adet', 'amz_urun_adet',
    'genel_ciro', 'genel_maliyet', 'genel_kar', 'genel_sip_adet', 'genel_urun_adet', 'genel_marj'
]
for k in v_list:
    if k not in st.session_state: st.session_state[k] = 0.0
if 'hesaplandi' not in st.session_state: st.session_state['hesaplandi'] = False

tab_yükleme, tab_rapor = st.tabs(["📥 Veri Yükleme İstasyonu", "📊 Konsolide Finans & Raporlar"])

with tab_yükleme:
    st.subheader("1. Trendyol Raporları")
    col1, col2, col3 = st.columns(3)
    with col1: finans_file = st.file_uploader("SiparisNo dosyası", type=["xlsx", "xls"], key="finans")