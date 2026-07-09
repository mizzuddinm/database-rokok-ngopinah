import streamlit as st
from streamlit_gsheets import GSheetsConnection

# Inisialisasi koneksi dengan mode oauth
conn = st.connection("gsheets", type=GSheetsConnection)

# Membaca data
df = conn.read(spreadsheet="MASUKKAN_URL_GOOGLE_SHEET_ANDA_DI_SINI")

st.dataframe(df)
