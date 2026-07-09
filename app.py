import streamlit as st
from streamlit_gsheets import GSheetsConnection

# Inisialisasi koneksi dengan mode oauth
conn = st.connection("gsheets", type=GSheetsConnection)

# Membaca data
df = conn.read(spreadsheet="https://docs.google.com/spreadsheets/d/1XwYfBSF4fULdXJFo3TJsJ78JBfANmNtcw8kajWV4nK0/edit?usp=sharing")

st.dataframe(df)
