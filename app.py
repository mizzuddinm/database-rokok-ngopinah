import streamlit as st
import pandas as pd
import io
from datetime import datetime
import plotly.graph_objects as go
from streamlit_gsheets import GSheetsConnection

# Struktur Kolom Master Database Toko
COLUMNS = [
    "Batch", "Tanggal", "Merek Rokok", "Satuan Beli", "Harga Beli Total", 
    "Isi per Slop", "Isi per Bungkus", "Total Isi Batang", 
    "HPP per Bungkus", "HPP per Batang", 
    "Harga Jual Bungkus", "Harga Jual Batang",
    "Stok Bungkus (Utuh)", "Stok Batang (Eceran)",
    "Terjual Bungkus", "Terjual Batang"
]

# Inisialisasi Koneksi Cloud ke Google Sheets
conn = st.connection("gsheets", type=GSheetsConnection)

def load_data():
    """Fungsi mengambil data secara real-time dari Google Sheets"""
    try:
        df = conn.read(ttl=0) # ttl=0 memaksa server mengambil data paling segar dari Google tanpa cache kaku
        if df.empty or df.columns[0] == "" or "Batch" not in df.columns:
            return pd.DataFrame(columns=COLUMNS)
        return df
    except:
        return pd.DataFrame(columns=COLUMNS)

def save_data(df):
    """Fungsi mengunci dan menulis ulang data ke Google Sheets Cloud"""
    try:
        conn.update(data=df)
        return True
    except Exception as e:
        st.error(f"Gagal sinkronisasi ke Google Cloud Sheets: {e}")
        return False

# ==========================================
# SUNTIKAN CSS UTAMA - INTERFACE CRM PREMIUM REPOSNSIF
# ==========================================
st.markdown("""
    <style>
    .stApp { background-color: #0B0F19 !important; color: #E2E8F0 !important; }
    [data-testid="stSidebar"] { background-color: #111827 !important; border-right: 1px solid #1F2937 !important; }
    [data-testid="stSidebar"] .stRadio > div { gap: 8px !important; padding: 10px 0px !important; }
    [data-testid="stSidebar"] label { color: #9CA3AF !important; font-size: 14px !important; font-weight: 600 !important; padding: 8px 12px !important; border-radius: 8px !important; transition: all 0.3s ease !important; }
    .kpi-card-purple { background: linear-gradient(135deg, #A855F7 0%, #D946EF 100%); padding: 22px; border-radius: 16px; box-shadow: 0 8px 32px rgba(217, 70, 239, 0.15); margin-bottom: 16px; width: 100%; }
    .kpi-card-cyan { background: linear-gradient(135deg, #06B6D4 0%, #3B82F6 100%); padding: 22px; border-radius: 16px; box-shadow: 0 8px 32px rgba(6, 182, 212, 0.15); margin-bottom: 16px; width: 100%; }
    .kpi-card-dark { background-color: #111827; border: 1px solid #1F2937; padding: 22px; border-radius: 16px; margin-bottom: 16px; box-shadow: 0 4px 20px rgba(0,0,0,0.2); width: 100%; }
    .card-title { font-size: 11px; color: rgba(255, 255, 255, 0.7); font-weight: 600; text-transform: uppercase; letter-spacing: 0.8px; margin-bottom: 6px; }
    .card-value { font-size: 28px; font-weight: 700; color: #FFFFFF; letter-spacing: -0.5px; }
    .stDataFrame { border: 1px solid #1F2937 !important; border-radius: 12px !important; overflow: hidden !important; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# GENERATOR DOKUMEN (DOWNLOAD LAPORAN)
# ==========================================
def buat_excel(df):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Laporan Stok & Jual')
    return buffer.getvalue()

def buat_pdf(df, profit, omset, bgk_jual, btg_jual):
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    elements = []
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('Title', parent=styles['Heading1'], fontSize=16, spaceAfter=8, alignment=1)
    meta_style = ParagraphStyle('Meta', parent=styles['Normal'], fontSize=10, spaceAfter=15, alignment=1)
    normal_style = ParagraphStyle('TableText', parent=styles['Normal'], fontSize=8, alignment=1)
    header_style = ParagraphStyle('TableHeader', parent=styles['Normal'], fontSize=8, fontName="Helvetica-Bold", textColor=colors.whitesmoke, alignment=1)
    
    elements.append(Paragraph("<b>LAPORAN KINERJA & INVENTARIS TOKO ROKOK</b>", title_style))
    elements.append(Paragraph(f"Tanggal Cetak: {datetime.today().strftime('%d-%m-%Y %H:%M')} | Total Profit Laporan: Rp {profit:,.0f} | Total Omset: Rp {omset:,.0f}", meta_style))
    
    pdf_df = df[["Batch", "Tanggal", "Merek Rokok", "Satuan Beli", "Stok Bungkus (Utuh)", "Stok Batang (Eceran)", "Terjual Bungkus", "Terjual Batang", "Harga Jual Bungkus", "Harga Jual Batang"]].copy()
    pdf_df["Harga Jual Bungkus"] = pdf_df["Harga Jual Bungkus"].apply(lambda x: f"Rp {x:,.0f}")
    pdf_df["Harga Jual Batang"] = pdf_df["Harga Jual Batang"].apply(lambda x: f"Rp {x:,.0f}")
    
    tabel_data = [[Paragraph(f"<b>{col}</b>", header_style) for col in pdf_df.columns]]
    for _, row in pdf_df.iterrows():
        tabel_data.append([Paragraph(str(item), normal_style) for item in row])
        
    t = Table(tabel_data, repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1E3A8A')), ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('BOTTOMPADDING', (0,0), (-1,-1), 6), ('TOPPADDING', (0,0), (-1,-1), 6),
        ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey), ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F9FAFB')])
    ]))
    elements.append(t)
    doc.build(elements)
    return buffer.getvalue()

st.set_page_config(layout="wide")

# PANEL UTAMA MENU KIRI (SIDEBAR NAVIGASI)
with st.sidebar:
    st.markdown("<h2 style='color: #FFFFFF; font-size: 20px; font-weight: 700; margin-top: 10px; margin-bottom: 5px;'>📊 CRM Control</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #6B7280; font-size: 12px; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 20px;'>Menu Navigasi Cloud</p>", unsafe_allow_html=True)
    menu = st.radio("Pilih Halaman:", ["🏠 Dashboard Utama", "➕ Input Produk Baru", "🔄 Kelola & Update Data", "📋 Data Spreadsheet", "📥 Unduh Laporan"], label_visibility="collapsed")

# MEMUAT DATA DATABASE DARI GOOGLE SHEETS
df_db = load_data()

# ==========================================
# HALAMAN 1: DASHBOARD UTAMA
# ==========================================
if menu == "🏠 Dashboard Utama":
    st.markdown("### 🏠 Executive Overview Dashboard")
    
    if df_db.empty:
        st.info("Belum ada data bisnis terdeteksi di Google Sheets. Silakan isi data barang masuk di Menu 'Input Produk Baru'.")
    else:
        untung_bgk = df_db["Harga Jual Bungkus"] - df_db["HPP per Bungkus"]
        untung_btg = df_db["Harga Jual Batang"] - df_db["HPP per Batang"]
        total_omset = (df_db["Terjual Bungkus"] * df_db["Harga Jual Bungkus"]).sum() + (df_db["Terjual Batang"] * df_db["Harga Jual Batang"]).sum()
        total_profit = (df_db["Terjual Bungkus"] * untung_bgk).sum() + (df_db["Terjual Batang"] * untung_btg).sum()
        modal_mengendap = (df_db["Stok Bungkus (Utuh)"] * df_db["HPP per Bungkus"]).sum() + (df_db["Stok Batang (Eceran)"] * df_db["HPP per Batang"]).sum()
        total_stok_utuh = df_db["Stok Bungkus (Utuh)"].sum()
        total_stok_eceran = df_db["Stok Batang (Eceran)"].sum()

        db_col1, db_col2, db_col3 = st.columns(3)
        with db_col1:
            st.markdown(f'<div class="kpi-card-purple"><div class="card-title">Total Profit Bersih (Aktual)</div><div class="card-value">Rp {total_profit:,.0f}</div></div>', unsafe_allow_html=True)
        with db_col2:
            st.markdown(f'<div class="kpi-card-cyan"><div class="card-title">Total Pendapatan (Omset Toko)</div><div class="card-value">Rp {total_omset:,.0f}</div></div>', unsafe_allow_html=True)
        with db_col3:
            st.markdown(f'<div class="kpi-card-dark"><div class="card-title">Modal Mengendap di Etalase</div><div class="card-value">Rp {modal_mengendap:,.0f}</div></div>', unsafe_allow_html=True)
            
        st.markdown("<br>", unsafe_allow_html=True)
        db_col4, db_col5 = st.columns([1, 2])
        
        with db_col4:
            st.markdown("#### 📦 Status Total Unit Gudang")
            st.metric(label="Total Stok Unit Utuh (Bungkus/Kaleng)", value=f"{total_stok_utuh} Unit")
            st.metric(label="Total Stok Pecahan (Batangan Ecer)", value=f"{total_stok_eceran} Batang")
            stok_kritis = df_db[df_db["Stok Bungkus (Utuh)"] <= 2]
            if not stok_kritis.empty:
                for _, r in stok_kritis.iterrows():
                    st.error(f"**{r['Merek Rokok']}** sisa {r['Stok Bungkus (Utuh)']} bks!")
            else:
                st.success("Semua stok aman.")
                
        with db_col5:
            st.markdown("#### 📊 Grafik Komparasi Stok (Bungkus vs Eceran)")
            df_db["Identitas"] = df_db["Merek Rokok"] + "<br><sup>(" + df_db["Batch"] + ")</sup>"
            fig = go.Figure()
            fig.add_trace(go.Bar(x=df_db["Identitas"], y=df_db["Stok Bungkus (Utuh)"], name="Bungkus / Kaleng Utuh", marker_color="#00D2FF"))
            fig.add_trace(go.Bar(x=df_db["Identitas"], y=df_db["Stok Batang (Eceran)"], name="Batangan Eceran", marker_color="#FF3399"))
            fig.update_layout(
                barmode="group", template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="#111827",
                margin=dict(l=10, r=10, t=10, b=10), legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                xaxis=dict(showgrid=False), yaxis=dict(showgrid=True, gridcolor="#1F2937"), bargap=0.18, bargroupgap=0.04
            )
            st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})

        st.markdown("---")
        st.markdown("#### 📋 Overview Detail Jumlah Stok Setiap Produk")
        df_overview = df_db[["Merek Rokok", "Batch", "Satuan Beli", "Stok Bungkus (Utuh)", "Stok Batang (Eceran)"]].copy()
        df_overview.columns = ["Merek Rokok", "Nomor Batch", "Kemasan Beli", "Stok Utuh (Bks/Klg)", "Stok Eceran (Batang)"]
        st.dataframe(df_overview, use_container_width=True, hide_index=True)

# ==========================================
# HALAMAN 2: INPUT PRODUK BARU
# ==========================================
elif menu == "➕ Input Produk Baru":
    st.markdown("### ➕ Input Registrasi Produk Baru")
    col1, col2 = st.columns(2)
    with col1:
        batch = st.text_input("Batch Pembelian", value="Batch 1")
    with col2:
        tanggal = st.date_input("Tanggal", value=datetime.today())

    merek = st.text_input("Merek Rokok", value="Gudang Garam Surya Merah 16")
    satuan_beli = st.radio("Satuan Beli", ["Slop", "Kaleng", "Bungkus"], horizontal=True)
    harga_beli = st.number_input("Harga Beli Total (Rp)", min_value=0, value=349500, step=500)
    
    if satuan_beli == "Slop":
        col3, col4 = st.columns(2)
        with col3: isi_slop = st.number_input("Isi per-Slop (Bungkus)", min_value=1, value=10)
        with col4: isi_bungkus = st.number_input("Isi per-Bungkus (Batang)", min_value=1, value=16)
        total_batang = isi_slop * isi_bungkus
        hpp_bungkus = harga_beli / isi_slop
        hpp_batang = harga_beli / total_batang
        unit_label = "Bungkus"
    elif satuan_beli == "Kaleng":
        isi_slop, isi_bungkus = 1, st.number_input("Isi per-Kaleng (Batang)", min_value=1, value=50)
        total_batang = isi_bungkus
        hpp_bungkus, hpp_batang = harga_beli, harga_beli / total_batang
        unit_label = "Kaleng"
    else:
        isi_slop, isi_bungkus = 1, st.number_input("Isi per-Bungkus (Batang)", min_value=1, value=16)
        total_batang = isi_bungkus
        hpp_bungkus, hpp_batang = harga_beli, harga_beli / total_batang
        unit_label = "Bungkus"
    
    st.caption(f"*HPP {unit_label}: Rp {hpp_bungkus:,.2f} | HPP Batang: Rp {hpp_batang:,.2f}*")

    st.subheader("Rencana Pemasaran & Stok Awal")
    col5, col6 = st.columns(2)
    with col5:
        qty_bungkus_awal = st.number_input(f"Qty. {unit_label} untuk Dijual Utuh", min_value=0, value=8 if satuan_beli == "Slop" else 1)
        cara_harga_unit = st.radio(f"Metode Penentuan Harga {unit_label}:", ["Input Harga Manual (Rupiah)", "Berdasarkan Persentase Keuntungan dari Modal"])
        if cara_harga_unit == "Input Harga Manual (Rupiah)":
            harga_jual_bungkus = st.number_input(f"Harga Jual / {unit_label} (Rp)", min_value=0, value=39000 if satuan_beli != "Kaleng" else 120000)
        else:
            persen_untung_unit = st.number_input(f"Target Keuntungan Jual {unit_label} (%)", min_value=0.0, value=11.5, step=0.5)
            harga_jual_bungkus = hpp_bungkus * (1 + persen_untung_unit / 100)
            st.info(f"💡 Hasil Hitung Jual: **Rp {harga_jual_bungkus:,.0f} / {unit_label.lower()}**")
            
    with col6:
        qty_eceran_awal = st.number_input(f"Qty. {unit_label} untuk Dibongkar Ecer", min_value=0, value=2 if satuan_beli == "Slop" else 0)
        cara_harga_batang = st.radio("Metode Penentuan Harga Batangan:", ["Input Harga Manual (Rupiah)", "Berdasarkan Persentase Keuntungan dari Modal"])
        if cara_harga_batang == "Input Harga Manual (Rupiah)":
            harga_jual_batang = st.number_input("Harga Jual / Batang (Rp)", min_value=0, value=3000)
        else:
            persen_untung_batang = st.number_input("Target Keuntungan Jual Batangan (%)", min_value=0.0, value=37.0, step=0.5)
            harga_jual_batang = hpp_batang * (1 + persen_untung_batang / 100)
            st.info(f"💡 Hasil Hitung Jual: **Rp {harga_jual_batang:,.0f} / batang**")
        
    stok_batang_awal = qty_eceran_awal * isi_bungkus

    if st.button("Tambahkan Produk", type="primary"):
        if not df_db.empty and ((df_db['Batch'] == batch) & (df_db['Merek Rokok'] == merek)).any():
            st.error("❌ Gagal Simpan! Kombinasi Merek dan Nomor Batch ini sudah terdaftar.")
        else:
            new_data = {
                "Batch": batch, "Tanggal": tanggal.strftime("%Y-%m-%d"), "Merek Rokok": merek, 
                "Satuan Beli": satuan_beli, "Harga Beli Total": harga_beli, "Isi per Slop": isi_slop, "Isi per Bungkus": isi_bungkus, 
                "Total Isi Batang": total_batang, "HPP per Bungkus": hpp_bungkus, "HPP per Batang": hpp_batang, 
                "Harga Jual Bungkus": harga_jual_bungkus, "Harga Jual Batang": harga_jual_batang,
                "Stok Bungkus (Utuh)": qty_bungkus_awal, "Stok Batang (Eceran)": stok_batang_awal, "Terjual Bungkus": 0, "Terjual Batang": 0
            }
            df_db = pd.concat([df_db, pd.DataFrame([new_data])], ignore_index=True)
            if save_data(df_db):
                st.success(f"🎉 Sukses! Produk '{merek}' ({batch}) berhasil dikunci dan diunggah ke Google Sheets Cloud.")

# ==========================================
# HALAMAN 3: KELOLA & UPDATE DATA
# ==========================================
elif menu == "🔄 Kelola & Update Data":
    st.markdown("### 🔄 Pusat Pengaturan Stok & Harga Jual")
    if df_db.empty:
        st.info("Database Google Sheets kosong. Silakan isi produk terlebih dahulu.")
    else:
        df_db['Pilihan'] = df_db['Merek Rokok'] + " (" + df_db['Batch'] + ")"
        pilihan_produk = st.selectbox("Pilih Produk/Batch yang Akan Dikelola", df_db['Pilihan'].unique())
        idx = df_db[df_db['Pilihan'] == pilihan_produk].index[0]
        row = df_db.loc[idx]
        unit_stok = "Kaleng" if row['Satuan Beli'] == "Kaleng" else "Bungkus"
        
        st.info(f"📋 **Status** -> **Stok {unit_stok}:** {row['Stok Bungkus (Utuh)']} | **Stok Eceran:** {row['Stok Batang (Eceran)']} btg | **Harga Jual:** Rp {row['Harga Jual Bungkus']:,.0f}/{unit_stok.lower()} & Rp {row['Harga Jual Batang']:,.0f}/btg")
        aksi = st.radio("Pilih Jenis Tindakan", ["Catat Penjualan (Stok Berkurang)", "Tambah Stok / Kulakan Baru (Stok Bertambah)", "🔄 Ubah Rencana Harga Jual", "❌ Hapus Produk/Batch dari Database"])
        
        if aksi in ["Catat Penjualan (Stok Berkurang)", "Tambah Stok / Kulakan Baru (Stok Bertambah)"]:
            tipe_stok = st.radio("Satuan yang Diubah", [f"{unit_stok} (Utuh)", "Batangan (Eceran)"])
            jumlah_perubahan = st.number_input("Jumlah (Qty) Barang", min_value=1, value=1)
            if st.button("Eksekusi Perubahan Stok", type="primary"):
                if aksi == "Catat Penjualan (Stok Berkurang)":
                    if tipe_stok == f"{unit_stok} (Utuh)":
                        if row['Stok Bungkus (Utuh)'] >= jumlah_perubahan:
                            df_db.loc[idx, 'Stok Bungkus (Utuh)'] -= jumlah_perubahan
                            df_db.loc[idx, 'Terjual Bungkus'] += jumlah_perubahan
                            st.success("Penjualan utuh dicatat!")
                        else: st.error("Stok tidak mencukupi!")
                    else:
                        if row['Stok Batang (Eceran)'] >= jumlah_perubahan:
                            df_db.loc[idx, 'Stok Batang (Eceran)'] -= jumlah_perubahan
                            df_db.loc[idx, 'Terjual Batang'] += jumlah_perubahan
                            st.success("Penjualan eceran dicatat!")
                        elif row['Stok Bungkus (Utuh)'] > 0:
                            df_db.loc[idx, 'Stok Bungkus (Utuh)'] -= 1
                            df_db.loc[idx, 'Stok Batang (Eceran)'] += (row['Isi per Bungkus'] - jumlah_perubahan)
                            df_db.loc[idx, 'Terjual Batang'] += jumlah_perubahan
                            st.success("1 Unit utuh otomatis dibongkar untuk eceran!")
                        else: st.error("Stok habis!")
                else:
                    if tipe_stok == f"{unit_stok} (Utuh)": df_db.loc[idx, 'Stok Bungkus (Utuh)'] += jumlah_perubahan
                    else: df_db.loc[idx, 'Stok Batang (Eceran)'] += jumlah_perubahan
                
                df_db = df_db.drop(columns=['Pilihan'])
                if save_data(df_db): st.rerun()
                
        elif aksi == "🔄 Ubah Rencana Harga Jual":
            col_h1, col_h2 = st.columns(2)
            with col_h1: harga_jual_bungkus_baru = st.number_input(f"Harga Jual Baru / {unit_stok} (Rp)", min_value=0, value=int(row['Harga Jual Bungkus']), step=500)
            with col_h2: harga_jual_batang_baru = st.number_input("Harga Jual Baru / Batang (Rp)", min_value=0, value=int(row['Harga Jual Batang']), step=100)
            if st.button("Simpan Perubahan Harga", type="primary"):
                df_db.loc[idx, 'Harga Jual Bungkus'] = harga_jual_bungkus_baru
                df_db.loc[idx, 'Harga Jual Batang'] = harga_jual_batang_baru
                df_db = df_db.drop(columns=['Pilihan'])
                if save_data(df_db): st.success("Harga berhasil diperbarui!"); st.rerun()
                
        elif aksi == "❌ Hapus Produk/Batch dari Database":
            konfirmasi = st.checkbox("Ya, saya yakin ingin menghapus data ini secara permanen.")
            if st.button("Hapus Permanen Sekarang", disabled=not konfirmasi):
                df_db = df_db.drop(idx).drop(columns=['Pilihan'])
                if save_data(df_db): st.rerun()

# ==========================================
# HALAMAN 4: SPREADSHEET VIEW
# ==========================================
elif menu == "📋 Data Spreadsheet":
    st.markdown("### 📋 Master Database Spreadsheet (Google Sheets)")
    if df_db.empty: st.info("Belum ada data.")
    else:
        df_display = df_db.copy()
        df_display["Untung /Bungkus (Rp)"] = df_display["Harga Jual Bungkus"] - df_display["HPP per Bungkus"]
        df_display["Untung /Batang (Rp)"] = df_display["Harga Jual Batang"] - df_display["HPP per Batang"]
        df_display["Margin Bungkus (%)"] = (df_display["Untung /Bungkus (Rp)"] / df_display["Harga Jual Bungkus"]) * 100
        df_display["Margin Eceran (%)"] = (df_display["Untung /Batang (Rp)"] / df_display["Harga Jual Batang"]) * 100
        
        col_fv1, col_fv2 = st.columns(2)
        with col_fv1: filter_merek_view = st.selectbox("Saring Merek:", ["-- Tampilkan Semua Merek --"] + list(df_display["Merek Rokok"].unique()))
        with col_fv2: filter_batch_view = st.selectbox("Saring Batch:", ["-- Tampilkan Semua Batch --"] + list(df_display["Batch"].unique()))
        
        if filter_merek_view != "-- Tampilkan Semua Merek --": df_display = df_display[df_display["Merek Rokok"] == filter_merek_view]
        if filter_batch_view != "-- Tampilkan Semua Batch --": df_display = df_display[df_display["Batch"] == filter_batch_view]
        
        st.dataframe(df_display.style.format({
            "Harga Beli Total": "Rp {:,.0f}", "HPP per Bungkus": "Rp {:,.2f}", "HPP per Batang": "Rp {:,.2f}",
            "Harga Jual Bungkus": "Rp {:,.0f}", "Harga Jual Batang": "Rp {:,.0f}",
            "Untung /Bungkus (Rp)": "Rp {:,.0f}", "Untung /Batang (Rp)": "Rp {:,.0f}",
            "Margin Bungkus (%)": "{:.2f}%", "Margin Eceran (%)": "{:.2f}%"
        }), use_container_width=True)

# ==========================================
# HALAMAN 5: EXPORT LAPORAN
# ==========================================
elif menu == "📥 Unduh Laporan":
    st.markdown("### 📥 Pusat Cetak Laporan Keuangan Finansial")
    if df_db.empty: st.info("Database kosong.")
    else:
        df_dl = df_db.copy()
        df_dl["Untung /Bungkus (Rp)"] = df_dl["Harga Jual Bungkus"] - df_dl["HPP per Bungkus"]
        df_dl["Untung /Batang (Rp)"] = df_dl["Harga Jual Batang"] - df_dl["HPP per Batang"]
        
        col_f1, col_f2 = st.columns(2)
        with col_f1: filter_merek = st.selectbox("Laporan Merek:", ["-- Tampilkan Semua Merek --"] + list(df_dl["Merek Rokok"].unique()))
        with col_f2: filter_batch = st.selectbox("Laporan Batch:", ["-- Tampilkan Semua Batch --"] + list(df_dl["Batch"].unique()))
        
        if filter_merek != "-- Tampilkan Semua Merek --": df_dl = df_dl[df_dl["Merek Rokok"] == filter_merek]
        if filter_batch != "-- Tampilkan Semua Batch --": df_dl = df_dl[df_dl["Batch"] == filter_batch]
        
        u_bgk_f = df_dl["Harga Jual Bungkus"] - df_dl["HPP per Bungkus"]
        u_btg_f = df_dl["Harga Jual Batang"] - df_dl["HPP per Batang"]
        p_omset_f = (df_dl["Terjual Bungkus"] * df_dl["Harga Jual Bungkus"]).sum() + (df_dl["Terjual Batang"] * df_dl["Harga Jual Batang"]).sum()
        p_profit_f = (df_dl["Terjual Bungkus"] * u_bgk_f).sum() + (df_dl["Terjual Batang"] * u_btg_f).sum()
        
        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            st.download_button(label="🟢 Ekspor ke Excel (.xlsx)", data=buat_excel(df_dl), file_name="Laporan_Rokok.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        with col_dl2:
            st.download_button(label="🔴 Cetak Laporan PDF Resmi (.pdf)", data=buat_pdf(df_dl, p_profit_f, p_omset_f, df_dl["Terjual Bungkus"].sum(), df_dl["Terjual Batang"].sum()), file_name="Laporan_Resmi.pdf", mime="application/pdf", use_container_width=True)
        st.dataframe(df_dl, use_container_width=True)