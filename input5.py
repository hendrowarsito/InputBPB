import os
import streamlit as st
import pdfplumber
import pandas as pd
from openpyxl import load_workbook, Workbook
from io import BytesIO

# Fungsi untuk mengekstrak data dari PDF
def extract_data_from_pdf(pdf_file):
    data = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            if page.extract_tables():
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        data.append(row)
    return data

# Fungsi untuk membersihkan data dan mengonversi format angka
def clean_and_convert_data(data):
    df = pd.DataFrame(data)
    for col in df.columns:
        df[col] = df[col].apply(lambda x: str(x).replace(",", "").strip() if isinstance(x, str) else x)
        try:
            df[col] = pd.to_numeric(df[col], errors="ignore")
        except ValueError:
            pass
    return df

# Fungsi untuk menyimpan data ke Excel (dalam satu sheet)
def save_to_single_sheet(data, excel_path, kota_kabupaten, tahun):
    # Bersihkan dan konversi data
    df = clean_and_convert_data(data)
    
    # Tambahkan kolom Kota/Kabupaten dan Tahun
    df.insert(0, "Kota/Kabupaten", kota_kabupaten)
    df.insert(1, "Tahun", tahun)

    # Konversi kolom Tahun ke tipe numerik
    df["Tahun"] = pd.to_numeric(df["Tahun"], errors="coerce", downcast="integer")

    if not os.path.exists(excel_path):  # Jika file belum ada, buat baru
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="BTB Data")
    else:  # Jika file sudah ada, tambahkan data ke sheet
        existing_data = pd.read_excel(excel_path, sheet_name="BTB Data")
        
        # Pastikan kolom Tahun pada data lama juga numerik
        if "Tahun" in existing_data.columns:
            existing_data["Tahun"] = pd.to_numeric(existing_data["Tahun"], errors="coerce", downcast="integer")

        # Gabungkan data lama dengan data baru
        updated_data = pd.concat([existing_data, df], ignore_index=True)
        
        # Tulis kembali data ke file Excel
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            updated_data.to_excel(writer, index=False, sheet_name="BTB Data")

# Direktori untuk menyimpan data
DATA_DIR = "data_btb"
if not os.path.exists(DATA_DIR):
    os.makedirs(DATA_DIR)

# Aplikasi Streamlit
st.title("Aplikasi Input Data BPB MAPPI")

# Tab untuk input data, melihat data yang sudah diinput, dan download data
tab1, tab2, tab3 = st.tabs(["Input Data", "Data Telah Diinput", "Download Data"])

# Tab 1: Input Data
with tab1:
    st.write("Unggah file PDF Data BPB MAPPI.")

    # Form input data
    kota_kabupaten = st.text_input("Masukkan Kota/Kabupaten:")
    tahun = st.text_input("Masukkan Tahun:")
    uploaded_file = st.file_uploader("Upload PDF file", type=["pdf"])

    if uploaded_file and kota_kabupaten and tahun:
        # Ekstraksi data
        extracted_data = extract_data_from_pdf(uploaded_file)

        # Konversi data
        converted_data = clean_and_convert_data(extracted_data)
        st.write(converted_data)

        # Simpan data ke Excel
        if st.button("Simpan ke Excel"):
            excel_path = os.path.join(DATA_DIR, "btb_data.xlsx")
            save_to_single_sheet(converted_data.values.tolist(), excel_path, kota_kabupaten, tahun)
            st.success(f"Data untuk {kota_kabupaten} tahun {tahun} berhasil disimpan ke {excel_path}")

# Tab 2: Data Telah Diinput
with tab2:
    st.write("Berikut adalah daftar kota/kabupaten dan tahun yang telah diinput:")

    excel_path = os.path.join(DATA_DIR, "btb_data.xlsx")
    if os.path.exists(excel_path):
        all_data = pd.read_excel(excel_path, sheet_name="BTB Data")
        
        # Pastikan nama kolom tidak memiliki spasi tambahan
        all_data.columns = all_data.columns.map(str).str.strip()

        if "Kota/Kabupaten" in all_data.columns and "Tahun" in all_data.columns:
            # Pastikan kolom Tahun adalah numerik
            all_data["Tahun"] = pd.to_numeric(all_data["Tahun"], errors="coerce", downcast="integer")

            # Filter kolom unik untuk kota/kabupaten dan tahun
            summary_data = all_data[["Kota/Kabupaten", "Tahun"]].drop_duplicates()
            st.write(summary_data)
        else:
            st.error("Kolom 'Kota/Kabupaten' dan 'Tahun' tidak ditemukan dalam data!")
    else:
        st.info("Belum ada data yang tersedia.")
        
# Tab 3: Download Data
with tab3:
    st.write("Pilih Data BPB MAPPI yang akan diunduh")

    excel_path = os.path.join(DATA_DIR, "btb_data.xlsx")
    if os.path.exists(excel_path):
        all_data = pd.read_excel(excel_path, sheet_name="BTB Data")
        
        # Filter berdasarkan kota/kabupaten dan tahun
        kota_options = all_data["Kota/Kabupaten"].unique()
        tahun_options = all_data["Tahun"].unique()

        selected_kota = st.selectbox("Pilih Kota/Kabupaten:", kota_options)
        selected_tahun = st.selectbox("Pilih Tahun:", tahun_options)

        filtered_data = all_data[
            (all_data["Kota/Kabupaten"] == selected_kota) &
            (all_data["Tahun"] == selected_tahun)
        ]

        st.write(filtered_data)

        # Fungsi untuk menghasilkan file Excel dalam mode transpose
        def generate_transposed_excel(data):
            transposed_data = data.transpose()
            transposed_data.columns = transposed_data.iloc[0]
            transposed_data = transposed_data[1:]
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                transposed_data.to_excel(writer, index=False, sheet_name="Transposed Data")
            output.seek(0)
            return output

        if st.button("Download Data"):
            transposed_file = generate_transposed_excel(filtered_data)
            st.download_button(
                label="Download File Excel (Transposed)",
                data=transposed_file,
                file_name=f"BTB_Data_{selected_kota}_{selected_tahun}_Transposed.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.info("Belum ada data yang tersedia untuk diunduh.")
