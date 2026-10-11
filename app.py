import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
from datetime import datetime

# --- Konfigurasi Tampilan Web (Form Mode) ---
st.set_page_config(page_title="Form Psikotes CBI", page_icon="📝", layout="centered")

st.title("📝 Form Psikotes Counterproductive Behavior Index (CBI)")
st.write("Silakan isi data diri Anda dan jawablah 140 pernyataan di bawah ini. Waktu Anda adalah 15 menit. Apapun yang Anda pilih, hendaknya sungguh-sungguh menggambarkan diri Anda.")
st.markdown("---")

# --- 1. Form Biodata Peserta ---
st.subheader("A. Data Diri")
col1, col2 = st.columns(2)

with col1:
    nama = st.text_input("Nama Lengkap:")
    jenis_kelamin = st.selectbox("Jenis Kelamin:", ["Laki-laki", "Perempuan"])
    tanggal_lahir = st.text_input("Tanggal Lahir (Misal: 16-08-2001):")

with col2:
    usia = st.text_input("Usia (Misal: 25):")
    pendidikan = st.text_input("Pendidikan (Misal: S1 Akuntansi):")
    tanggal_isi = st.date_input("Tanggal Pengisian:", datetime.today())

st.markdown("---")

# --- 2. Kuesioner 140 Soal ---
st.subheader("B. Pernyataan (1 - 140)")
st.info("Pilih **B (Benar)** jika pernyataan sesuai dengan diri Anda. Pilih **S (Salah)** jika tidak sesuai.")

# Dictionary untuk menyimpan jawaban B/S dari peserta
jawaban_peserta = {}

# Looping untuk membuat 140 soal di halaman web
for i in range(1, 141):
    jawaban_peserta[i] = st.radio(
        f"Soal No {i}:", 
        options=['B', 'S'], 
        horizontal=True, # Membuat pilihan B dan S berjejer ke samping
        index=None,      # Tidak ada jawaban default, peserta wajib klik
        key=f"soal_{i}"  # ID unik untuk tiap soal
    )

st.markdown("---")

# --- 3. Tombol Submit & Logika Ekspor ke Excel ---
if st.button("Kirim Jawaban (Submit) 🚀", type="primary", use_container_width=True):
    
    # Validasi: Cek apakah nama sudah diisi
    if not nama:
        st.error("Mohon isi Nama Lengkap Anda terlebih dahulu di bagian atas.")
    
    # Validasi: Cek apakah ada soal yang belum dijawab
    elif None in jawaban_peserta.values():
        soal_kosong = [k for k, v in jawaban_peserta.items() if v is None]
        st.error(f"Mohon lengkapi semua soal. Anda belum menjawab soal nomor: {', '.join(map(str, soal_kosong))}")
    
    else:
        with st.spinner("Memproses hasil tes..."):
            try:
                # Buka Template Excel
                # Pastikan file "Skoring CBI.xlsx" ada di folder yang sama
                wb = openpyxl.load_workbook("Skoring CBI.xlsx")
                sheet_soal = wb["Soal"]
                
                # --- A. Masukkan Biodata ke Kolom D ---
                sheet_soal["D3"] = tanggal_isi.strftime("%d-%m-%Y")
                sheet_soal["D4"] = nama
                sheet_soal["D5"] = jenis_kelamin
                sheet_soal["D6"] = tanggal_lahir
                sheet_soal["D7"] = usia
                sheet_soal["D8"] = pendidikan
                
                # --- B. Masukkan Tanda Silang 'x' ke Kolom M (Benar) / N (Salah) ---
                # Rumus baris: Soal 1 = 20, Soal 140 = 159
                for soal_no, jawaban in jawaban_peserta.items():
                    baris_excel = 19 + soal_no 
                    
                    if jawaban == 'B':
                        sheet_soal[f"M{baris_excel}"] = "x"
                        sheet_soal[f"N{baris_excel}"] = ""
                    elif jawaban == 'S':
                        sheet_soal[f"M{baris_excel}"] = ""
                        sheet_soal[f"N{baris_excel}"] = "x"
                
                # --- C. Siapkan File Excel untuk HR (mas Ade) ---
                output = BytesIO()
                wb.save(output)
                output.seek(0)
                
                st.success(f"Terima kasih {nama}! Jawaban Anda telah tersimpan.")
                
                # Tombol Download untuk HR (Peserta melihat form, HR mendownload)
                nama_file_baru = f"Hasil_CBI_{nama}.xlsx"
                st.download_button(
                    label="📥 Download Hasil Skoring (Khusus HR)",
                    data=output,
                    file_name=nama_file_baru,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except FileNotFoundError:
                st.error("File 'Skoring CBI.xlsx' tidak ditemukan! Pastikan file tersebut ada di folder yang sama.")
            except Exception as e:
                st.error(f"Terjadi kesalahan teknis: {e}")