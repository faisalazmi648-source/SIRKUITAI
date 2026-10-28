import os
import google.generativeai as genai
from docx import Document
from docx.shared import Pt, Inches
import streamlit as st

# ==========================================
# KONFIGURASI API KEY & MODEL (GLOBAL)
# ==========================================
API_KEY = st.secrets["GEMINI_API_KEY"]
genai.configure(api_key=API_KEY)

# Gunakan @st.cache_resource agar inisialisasi model hanya terjadi sekali
@st.cache_resource
def init_gemini_model():
    generation_config = {
        "temperature": 0.4,
    }
    return genai.GenerativeModel(
        model_name="gemini-3.6-flash", # Pastikan menggunakan model flash yang stabil
        system_instruction="Anda adalah Pakar Kurikulum dan Guru Penggerak Indonesia yang handal membuat kegiatan pembelajaran.",
        generation_config=generation_config
    )

model = init_gemini_model()

# ==========================================
# FUNGSI PERAKIT PROMPT AUTOMATIS
# ==========================================
def buat_prompt_rpp(data_input):
    prompt_template = f"""
Buatkan dokumen Rencana Kegiatan Pembelajaran lengkap siap cetak (.docx) untuk Kurikulum Merdeka.

**IDENTITAS:**
- Penulis: {data_input.get('nama_penulis')} (NIP: {data_input.get('nip')})
- Instansi: {data_input.get('nama_instansi')}
- Kelas / Semester: {data_input.get('kelas')} / {data_input.get('semester')}
- Mata Pelajaran: {data_input.get('mata_pelajaran')}
- Fokus Materi: {data_input.get('fokus_materi')}
- Tanggal: {data_input.get('tanggal_pembuatan')}

**KOMPONEN UTAMA:**
1. Profil Pelajar Pancasila, CP, ATP, serta Alat & Bahan.
2. Alokasi waktu tiap tahapan.
3. Rincian Aktivitas Guru & Siswa menggunakan **Model SIRKUIT AI (7 Langkah)**:
   - S: Selidiki (Orientasi Masalah Nyata)
   - I: Investigasi (Pengumpulan Data Primer)
   - R: Rancang Pertanyaan untuk AI (Prompting)
   - K: Kritis & Verifikasi (AI Verification Loop)
   - U: Ubah Menjadi Solusi (Perancangan Produk/Solusi)
   - I: Implementasi (Uji Produk & Presentasi)
   - T: Tindak Lanjut (Refleksi & Perbaikan)

Susun secara sistematis, terstruktur rapi, langsung pada isi, dan gunakan format teks tebal (**bold**) pada bagian penting.
"""
    return prompt_template

# ==========================================
# FUNGSI PEMBUAT FILE DOCX
# ==========================================
def buat_file_docx(teks_rpp, nama_file_output):
    doc = Document()
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    baris_list = teks_rpp.split('\n')
    for baris in baris_list:
        baris_clean = baris.strip()
        if not baris_clean:
            doc.add_paragraph()
            continue
        
        p = doc.add_paragraph()
        parts = baris_clean.split('**')
        for i, part in enumerate(parts):
            if not part:
                continue
            run = p.add_run(part)
            run.font.name = 'Calibri'
            run.font.size = Pt(11)
            if i % 2 == 1:
                run.bold = True

    doc.save(nama_file_output)

# ==========================================
# TAMPILAN ANTARMUKA APLIKASI (STREAMLIT)
# ==========================================
st.set_page_config(page_title="Rencana Kegiatan Pembelajaran Berbasis SIRKUIT AI", page_icon="📚", layout="centered")

st.title("📚 RKP SIRKUIT AI)")
st.markdown("Aplikasi berbasis AI untuk menyusun Rencana Kegiatan Pembelajaran secara instan dan siap diunduh dalam format Word (`.docx`).")
st.divider()

# Inisialisasi session state untuk mencegah klik ganda
if "is_generating" not in st.session_state:
    st.session_state.is_generating = False

# Form Input Data di Web
with st.form("form_rpp"):
    st.subheader("✍️ Masukkan Identitas & Materi Pembelajaran")
    
    col1, col2 = st.columns(2)
    with col1:
        nama_penulis = st.text_input("Nama Penulis & Gelar", "Tulis Disini")
        nip = st.text_input("NIP", "Tulis Disini / -")
        nama_instansi = st.text_input("Nama Instansi / Sekolah", "Tulis Disini")
    with col2:
        tanggal_pembuatan = st.text_input("Tanggal Pembuatan", "27 Oktober 2026")
        semester = st.selectbox("Semester", ["Ganjil (1)", "Genap (2)"])
        kelas = st.text_input("Kelas", "V (Lima)")

    mata_pelajaran = st.text_input("Mata Pelajaran", "Tulis Disini")
    fokus_materi = st.text_input("Fokus Materi / Topik", "Tulis Disini")

    # Tombol Eksekusi
    submit_button = st.form_submit_button(label="🚀 Buat & Generate RKP Sekarang")

# Proses ketika tombol diklik
if submit_button:
    data_rpp = {
        "nama_penulis": nama_penulis,
        "nip": nip,
        "nama_instansi": nama_instansi,
        "tanggal_pembuatan": tanggal_pembuatan,
        "semester": semester,
        "kelas": kelas,
        "mata_pelajaran": mata_pelajaran,
        "fokus_materi": fokus_materi
    }

    with st.spinner("⏳ Sedang merakit promptnya, kurang lebih 5 menitanlah jadi tunggu sebentar yah..."):
        try:
            # 1. Buat Prompt
            prompt_final = buat_prompt_rpp(data_rpp)

            # 2. Panggil Gemini API
            respon = model.generate_content(
                prompt_final,
                request_options={'timeout': 300}
            )
            teks_hasil = respon.text

            # 3. Simpan ke File Word (.docx)
            nama_file = f"RPP_{mata_pelajaran.replace(' ', '_')}_Kelas_{kelas}.docx"
            buat_file_docx(teks_hasil, nama_file)

            st.success("🎉 RPP Berhasil Dibuat!")
            
            # Simpan hasil ke session_state agar tidak hilang saat interaksi unduh
            st.session_state["teks_hasil"] = teks_hasil
            st.session_state["nama_file"] = nama_file

        except Exception as e:
            st.error(f"Terjadi kesalahan saat memproses API: {e}")
            st.info("💡 Tips: Batas Free Tier adalah 5 RPM. Tunggu sekitar 60 detik jika Anda baru saja melakukan beberapa kali klik beruntun sebelum mencoba lagi.")

# Menampilkan tombol download dan pratinjau jika file sudah berhasil dibuat di session
if "teks_hasil" in st.session_state and "nama_file" in st.session_state:
    with open(st.session_state["nama_file"], "rb") as file_dokumen:
        st.download_button(
            label="📥 Unduh Dokumen RKP (Format Word)",
            data=file_dokumen,
            file_name=st.session_state["nama_file"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    with st.expander("👁️ Lihat Pratinjau Hasil RKP di Layar"):
        st.markdown(st.session_state["teks_hasil"])
