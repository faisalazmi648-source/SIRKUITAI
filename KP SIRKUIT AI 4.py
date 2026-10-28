import os
import google.generativeai as genai
from docx import Document
from docx.shared import Pt, Inches, RGBColor
import streamlit as st

# ==========================================
# KONFIGURASI API KEY & MODEL (GLOBAL)
# ==========================================
API_KEY = st.secrets["GEMINI_API_KEY"]
genai.configure(api_key=API_KEY)

# Gunakan @st.cache_resource agar inisialisasi model hanya terjadi sekali, 
# mencegah pemborosan request/token saat aplikasi berjalan.
@st.cache_resource
def init_gemini_model():
    generation_config = {
        "temperature": 0.4,
    }
    # Pastikan menggunakan nama model yang valid dan stabil di Google AI Studio
    return genai.GenerativeModel(
        model_name="gemini-3-flash", # Ubah ke gemini-1.5-flash atau model flash terbaru yang stabil
        system_instruction="Anda adalah Pakar Kurikulum dan Guru Penggerak Indonesia yang handal membuat modul ajar.",
        generation_config=generation_config
    )

model = init_gemini_model()

# ==========================================
# FUNGSI PERAKIT PROMPT AUTOMATIS
# ==========================================
def buat_prompt_rpp(data_input):
    prompt_template = f"""
Anda adalah seorang guru profesional yang handal dalam membuat rencana kegiatan pembelajaran. buatkan saya kegiatan pembelajaran dalam bentuk documen atau format word. Kegiatan pembelajaran untuk Untuk kelas {data_input.get('kelas')} dengan mata pelajaran {data_input.get('mata_pelajaran')} dengan fokus materi {data_input.get('fokus_materi')}. identitas kegiatan pembelajaran meliputi: nama penulis :{data_input.get('nama_penulis')}; NIP. {data_input.get('nip')}; Nama instansi: {data_input.get('nama_instansi')}; tanggal pembuatan: {data_input.get('tanggal_pembuatan')}; dan semester: {data_input.get('semester')}. kegiatan pembelajaran meliputi  aktifitas guru dan siswanya, alokasi waktu tiap kegiatan, alat dan bahan yang perlu disiapkan, CP, ATP, Profil pelajar pancasila. Kegiatan pembelajaran disesuaikan dengan model pembelajaran di bawah ini: 

Pengembangan model SIRKUIT AI dilandaskan pada kerangka Inkuiri Kolaboratif (Collaborative Inquiry) melalui siklus Assess-Design-Implement-Measure/Reflect/Change. Model ini terdiri dari 7 tahapan yang membentuk akronim S-I-R-K-U-I-T. Berikut adalah penjelasan detail beserta contoh penerapannya yang bisa diadaptasi untuk semua mata pelajaran.

1. S - Selidiki (Orientasi Masalah Nyata), Aktifitas guru: Guru menyajikan masalah yang benar-benar ada di lingkungan sekitar sesuai mata pelajaran dan Fokus Materi / Topik yang telah ditentukan. Aktifitas Siswa yaitu mengamati, membangun rasa ingin tahu, dan merumuskan pertanyaan investigasi berdasarkan mata pelajaran dan Fokus Materi / Topik yang telah ditentukan.
2. I - Investigasi (Pengumpulan Data), Aktifitas guru: meminta siswa membentuk kelompok untuk melakukan investigasi tentang masalah sesuai mata pelajaran dan Fokus Materi / Topik yang telah ditentukan. Atifitas siswa yaitu: secara berkelompok, siswa mengumpulkan data primer (wawancara, observasi fisik, data numerik lapangan, buku cetak).
3. R - Rancang Pertanyaan untuk AI (Prompting), Aktifitas guru yaitu memberikan contoh prompt untuk menyelesaikan masalah sesuai mata pelajaran dan Fokus Materi / Topik yang telah ditentukan untuk di copy paste di platform AI. Sedangkan aktifitas siswa yaitu: siswa belajar membuat prompt spesifik untuk AI untuk mencari dugaan awal/teori umum sesuai mata pelajaran dan Fokus Materi / Topik yang telah ditentukan.
4. K - Kritis dan Verifikasi (AI Verification Loop), Aktifitas guru: meminta agar siswa membandingkan data yang diperoleh lewat AI dan data primer tetantang masalah sesuai mata pelajaran dan Fokus Materi / Topik yang telah ditentukan. aktifitas siswa yaitu: dalam kelompok Siswa membuat Tabel Verifikasi untuk membandingkan klaim AI dengan data primer / hasil Investigasi nyata di lapangan terkait masalah sesuai mata pelajaran dan Fokus Materi / Topik yang telah ditentukan.
5. U - Ubah Menjadi Solusi, untuk aktifitas guru: meminta siswa untuk membuat produk yang disesuaikan dengan solusi masalah terkait mata pelajaran dan Fokus Materi / Topik yang telah ditentukan, aktifitas siswa membuat produk Berdasarkan data nyata / primer yang diverifikasi dengan teori AI, siswa merancang solusi logis (karya fisik, algoritma/computational thinking, atau maket yang disesuaikan dengan mata pelajaran dan Fokus Materi / Topik yang telah ditentukan).
6. I - Implementasi (Uji Produk), aktifitas guru yaitu: meminta kelompok untuk melakukan presentasi, menunjuk kelompok lain untuk memberi pertanyaan sebagai "tim penguji", dan meberikan penguatan dan apresiasi kepada kelompok yang telah presentasi, aktiftas siswa yaitu: Kelompok mempresentasikan karya dan diuji oleh kelompok lain sebagai "Tim Penguji" validitas data.
7. T - Tindak Lanjut (Refleksi dan Perbaikan), aktifitas guru dan siswa yaitu: melakukan refleksi proses belajar dan menyimpulkan bahwa kebenaran mutlak diukur dari realitas lapangan atau data primer.

Buatkan kegiatan pembelajaran secara sederhana, sistematis, rapi, dan langsung mencakup seluruh 7 tahapan SIRKUIT AI di atas secara rinci untuk mata pelajaran dan fokus materi / topik yang telah ditentukan.
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
st.set_page_config(page_title="Kegiatan Pembelajaran Berbasis SIRKUIT AI", page_icon="📚", layout="centered")

st.title("📚 Kegiatan Pembelajaran (Model SIRKUIT AI)")
st.markdown("Aplikasi berbasis AI untuk menyusun Rencana Pelaksanaan Pembelajaran secara instan dan siap diunduh dalam format Word (`.docx`).")
st.divider()

# Form Input Data di Web
with st.form("form_rpp"):
    st.subheader("✍️ Masukkan Identitas & Materi Pembelajaran")
    
    col1, col2 = st.columns(2)
    with col1:
        nama_penulis = st.text_input("Nama Penulis & Gelar", "Tulis Namamu Disini")
        nip = st.text_input("NIP", "tulis nip disini, jika tidak ada beri tanda strip (-)")
        nama_instansi = st.text_input("Nama Instansi / Sekolah", "Tulis Nama SDmu Disini")
    with col2:
        tanggal_pembuatan = st.text_input("Tanggal Pembuatan", "27 Oktober 2026")
        semester = st.selectbox("Semester", ["Ganjil (1)", "Genap (2)"])
        kelas = st.text_input("Kelas", "V (Lima)")

    mata_pelajaran = st.text_input("Mata Pelajaran", "Tulis Mata Pelajaranmu Disini")
    fokus_materi = st.text_input("Fokus Materi / Topik", "Tulis Fokus Materi Disini")

    # Tombol Eksekusi
    submit_button = st.form_submit_button(label="🚀 Buat & Generate Sekarang")

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

    with st.spinner("⏳ Sedang merakit prompt otomatis dan memproses ke mesin AI... Mohon tunggu sebentar."):
        # 1. Buat Prompt
        prompt_final = buat_prompt_rpp(data_rpp)

        # 2. Panggil Gemini API menggunakan model yang sudah di-cache di atas
        respon = model.generate_content(
            prompt_final,
            request_options={'timeout': 180}
        )
        teks_hasil = respon.text

        # 3. Simpan ke File Word (.docx)
        nama_file = f"RPP_{mata_pelajaran.replace(' ', '_')}_Kelas_{kelas}.docx"
        buat_file_docx(teks_hasil, nama_file)

    st.success("🎉 RPP Berhasil Dibuat!")
    
    # Tombol Unduh File (.docx) di Browser
    with open(nama_file, "rb") as file_dokumen:
        st.download_button(
            label="📥 Unduh Dokumen RPP (Format Word)",
            data=file_dokumen,
            file_name=nama_file,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    # Pratinjau Teks di Layar
    with st.expander("👁️ Lihat Pratinjau Hasil RPP di Layar"):
        st.markdown(teks_hasil)
