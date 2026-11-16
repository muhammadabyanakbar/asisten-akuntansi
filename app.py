import streamlit as st

from ai_service import buat_chat, kirim_stream
from document_reader import baca_dokumen

BATAS_PERINGATAN = 50 * 1024 * 1024  # 50 MB

ANALISIS_CEPAT = {
    "Ringkasan Eksekutif": (
        "Buat ringkasan eksekutif dari materi ini: kondisi utama, angka penting, "
        "tren, dan risiko, dalam poin-poin singkat."
    ),
    "Analisis Rasio Keuangan": (
        "Hitung dan jelaskan rasio keuangan utama dari materi ini (likuiditas, "
        "solvabilitas, profitabilitas, dan efisiensi) jika datanya tersedia. "
        "Tampilkan rumus, langkah, dan hasilnya. Sebutkan jika data tidak tersedia."
    ),
    "Koreksi Jurnal": (
        "Periksa jurnal atau perlakuan akuntansi dalam materi ini. Tunjukkan "
        "kesalahan jika ada, lalu berikan jurnal koreksi beserta penjelasannya."
    ),
}

st.title("Asisten Akuntansi")

files = st.file_uploader(
    "Unggah materi (PDF, PPTX, DOCX, XLSX, CSV), opsional",
    type=["pdf", "pptx", "docx", "xlsx", "csv"],
    accept_multiple_files=True,
)


@st.cache_data(show_spinner=False)
def ekstrak_materi(nama: str, data: bytes) -> str:
    return baca_dokumen(nama, data)


materi = None
kunci = tuple((f.name, f.size) for f in files) if files else ()

if files:
    for file in files:
        if file.size > BATAS_PERINGATAN:
            st.warning(
                f"'{file.name}' berukuran lebih dari 50 MB. "
                "Pemrosesan bisa memakan waktu lebih lama."
            )
    bagian = []
    for file in files:
        bagian.append(f"=== Materi: {file.name} ===")
        bagian.append(ekstrak_materi(file.name, file.getvalue()))
    materi = "\n".join(bagian)

# Mulai percakapan baru saat pertama kali dibuka, atau saat file berubah.
if st.session_state.get("kunci_materi") != kunci:
    st.session_state.kunci_materi = kunci
    st.session_state.chat = buat_chat()
    st.session_state.riwayat = []
    st.session_state.materi_terkirim = False

if files:
    st.caption(f"{len(files)} materi aktif. Mengubah file akan memulai percakapan baru.")
else:
    st.caption("Tanpa materi. Anda bisa langsung bertanya tentang akuntansi, keuangan, dan bisnis.")

for pesan in st.session_state.riwayat:
    with st.chat_message(pesan["role"]):
        st.markdown(pesan["isi"])

# Tombol analisis cepat, hanya muncul jika ada materi.
aksi = None
if files:
    kolom = st.columns(len(ANALISIS_CEPAT))
    for k, (label, perintah) in zip(kolom, ANALISIS_CEPAT.items()):
        if k.button(label):
            aksi = (label, perintah)

masukan = st.chat_input("Tanyakan sesuatu...")
if masukan and aksi is None:
    aksi = (masukan, masukan)

if aksi:
    tampilan, perintah = aksi
    st.session_state.riwayat.append({"role": "user", "isi": tampilan})
    with st.chat_message("user"):
        st.markdown(tampilan)

    materi_awal = None if st.session_state.materi_terkirim else materi
    with st.chat_message("assistant"):
        jawaban = st.write_stream(
            kirim_stream(st.session_state.chat, materi_awal, perintah)
        )

    if materi_awal is not None:
        st.session_state.materi_terkirim = True
    st.session_state.riwayat.append({"role": "assistant", "isi": jawaban})
