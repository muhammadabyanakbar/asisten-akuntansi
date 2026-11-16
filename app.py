import os
from collections.abc import Iterator
from io import BytesIO

import pandas as pd
import streamlit as st
from docx import Document
from google import genai
from google.genai import types
from pptx import Presentation
from pypdf import PdfReader

MODEL = "gemini-3.5-flash-lite"
BATAS_PERINGATAN = 50 * 1024 * 1024  # 50 MB

SYSTEM_INSTRUCTION = """Kamu adalah asisten AI untuk akuntansi, analisis keuangan, dan bisnis. Kamu berpikir seperti akuntan senior dan analis bisnis, dan membantu mahasiswa akuntansi serta profesional.

Aturan menjawab:
- Jika pengguna sudah mengunggah materi, jawab berdasarkan materi tersebut. Jika informasinya tidak ada di materi, katakan dengan jelas bahwa materi tidak memuat informasi itu. Boleh menjelaskan konsep umumnya, dengan menandainya sebagai pengetahuan umum. Jangan mengarang angka dari materi.
- Jika tidak ada materi, jawab sebagai asisten akuntansi umum: jelaskan konsep, bantu perhitungan, dan koreksi jurnal.
- Sebutkan nomor halaman atau slide jika relevan.
- Untuk perhitungan, tampilkan rumus, langkah, dan hasil akhirnya.
- Tulis angka dengan presisi dan satuan yang jelas.
- Gunakan bahasa Indonesia yang formal namun mudah dipahami."""

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


# ---------- Pembacaan dokumen ----------

def baca_pdf(data: bytes) -> str:
    reader = PdfReader(BytesIO(data))
    hasil = []
    for nomor, halaman in enumerate(reader.pages, start=1):
        hasil.append(f"Halaman {nomor}:")
        hasil.append(halaman.extract_text() or "")
    return "\n".join(hasil)


def baca_pptx(data: bytes) -> str:
    prs = Presentation(BytesIO(data))
    hasil = []
    for nomor, slide in enumerate(prs.slides, start=1):
        hasil.append(f"Slide {nomor}:")
        for shape in slide.shapes:
            if shape.has_text_frame:
                hasil.append(shape.text_frame.text)
    return "\n".join(hasil)


def baca_docx(data: bytes) -> str:
    dokumen = Document(BytesIO(data))
    hasil = []
    for paragraf in dokumen.paragraphs:
        if paragraf.text.strip():
            hasil.append(paragraf.text)
    for nomor, tabel in enumerate(dokumen.tables, start=1):
        hasil.append(f"Tabel {nomor}:")
        for baris in tabel.rows:
            hasil.append(" | ".join(sel.text.strip() for sel in baris.cells))
    return "\n".join(hasil)


def baca_excel(data: bytes) -> str:
    semua_sheet = pd.read_excel(BytesIO(data), sheet_name=None)
    hasil = []
    for nama_sheet, tabel in semua_sheet.items():
        hasil.append(f"Sheet: {nama_sheet}")
        hasil.append(tabel.to_csv(index=False))
    return "\n".join(hasil)


def baca_csv(data: bytes) -> str:
    tabel = pd.read_csv(BytesIO(data))
    return tabel.to_csv(index=False)


def baca_dokumen(nama: str, data: bytes) -> str:
    nama = nama.lower()
    if nama.endswith(".pptx"):
        return baca_pptx(data)
    if nama.endswith(".docx"):
        return baca_docx(data)
    if nama.endswith(".xlsx"):
        return baca_excel(data)
    if nama.endswith(".csv"):
        return baca_csv(data)
    return baca_pdf(data)


# ---------- Layanan AI ----------

@st.cache_resource
def get_client() -> genai.Client:
    api_key = os.environ.get("GEMINI_API_KEY") or st.secrets["GEMINI_API_KEY"]
    return genai.Client(api_key=api_key)


def buat_chat():
    return get_client().chats.create(
        model=MODEL,
        config=types.GenerateContentConfig(system_instruction=SYSTEM_INSTRUCTION),
    )


def kirim_stream(chat, materi_awal: str | None, pertanyaan: str) -> Iterator[str]:
    # Materi hanya dikirim pada pesan pertama. Setelahnya, Gemini mengingat konteksnya.
    isi = f"{materi_awal}\n\nPertanyaan: {pertanyaan}" if materi_awal else pertanyaan
    for potongan in chat.send_message_stream(isi):
        if potongan.text:
            yield potongan.text


# ---------- Antarmuka ----------

@st.cache_data(show_spinner=False)
def ekstrak_materi(nama: str, data: bytes) -> str:
    return baca_dokumen(nama, data)


st.title("Asisten Akuntansi")

files = st.file_uploader(
    "Unggah materi (PDF, PPTX, DOCX, XLSX, CSV), opsional",
    type=["pdf", "pptx", "docx", "xlsx", "csv"],
    accept_multiple_files=True,
)

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

# Tombol unduh percakapan.
if st.session_state.riwayat:
    teks = "\n\n".join(
        f"**{'Pengguna' if p['role'] == 'user' else 'Asisten'}:**\n\n{p['isi']}"
        for p in st.session_state.riwayat
    )
    st.download_button(
        "Unduh percakapan (.md)",
        data=teks,
        file_name="percakapan-asisten-akuntansi.md",
        mime="text/markdown",
    )
        mime="text/markdown",
    )
