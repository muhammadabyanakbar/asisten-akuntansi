import os
import streamlit as st
from pypdf import PdfReader
from pptx import Presentation
from google import genai

api_key = os.environ.get("GEMINI_API_KEY") or st.secrets["GEMINI_API_KEY"]
client = genai.Client(api_key=api_key)

st.title("Asisten Akuntansi")

files = st.file_uploader(
    "Unggah materi (PDF atau PPTX)",
    type=["pdf", "pptx"],
    accept_multiple_files=True,
)


def baca_pdf(file):
    reader = PdfReader(file)
    hasil = []
    for nomor, halaman in enumerate(reader.pages, start=1):
        hasil.append(f"Halaman {nomor}:")
        hasil.append(halaman.extract_text() or "")
    return "\n".join(hasil)


def baca_pptx(file):
    prs = Presentation(file)
    hasil = []
    for nomor, slide in enumerate(prs.slides, start=1):
        hasil.append(f"Slide {nomor}:")
        for shape in slide.shapes:
            if shape.has_text_frame:
                hasil.append(shape.text_frame.text)
    return "\n".join(hasil)


if files:
    teks = []
    for file in files:
        teks.append(f"=== Materi: {file.name} ===")
        if file.name.lower().endswith(".pptx"):
            teks.append(baca_pptx(file))
        else:
            teks.append(baca_pdf(file))
    materi = "\n".join(teks)
    st.success(f"{len(files)} materi berhasil dibaca.")

    pertanyaan = st.text_input("Tulis pertanyaan lu:")
    if st.button("Tanya") and pertanyaan:
        with st.spinner("Sedang menjawab..."):
            response = client.models.generate_content(
                model="gemini-3.5-flash-lite",
                contents=materi + "\n\nPertanyaan: " + pertanyaan + "\nJawab berdasarkan materi di atas.",
            )
        st.write(response.text)
