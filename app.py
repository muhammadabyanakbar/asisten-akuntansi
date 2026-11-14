import os
import streamlit as st
from pypdf import PdfReader
from google import genai

api_key = os.environ.get("GEMINI_API_KEY") or st.secrets["GEMINI_API_KEY"]
client = genai.Client(api_key=api_key)

st.title("Asisten Akuntansi")

file = st.file_uploader("Unggah materi (PDF)", type="pdf")

if file:
    reader = PdfReader(file)
    teks = []
    for nomor, halaman in enumerate(reader.pages, start=1):
        teks.append(f"Halaman {nomor}:")
        teks.append(halaman.extract_text() or "")
    materi = "\n".join(teks)
    st.success("Materi berhasil dibaca.")

    pertanyaan = st.text_input("Tulis pertanyaan lu:")
    if st.button("Tanya") and pertanyaan:
        with st.spinner("Sedang menjawab..."):
            response = client.models.generate_content(
                model="gemini-3.5-flash-lite",
                contents=materi + "\n\nPertanyaan: " + pertanyaan + "\nJawab berdasarkan materi di atas.",
            )
        st.write(response.text)