import os

import streamlit as st
from google import genai
from google.genai import types

MODEL = "gemini-3.5-flash-lite"

SYSTEM_INSTRUCTION = """Kamu adalah akuntan senior dan analis bisnis yang membantu mahasiswa akuntansi dan profesional.

Aturan menjawab:
- Jawab hanya berdasarkan materi yang diberikan. Jika informasinya tidak ada di materi, katakan dengan jelas bahwa materi tidak memuat informasi tersebut. Jangan mengarang angka.
- Sebutkan nomor halaman atau slide sebagai sumber jika relevan.
- Untuk perhitungan, tampilkan rumus, langkah, dan hasil akhirnya.
- Tulis angka dengan presisi dan satuan yang jelas.
- Gunakan bahasa Indonesia yang formal namun mudah dipahami."""


@st.cache_resource
def get_client() -> genai.Client:
    api_key = os.environ.get("GEMINI_API_KEY") or st.secrets["GEMINI_API_KEY"]
    return genai.Client(api_key=api_key)


def buat_chat():
    return get_client().chats.create(
        model=MODEL,
        config=types.GenerateContentConfig(system_instruction=SYSTEM_INSTRUCTION),
    )


def kirim(chat, materi_awal: str | None, pertanyaan: str) -> str:
    # Materi hanya dikirim pada pesan pertama. Pesan berikutnya
    # cukup membawa pertanyaan, karena Gemini sudah mengingat konteksnya.
    isi = f"{materi_awal}\n\nPertanyaan: {pertanyaan}" if materi_awal else pertanyaan
    return chat.send_message(isi).text
