import os

import streamlit as st
from google import genai

MODEL = "gemini-3.5-flash-lite"


@st.cache_resource
def get_client() -> genai.Client:
    api_key = os.environ.get("GEMINI_API_KEY") or st.secrets["GEMINI_API_KEY"]
    return genai.Client(api_key=api_key)


def tanya_materi(materi: str, pertanyaan: str) -> str:
    response = get_client().models.generate_content(
        model=MODEL,
        contents=(
            f"{materi}\n\n"
            f"Pertanyaan: {pertanyaan}\n"
            "Jawab berdasarkan materi di atas."
        ),
    )
    return response.text