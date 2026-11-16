from io import BytesIO

from pptx import Presentation
from pypdf import PdfReader


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


def baca_dokumen(nama: str, data: bytes) -> str:
    if nama.lower().endswith(".pptx"):
        return baca_pptx(data)
    return baca_pdf(data)
