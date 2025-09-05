import os, tempfile, mimetypes, subprocess, glob
from flask import Flask, request, jsonify
from markitdown import MarkItDown
import google.generativeai as genai

# --- Gemini (OCR) ---
GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY", "")
if GOOGLE_API_KEY:
    genai.configure(api_key=GOOGLE_API_KEY)
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-1.5-flash")
MIN_TEXT_CHARS = int(os.getenv("MIN_TEXT_CHARS", "80"))
PDF_DPI = int(os.getenv("PDF_DPI", "200"))

app = Flask(__name__)

def _ok_ext(name: str) -> bool:
    name = (name or "").lower()
    return name.endswith(".ppt") or name.endswith(".pptx")

def _run(cmd: list):
    return subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

def ppt_any_to_pdf(src_path: str, outdir: str) -> str:
    _run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", outdir, src_path])
    base = os.path.splitext(os.path.basename(src_path))[0]
    pdf_path = os.path.join(outdir, base + ".pdf")
    if not os.path.exists(pdf_path):
        pdfs = sorted(glob.glob(os.path.join(outdir, "*.pdf")))
        if not pdfs:
            raise RuntimeError("Falha na conversão para PDF")
        pdf_path = pdfs[0]
    return pdf_path

def pdf_to_images(pdf_path: str, outdir: str) -> list:
    prefix = os.path.join(outdir, "slide")
    _run(["pdftoppm", "-png", "-r", str(PDF_DPI), pdf_path, prefix])
    return sorted(glob.glob(prefix + "-*.png"))

def gemini_ocr(img_bytes: bytes, mime: str) -> str:
    model = genai.GenerativeModel(GEMINI_MODEL)
    parts = [
        {"text": "Transcreva fielmente todo o texto visível nesta imagem. "
                 "Apenas texto; sem descrição. Preserve quebras e ordem."},
        {"inline_data": {"mime_type": mime, "data": img_bytes}},
    ]
    resp = model.generate_content(parts)
    return (getattr(resp, "text", "") or "").strip()

def ocr_pdf_with_gemini(pdf_path: str) -> str:
    out = []
    with tempfile.TemporaryDirectory() as tdir:
        images = pdf_to_images(pdf_path, tdir)
        for i, img in enumerate(images, 1):
            with open(img, "rb") as f:
                out.append(f"# Slide {i}\n{gemini_ocr(f.read(), 'image/png')}".strip())
    return "\n\n".join(out).strip()

def pptx_native_text(path: str) -> str:
    md = MarkItDown()
    res = md.convert(path)
    return (getattr(res, "text_content", "") or "").strip()

def pdf_native_text(path: str) -> str:
    md = MarkItDown()
    res = md.convert(path)
    return (getattr(res, "text_content", "") or "").strip()

@app.get("/health")
def health():
    return {"ok": True}

@app.post("/ppt2md")
def ppt2md():
    """
    Transcreve PPT/PPTX -> Markdown.
    OCR modes: ?ocr=auto|force|off  (padrão: auto)
    """
    ocr_mode = request.args.get("ocr", "auto").lower()
    ctype = request.headers.get("Content-Type", "")
    filename = request.headers.get("X-Filename", "arquivo.pptx")

    if "multipart/form-data" in ctype and "file" in request.files:
        f = request.files["file"]; data = f.read(); filename = f.filename or filename
    else:
        data = request.get_data()

    if not data:
        return jsonify({"error": "Nenhum arquivo recebido"}), 400
    if not _ok_ext(filename):
        return jsonify({"error": "Envie .ppt ou .pptx"}), 415

    try:
        with tempfile.TemporaryDirectory() as tdir:
            src = os.path.join(tdir, filename)
            with open(src, "wb") as fp: fp.write(data)

            ext = filename.lower().split(".")[-1]
            # 1) tenta texto nativo
            if ext == "pptx":
                native = pptx_native_text(src)
            else:
                native = pdf_native_text(ppt_any_to_pdf(src, tdir))

            do_ocr = (ocr_mode == "force") or (ocr_mode == "auto" and len(native) < MIN_TEXT_CHARS)

            if do_ocr:
                if not GOOGLE_API_KEY:
                    return jsonify({"error": "OCR necessário, mas GOOGLE_API_KEY não configurada",
                                    "hint": "defina GOOGLE_API_KEY ou use ?ocr=off"}), 500
                txt = ocr_pdf_with_gemini(ppt_any_to_pdf(src, tdir))
                engine, strategy = "gemini-ocr", ("ocr-forced" if ocr_mode == "force" else "ocr-fallback")
            else:
                txt, engine, strategy = native, "markitdown-native", "native-only"

            return jsonify({
                "format": ext, "engine": engine, "strategy": strategy,
                "chars": len(txt), "content_markdown": txt
            }), 200

    except subprocess.CalledProcessError as e:
        return jsonify({"error": "Erro em conversão externa", "detail": e.stderr.decode('utf-8','ignore')}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", "5000")))
