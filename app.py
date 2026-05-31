import os
import json
import glob
from datetime import datetime
from flask import Flask, render_template, request, jsonify
from groq import Groq
import docx
import PyPDF2

app = Flask(__name__)

DOCS_FOLDER = os.path.join(os.path.dirname(__file__), "docs")
API_KEY_FILE = os.path.join(os.path.dirname(__file__), "api_key.txt")

OFFICE_START = 8
OFFICE_END = 18
OFFICE_DAYS = [0, 1, 2, 3, 4]  # Lunes=0 ... Viernes=4


def get_api_key():
    # En Railway la key viene como variable de entorno
    env_key = os.environ.get("GROQ_API_KEY", "")
    if env_key:
        return env_key
    if os.path.exists(API_KEY_FILE):
        with open(API_KEY_FILE) as f:
            return f.read().strip()
    return ""


def save_api_key(key):
    with open(API_KEY_FILE, "w") as f:
        f.write(key.strip())


def is_office_hours():
    now = datetime.now()
    return now.weekday() in OFFICE_DAYS and OFFICE_START <= now.hour < OFFICE_END


def read_docx(path):
    try:
        doc = docx.Document(path)
        return "\n".join(p.text for p in doc.paragraphs if p.text.strip())
    except Exception:
        return ""


def read_pdf(path):
    try:
        text = []
        with open(path, "rb") as f:
            reader = PyPDF2.PdfReader(f)
            for page in reader.pages:
                text.append(page.extract_text() or "")
        return "\n".join(text)
    except Exception:
        return ""


def read_txt(path):
    try:
        with open(path, encoding="utf-8", errors="ignore") as f:
            return f.read()
    except Exception:
        return ""


def load_routing_matrix():
    """Carga la matriz de routing del Excel convertida a texto."""
    path = os.path.join(os.path.dirname(__file__), "knowledge.txt")
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            return f.read()
    return ""


def load_knowledge_base():
    docs = []
    patterns = ["**/*.docx", "**/*.pdf", "**/*.txt", "**/*.md"]
    for pattern in patterns:
        for filepath in glob.glob(os.path.join(DOCS_FOLDER, pattern), recursive=True):
            name = os.path.basename(filepath)
            ext = os.path.splitext(filepath)[1].lower()
            if ext == ".docx":
                content = read_docx(filepath)
            elif ext == ".pdf":
                content = read_pdf(filepath)
            else:
                content = read_txt(filepath)
            if content.strip():
                docs.append(f"=== DOCUMENTO: {name} ===\n{content}\n")
    return "\n".join(docs) if docs else ""


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/check-key", methods=["GET"])
def check_key():
    key = get_api_key()
    return jsonify({"has_key": bool(key)})


@app.route("/api/save-key", methods=["POST"])
def save_key():
    data = request.json
    key = data.get("key", "").strip()
    if not key:
        return jsonify({"ok": False, "error": "Clave vacía"})
    save_api_key(key)
    return jsonify({"ok": True})


@app.route("/api/docs-status", methods=["GET"])
def docs_status():
    docs = []
    for pattern in ["**/*.docx", "**/*.pdf", "**/*.txt", "**/*.md"]:
        for f in glob.glob(os.path.join(DOCS_FOLDER, pattern), recursive=True):
            docs.append(os.path.basename(f))
    return jsonify({"docs": docs, "folder": DOCS_FOLDER})


@app.route("/api/analyze", methods=["POST"])
def analyze():
    api_key = get_api_key()
    if not api_key:
        return jsonify({"error": "No hay API key configurada"}), 400

    data = request.json
    titulo = data.get("titulo", "").strip()
    descripcion = data.get("descripcion", "").strip()
    servicio = data.get("servicio", "").strip()
    categoria = data.get("categoria", "").strip()
    comentarios = data.get("comentarios", "").strip()

    if not titulo and not descripcion:
        return jsonify({"error": "Introduce al menos título o descripción"}), 400

    fuera_horario = not is_office_hours()
    now = datetime.now()
    hora_actual = now.strftime("%H:%M")
    dia_semana = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"][now.weekday()]

    routing_matrix = load_routing_matrix()
    knowledge = load_knowledge_base()

    docs_context = ""
    if routing_matrix:
        docs_context += f"\n\nMATRIZ DE ROUTING (usa esto como referencia principal para asignar grupos):\n{routing_matrix}"
    if knowledge:
        docs_context += f"\n\nDOCUMENTACIÓN ADICIONAL:\n{knowledge}"
    if not docs_context:
        docs_context = "\n\n[Sin matriz de routing cargada. Usa conocimiento general.]"

    incidencia_text = f"""TÍTULO: {titulo}
SERVICIO/CATEGORÍA: {servicio} / {categoria}
DESCRIPCIÓN:
{descripcion}
COMENTARIOS PREVIOS:
{comentarios if comentarios else 'Ninguno'}"""

    horario_text = (
        f"HORARIO: Fuera de horario de oficina ({dia_semana} {hora_actual}). "
        "Si existe guardia, proporciona el número de contacto."
        if fuera_horario else
        f"HORARIO: Dentro de horario de oficina ({dia_semana} {hora_actual})."
    )

    prompt = f"""Eres un asistente experto en gestión de incidencias IT para el equipo de Kyndryl/ECI (El Corte Inglés).
Analiza la siguiente incidencia y responde SIEMPRE en este formato JSON exacto (sin markdown, solo JSON puro):

{{
  "grupo_responsable": "Nombre del grupo que debe gestionar la incidencia",
  "confianza_grupo": "alta|media|baja",
  "fuera_horario": true,
  "guardia": {{
    "aplica": true,
    "nombre": "Nombre del servicio de guardia o null",
    "contacto": "Número o contacto o null"
  }},
  "puede_resolver": true,
  "pasos_resolucion": ["paso 1", "paso 2"],
  "resumen_diagnostico": "Explicación breve de qué está pasando y por qué",
  "escalado_recomendado": "A quién escalar si no se resuelve, o null",
  "notas_adicionales": "Cualquier info relevante adicional o null"
}}

{horario_text}

INCIDENCIA A ANALIZAR:
{incidencia_text}
{docs_context}"""

    try:
        client = Groq(api_key=api_key)
        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=1500,
            temperature=0.2,
        )
        raw = response.choices[0].message.content.strip()

        # Limpiar markdown si viene con bloques de código
        if "```" in raw:
            parts = raw.split("```")
            for part in parts:
                part = part.strip()
                if part.startswith("json"):
                    part = part[4:].strip()
                try:
                    result = json.loads(part)
                    result["fuera_horario"] = fuera_horario
                    result["hora_consulta"] = f"{dia_semana} {hora_actual}"
                    return jsonify(result)
                except Exception:
                    continue

        result = json.loads(raw)
        result["fuera_horario"] = fuera_horario
        result["hora_consulta"] = f"{dia_semana} {hora_actual}"
        return jsonify(result)

    except json.JSONDecodeError:
        return jsonify({"error": "Error parseando respuesta de IA", "raw": raw}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    print("=" * 50)
    print("  Asistente de Incidencias — EasyWEB")
    print("  Abre tu navegador en: http://localhost:5000")
    print("=" * 50)
    app.run(debug=False, port=5000)
