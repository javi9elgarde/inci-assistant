import os
import json
import glob
from datetime import datetime
from flask import Flask, render_template, request, jsonify
from groq import Groq
import docx
import PyPDF2
try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

app = Flask(__name__)

@app.after_request
def add_cors(response):
    response.headers["Access-Control-Allow-Origin"] = "*"
    response.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    response.headers["Access-Control-Allow-Headers"] = "Content-Type"
    return response

@app.route("/api/analyze", methods=["OPTIONS"])
def analyze_options():
    return "", 204

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


def build_category_tree():
    """Construye el árbol N1>N2>N3 desde el Excel."""
    xlsx = os.path.join(os.path.dirname(__file__), "docs", "CategorizacionesIncidenciaseCommerceOK.xlsx")
    if not HAS_OPENPYXL or not os.path.exists(xlsx):
        return {}
    try:
        wb = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)
        ws = wb['CategorizacioneseCommerce']
        tree = {}
        for row in ws.iter_rows(min_row=6, max_row=831, min_col=2, max_col=5, values_only=True):
            b, c, d, e = row
            if not b:
                continue
            n1 = str(b).strip()
            n2 = str(c).strip() if c else ''
            n3 = str(d).strip() if d else ''
            if n1 not in tree:
                tree[n1] = {}
            if n2 and n2 not in tree[n1]:
                tree[n1][n2] = []
            if n2 and n3 and n3 not in tree[n1][n2]:
                tree[n1][n2].append(n3)
        return tree
    except Exception:
        return {}


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


@app.route("/api/categories", methods=["GET"])
def categories():
    tree = build_category_tree()
    return jsonify(tree)


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
    servicio = data.get("servicio", "").strip()
    criticidad = data.get("criticidad", "").strip()
    cat_n1 = data.get("cat_n1", "").strip()
    cat_n2 = data.get("cat_n2", "").strip()
    cat_n3 = data.get("cat_n3", "").strip()
    pistas = data.get("pistas", "").strip()

    cat_operacional = " > ".join(filter(None, [cat_n1, cat_n2, cat_n3]))
    es_alta_critica = criticidad in ["Alta", "Crítica"]

    if not titulo and not cat_operacional:
        return jsonify({"error": "Introduce al menos el título o la categoría operacional"}), 400

    fuera_horario = not is_office_hours()
    now = datetime.now()
    hora_actual = now.strftime("%H:%M")
    dia_semana = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"][now.weekday()]

    # Filtrar solo las filas relevantes de la matriz según la categoría seleccionada
    routing_matrix = load_routing_matrix()
    relevant_rows = []
    if routing_matrix and cat_operacional:
        for line in routing_matrix.splitlines():
            if any(part.lower() in line.lower() for part in [cat_n1, cat_n2, cat_n3] if part):
                relevant_rows.append(line)
    elif routing_matrix:
        # Sin categoría, enviar solo las primeras 80 líneas como muestra
        relevant_rows = routing_matrix.splitlines()[:80]

    docs_context = ""
    if relevant_rows:
        docs_context += f"\n\nREGLAS DE ROUTING RELEVANTES (del Excel de categorizaciones):\n" + "\n".join(relevant_rows)
    if not docs_context:
        docs_context = "\n\n[Sin reglas de routing específicas. Usa criterio general.]"

    incidencia_text = f"""TÍTULO: {titulo}
SERVICIO: {servicio}
CRITICIDAD: {criticidad if criticidad else 'No especificada'}
CATEGORÍA OPERACIONAL: {cat_operacional}
PISTAS ADICIONALES: {pistas if pistas else 'Ninguna'}"""

    horario_text = (
        f"HORARIO: Fuera de horario de oficina ({dia_semana} {hora_actual}). "
        "Si existe guardia, proporciona el número de contacto."
        if fuera_horario else
        f"HORARIO: Dentro de horario de oficina ({dia_semana} {hora_actual})."
    )

    prompt = f"""Eres un asistente experto en gestión de incidencias IT del equipo Sala Ebusiness WEB/BBDD (nivel 1) de Kyndryl/ECI.

REGLAS DE NEGOCIO (MUY IMPORTANTES):
1. Nosotros somos nivel 1 (Sala Ebusiness WEB/BBDD). Siempre pasamos al grupo de nivel 2 indicado en las reglas de routing.
2. El nombre del grupo debe extraerse del formato "Empresa - Org - NOMBRE_GRUPO": usa solo el NOMBRE_GRUPO final.
   Ejemplo: "El Corte Inglés, S.A. - Mantenimiento - Run The Business" → grupo = "Run The Business"
   Ejemplo: "Hiberus Digital Business, S.L. - Mantenimiento - Soporte Customer Experience" → grupo = "Soporte Customer Experience"
3. LÓGICA DE HORARIO Y CRITICIDAD:
   - DENTRO de horario de oficina (L-V 8:00-18:00): SIEMPRE asignar al grupo sin llamar, sea la criticidad que sea.
   - FUERA de horario (tardes +18h, fines de semana, festivos):
     * Si criticidad es Alta o Crítica: asignar Y avisar/llamar al grupo.
     * Si criticidad es Media o Baja: solo asignar, NO llamar.
4. Grupos que pertenecen a "Run The Business" (en horario o si no es Alta/Crítica, se pasa a Run The Business en vez del subgrupo):
   Soporte Customer Experience, Soporte Firefly Checkout, Soporte Cesta-Checkout, Front Cesta-Checkout, y similares de ECI/Accenture/Hiberus relacionados con web ecommerce.
   - Fuera de horario + Alta/Crítica → pasar directamente al subgrupo específico.
   - Resto → pasar a "Run The Business".
5. CUANDO CATEGORÍA ES "Servicio sin identificar" O ESTÁ VACÍA:
   Ignora la categoría y busca pistas en el TÍTULO y la DESCRIPCIÓN:
   - "openshift", "kubernetes", "k8s", "prometheus", "collector", "devops" → grupo = "DevOps / OpenShift"
   - "elastic", "elasticsearch", "kibana", "ELK" → grupo = "Elastic / Monitorización"
   - "linux", "CPU", "memoria", "disco", "mongod", "mongo" → grupo de infraestructura Unix/Linux
   - "oracle", "bbdd", "database", "sql" → grupo DBA
   - "kafka", "rabbit", "broker", "WMB" → grupo de integración/mensajería
   - Si el título contiene "InstanceId: XXXX" o "Component: XXXX", ese XXXX indica el sistema afectado
   - Usa el contexto del título completo para deducir el grupo más apropiado

Responde SIEMPRE en este formato JSON exacto (sin markdown, solo JSON puro):
{{
  "grupo_responsable": "Nombre limpio del grupo de nivel 2",
  "grupo_via": "Run The Business (si aplica la regla) o null",
  "confianza_grupo": "alta|media|baja",
  "fuera_horario": true,
  "accion": "asignar_y_llamar | solo_asignar",
  "motivo_accion": "Explicación breve de por qué llamar o no",
  "guardia": {{
    "aplica": true,
    "nombre": "Nombre del servicio de guardia o null",
    "contacto": "Número o contacto si lo conoces, si no null"
  }},
  "resumen_diagnostico": "Qué está pasando según la categoría y pistas",
  "notas_adicionales": "Cualquier info relevante adicional o null"
}}

{horario_text}
CRITICIDAD DE LA INCIDENCIA: {criticidad if criticidad else 'No especificada'}

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
