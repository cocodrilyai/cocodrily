import asyncio
import ctypes
import base64
import io
import os
import sys
import threading
import time
import urllib.parse
from http.server import BaseHTTPRequestHandler, HTTPServer
import edge_tts
from gtts import gTTS
import requests
from google import genai

# --- CONFIGURACIÓN DE CLAVES ---
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "AQ.Ab8RN6LZ98ckWn8YiZeBcJkKpOeCX-YTsQbjXXfnKdDUpkRpEg")
ELEVENLABS_API_KEY = os.environ.get("ELEVENLABS_API_KEY", "sk_a875d0a08d684229b591825019dd115bd2f8e21c1060de52")

PUERTO_SERVIDOR = int(os.environ.get("PORT", 5000))

MOTOR_ACTUAL = "edge"
INDICE_VOZ_ACTUAL = 0
COCODRILO_VIVO = True  
GENERO_ACTUAL = "cocodrily"
ULTIMO_AUDIO_BASE64 = ""

try:
    gemini_client = genai.Client(api_key=GEMINI_API_KEY) if GEMINI_API_KEY else None
except Exception:
    gemini_client = None

voces_sapi_instaladas = ["Microsoft Helena", "Microsoft Pablo"]
try:
    import win32com.client
    ctypes.windll.ole32.CoInitialize(None)
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    voces_sapi_instaladas = [voice.GetDescription() for voice in speaker.GetVoices()]
    ctypes.windll.ole32.CoUninitialize()
except Exception:
    pass

voces_edge_cache = {"mujeres": [], "hombres": []}
def cargar_todas_voces_edge():
    global voces_edge_cache
    try:
        async def _obtener():
            return await edge_tts.list_voices()
        try:
            loop = asyncio.get_event_loop()
            if loop.is_running():
                todas = asyncio.run_coroutine_threadsafe(_obtener(), loop).result(timeout=10)
            else:
                todas = asyncio.run(_obtener())
        except Exception:
            todas = asyncio.run(_obtener())

        voces_es = [v for v in todas if v.get("Locale", "").startswith("es-")]
        lista_completa = [{"nombre": v["ShortName"], "rate": "+0%", "pitch": "+0Hz"} for v in voces_es]
        if not lista_completa:
            lista_completa = [{"nombre": "es-ES-ElviraNeural", "rate": "+0%", "pitch": "+0Hz"}]
        voces_edge_cache["mujeres"] = lista_completa
    except Exception:
        voces_edge_cache["mujeres"] = [{"nombre": "es-ES-ElviraNeural", "rate": "+0%", "pitch": "+0Hz"}]

cargar_todas_voces_edge()

voces_elevenlabs_cache = {"mujeres": [], "hombres": []}
def cargar_todas_voces_elevenlabs():
    global voces_elevenlabs_cache
    if not ELEVENLABS_API_KEY or ELEVENLABS_API_KEY.startswith("sk_..."):
        return
    try:
        url = "https://api.elevenlabs.io/v1/voices"
        headers = {"xi-api-key": ELEVENLABS_API_KEY}
        response = requests.get(url, headers=headers, timeout=10)
        if response.status_code == 200:
            data = response.json()
            voces_elevenlabs_cache["mujeres"] = [{"nombre": v.get("name", "Desconocida"), "id": v.get("voice_id", "")} for v in data.get("voices", [])]
    except Exception:
        pass

cargar_todas_voces_elevenlabs()

voces_por_servicio = {
    "edge": voces_edge_cache,
    "elevenlabs": voces_elevenlabs_cache,
    "local": {"mujeres": voces_sapi_instaladas, "hombres": []},
    "gtts": {"mujeres": ["es-ES (España)", "es-MX (México)", "es-US (Estados Unidos)", "es-AR (Argentina)"], "hombres": []}
}

def obtener_nombre_actual():
    return "Cocodrila" if GENERO_ACTUAL == "cocodrila" else "Cocodrily"

def preparar_audio_base64(buffer_bytes):
    global ULTIMO_AUDIO_BASE64
    try:
        if not buffer_bytes or not isinstance(buffer_bytes, io.BytesIO):
            return
        buffer_bytes.seek(0)
        audio_bytes = buffer_bytes.read()
        if len(audio_bytes) > 50:
            ULTIMO_AUDIO_BASE64 = base64.b64encode(audio_bytes).decode("utf-8")
    except Exception as e:
        print(f"⚠️ [AUDIO ERROR]: {e}")

def _hablar_elevenlabs_bytes(texto):
    todos = voces_por_servicio["elevenlabs"]["mujeres"]
    if not todos:
        return
    item = todos[INDICE_VOZ_ACTUAL % len(todos)]
    voice_id = item["id"]
    url = f"https://api.elevenlabs.io/v1/text-to-speech/{voice_id}"
    headers = {"Accept": "audio/mpeg", "Content-Type": "application/json", "xi-api-key": ELEVENLABS_API_KEY}
    data = {"text": texto, "model_id": "eleven_multilingual_v2", "voice_settings": {"stability": 0.5, "similarity_boost": 0.75}}
    try:
        response = requests.post(url, json=data, headers=headers, timeout=15)
        if response.status_code == 200:
            preparar_audio_base64(io.BytesIO(response.content))
    except Exception:
        pass

def _hablar_edge_bytes(texto):
    elementos = voces_por_servicio["edge"]["mujeres"]
    voz_nombre = elementos[INDICE_VOZ_ACTUAL % len(elementos)]["nombre"] if elementos else "es-ES-ElviraNeural"

    async def _gen():
        comm = edge_tts.Communicate(texto, voz_nombre)
        data = bytearray()
        async for chunk in comm.stream():
            if chunk["type"] == "audio":
                data.extend(chunk["data"])
        return bytes(data)

    try:
        try:
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            ab = loop.run_until_complete(_gen())
            loop.close()
        except Exception:
            ab = asyncio.run(_gen())

        if ab and len(ab) > 100:
            preparar_audio_base64(io.BytesIO(ab))
    except Exception as e:
        print(f"⚠️ [EDGE TTS ERROR]: {e}")

def generar_audio_web(texto):
    if not texto or not texto.strip():
        return
    try:
        if MOTOR_ACTUAL == "elevenlabs":
            _hablar_elevenlabs_bytes(texto)
        elif MOTOR_ACTUAL == "edge":
            _hablar_edge_bytes(texto)
        elif MOTOR_ACTUAL == "gtts":
            buffer_ram = io.BytesIO()
            tts = gTTS(text=texto, lang="es", slow=False)
            tts.write_to_fp(buffer_ram)
            buffer_ram.seek(0)
            preparar_audio_base64(buffer_ram)
        else:
            _hablar_edge_bytes(texto)
    except Exception:
        pass

def procesar_inteligencia(prompt):
    global COCODRILO_VIVO
    texto_lower = prompt.lower()

    if "muérete" in texto_lower or "muere" in texto_lower:
        COCODRILO_VIVO = False
        return "..."
    if "pedo" in texto_lower or "gas" in texto_lower:
        return "¡Prrrt! Ups, ¡fue el cocodrilo!"
    if not COCODRILO_VIVO:
        return ""

    rol_genero = "Eres Cocodrila, alegre y divertida." if GENERO_ACTUAL == "cocodrila" else "Eres Cocodrily, amigable y curioso."
    
    if gemini_client:
        try:
            response = gemini_client.models.generate_content(
                model="gemini-2.5-flash",
                contents=f"{prompt}\n({rol_genero} Responde súper corto, amigable y directo en español, máximo 2 frases.)",
            )
            return response.text
        except Exception:
            return "¡Vaya, me quedé pensando un segundo!"
    else:
        return "¡Hola! Configura tu API key de Gemini."

class ServidorSistema(BaseHTTPRequestHandler):
    def do_GET(self):
        global MOTOR_ACTUAL, INDICE_VOZ_ACTUAL, COCODRILO_VIVO, ULTIMO_AUDIO_BASE64
        try:
            parsed_path = urllib.parse.urlparse(self.path)
            path = parsed_path.path

            if path == "/audio-nuevo":
                self.send_response(200)
                self.send_header("Content-type", "text/plain")
                self.end_headers()
                if ULTIMO_AUDIO_BASE64:
                    self.wfile.write(ULTIMO_AUDIO_BASE64.encode("utf-8"))
                    ULTIMO_AUDIO_BASE64 = "" # Se consume al instante para que nunca se repita
                return

            if path.startswith("/hablar/"):
                text = urllib.parse.unquote(path.split("/hablar/")[1])
                if text.strip():
                    resp = procesar_inteligencia(text)
                    threading.Thread(target=generar_audio_web, args=(resp,)).start()
                self.send_response(200)
                self.end_headers()
                return

            elif path.startswith("/revivir"):
                COCODRILO_VIVO = True
                self.send_response(303)
                self.send_header("Location", "/")
                self.end_headers()
                return

            elif path.startswith("/matar"):
                COCODRILO_VIVO = False
                self.send_response(303)
                self.send_header("Location", "/")
                self.end_headers()
                return

            elif path.startswith("/motor/"):
                MOTOR_ACTUAL = path.split("/")[-1]
                INDICE_VOZ_ACTUAL = 0
                self.send_response(303)
                self.send_header("Location", "/")
                self.end_headers()
                return

            elif path.startswith("/voz/"):
                try:
                    INDICE_VOZ_ACTUAL = int(path.split("/")[-1])
                except Exception:
                    pass
                self.send_response(303)
                self.send_header("Location", "/")
                self.end_headers()
                return

            self.send_response(200)
            self.send_header("Content-type", "text/html; charset=utf-8")
            self.end_headers()

            nombre_activo = obtener_nombre_actual()
            estado_texto = "🟢 VIVO Y ACTIVO (AUTOMÁTICO)" if COCODRILO_VIVO else "💀 MODO MUERTO"
            color_estado = "#4ade80" if COCODRILO_VIVO else "#f87171"

            voces_motor_actual = voces_por_servicio[MOTOR_ACTUAL]["mujeres"]
            if MOTOR_ACTUAL in ["edge", "elevenlabs"]:
                nombres_voces = [v["nombre"] for v in voces_motor_actual]
            else:
                nombres_voces = voces_motor_actual

            total_voces = len(nombres_voces)
            voces_html = "".join([f'<a href="/voz/{i}" class="btn {"btn-activo" if INDICE_VOZ_ACTUAL == i else ""}">{i+1}. {v}</a>\n' for i, v in enumerate(nombres_voces)])
            motores_html = "".join([f'<a href="/motor/{m}" class="btn btn-engine {"btn-activo" if MOTOR_ACTUAL==m else ""}">{m.upper()}</a> ' for m in voces_por_servicio.keys()])

            html = """<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Panel Web - """ + nombre_activo.upper() + """</title>
    <style>
        body { background-color: #000; color: #0ff; font-family: monospace; text-align: center; margin: 8px; }
        .container { max-width: 900px; margin: auto; background: #050505; padding: 18px; border-radius: 12px; border: 2px solid #00f2fe; }
        .btn { display: inline-block; margin: 2px; padding: 5px 8px; font-size: 10px; text-decoration: none; color: white; background: #334155; border-radius: 3px; }
        .btn-engine { background: #7c3aed; font-size: 10px; margin: 3px; padding: 7px 10px; text-decoration: none; color: white; border-radius: 4px; display: inline-block; font-weight: bold; }
        .btn-ai { background: #2563eb; margin: 3px; padding: 8px 14px; text-decoration: none; color: white; border-radius: 4px; display: inline-block; font-weight: bold; }
        .btn-dead { background: #dc2626; margin: 3px; padding: 8px 14px; text-decoration: none; color: white; border-radius: 4px; display: inline-block; font-weight: bold; }
        .btn-activo { background: #16a34a !important; font-weight: bold; border: 1px solid #86efac; }
        .section { margin-top: 15px; border-top: 1px solid #334155; padding-top: 10px; text-align: left; }
        .voice-group { background: #111827; padding: 10px; border-radius: 6px; margin-bottom: 8px; max-height: 250px; overflow-y: auto; }
    </style>
    <script>
        let recognition = null;
        const audioElem = new Audio();

        setInterval(async () => {
            try {
                let res = await fetch("/audio-nuevo");
                if (res.ok) {
                    let b64 = await res.text();
                    if (b64.length > 50) {
                        audioElem.src = "data:audio/mp3;base64," + b64;
                        audioElem.play().catch(e => console.log("Audio esperando toque en pantalla"));
                    }
                }
            } catch(e) {}
        }, 1500);

        function iniciarEscuchaContinua() {
            const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
            if (!SpeechRecognition) return;
            recognition = new SpeechRecognition();
            recognition.lang = 'es-ES';
            recognition.continuous = true;
            recognition.interimResults = false;
            
            recognition.onstart = function() {
                document.getElementById("estadoMic").innerText = "🟢 Micrófono Escuchando...";
                document.getElementById("estadoMic").style.background = "#16a34a";
            };
            recognition.onresult = function(event) {
                const speechResult = event.results[event.results.length - 1][0].transcript.trim();
                if (speechResult.length > 0) {
                    document.getElementById("textoTranscrito").innerText = "Última voz: " + speechResult;
                    fetch("/hablar/" + encodeURIComponent(speechResult));
                }
            };
            recognition.onend = function() { 
                setTimeout(() => { try { recognition.start(); } catch(e) {} }, 500); 
            };
            try { recognition.start(); } catch(e) {}
        }

        window.onload = function() { 
            // Desbloqueo automático de audio al hacer cualquier clic inicial en la página
            document.body.addEventListener('click', () => {
                if (audioElem.paused && audioElem.src) { audioElem.play().catch(e=>{}); }
            }, { once: true });
            iniciarEscuchaContinua(); 
        };
    </script>
</head>
<body>
    <div class="container">
        <h1>🐊 PANEL WEB: """ + nombre_activo.upper() + """</h1>
        <p>Estado: <strong style="color: """ + color_estado + """;">""" + estado_texto + """</strong></p>

        <div style="margin: 10px 0;">
            <a href="/revivir" class="btn-ai">🟢 Revivir / Activar</a>
            <a href="/matar" class="btn-dead">💀 Matar</a>
        </div>

        <div style="margin-top: 12px;">
            <h3>Seleccionar Motor de Síntesis:</h3>
            """ + motores_html + """
        </div>

        <div style="margin: 15px 0;">
            <div id="estadoMic" style="background: #16a34a; color: white; padding: 8px; border-radius: 12px; display: inline-block;">
                🟢 Iniciando Micrófono...
            </div>
            <p id="textoTranscrito" style="font-size: 11px; margin-top: 8px;">Habla cuando quieras.</p>
        </div>

        <div class="section">
            <h3>🎙️ Catálogo de Voces (""" + str(total_voces) + """ en """ + MOTOR_ACTUAL.upper() + """)</h3>
            <div class="voice-group">""" + voces_html + """</div>
        </div>
    </div>
</body>
</html>"""
            self.wfile.write(html.encode("utf-8"))
        except Exception:
            pass

    def log_message(self, format, *args):
        return

def iniciar_servidor():
    server = HTTPServer(("0.0.0.0", PUERTO_SERVIDOR), ServidorSistema)
    print(f"🚀 [SERVIDOR WEB ACTIVO] http://localhost:{PUERTO_SERVIDOR}")
    server.serve_forever()

if __name__ == "__main__":
    threading.Thread(target=iniciar_servidor, daemon=True).start()
    while True:
        try:
            time.sleep(1)
        except KeyboardInterrupt:
            sys.exit(0)
