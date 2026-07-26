import asyncio
import ctypes
import io
import os
import sys
import threading
import time
import urllib.parse
from http.server import BaseHTTPRequestHandler, HTTPServer

import edge_tts
from gtts import gTTS
import ollama
import requests

# --- CONFIGURACIÓN DE CLAVES ---
ELEVENLABS_API_KEY = "sk_a875d0a08d684229b591825019dd115bd2f8e21c1060de52"

# Puerto dinámico compatible con Render o local (5000)
PUERTO_SERVIDOR = int(os.environ.get("PORT", 5000))

# Variables globales de estado
MOTOR_ACTUAL = "edge"
INDICE_VOZ_ACTUAL = 0
COCODRILO_VIVO = True  
GENERO_ACTUAL = "cocodrily"
EFECTO_VOZ_ACTUAL = "normal"

# 🔒 Bandera para bloquear el micrófono mientras habla
BOT_HABLANDO = False

# Último buffer de audio generado listo para enviarse al HTML
ULTIMO_AUDIO_BYTES = None

# 1. Escaneo completo de TODAS las voces SAPI5 (Windows Local)
voces_sapi_instaladas = []
try:
    import win32com.client
    ctypes.windll.ole32.CoInitialize(None)
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    for voice in speaker.GetVoices():
        voces_sapi_instaladas.append(voice.GetDescription())
    ctypes.windll.ole32.CoUninitialize()
except Exception:
    voces_sapi_instaladas = ["Microsoft Helena", "Microsoft Pablo"]

# 2. Carga completa de TODAS las voces en español de Edge-TTS
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
        
        lista_completa = []
        for v in voces_es:
            item = {"nombre": v["ShortName"], "rate": "+0%", "pitch": "+0Hz"}
            lista_completa.append(item)
            
        if not lista_completa:
            lista_completa = [{"nombre": "es-ES-AlvaroNeural", "rate": "+0%", "pitch": "+0Hz"}, {"nombre": "es-ES-ElviraNeural", "rate": "+0%", "pitch": "+0Hz"}]

        voces_edge_cache["mujeres"] = lista_completa
        voces_edge_cache["hombres"] = []
    except Exception as e:
        print(f"⚠️ [EDGE WARNING CARGA]: {e}")
        voces_edge_cache["mujeres"] = [{"nombre": "es-ES-ElviraNeural", "rate": "+0%", "pitch": "+0Hz"}]

cargar_todas_voces_edge()

# 3. Carga completa de TODAS las voces de ElevenLabs desde la API
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
            lista_voces = data.get("voices", [])
            lista_completa = []
            for v in lista_voces:
                nombre = v.get("name", "Desconocida")
                voice_id = v.get("voice_id", "")
                lista_completa.append({"nombre": nombre, "id": voice_id})
            voces_elevenlabs_cache["mujeres"] = lista_completa
    except Exception:
        pass

cargar_todas_voces_elevenlabs()

voces_por_servicio = {
    "edge": voces_edge_cache,
    "elevenlabs": voces_elevenlabs_cache,
    "local": {
        "mujeres": voces_sapi_instaladas,
        "hombres": []
    },
    "gtts": {
        "mujeres": [
            "es-ES (España)", "es-MX (México)", "es-US (Estados Unidos)", 
            "es-AR (Argentina)", "es-CO (Colombia)", "es-CL (Chile)", 
            "es-PE (Perú)", "es-VE (Venezuela)"
        ],
        "hombres": []
    }
}

modo_actual = 0
FRASES_REALISTAS = [
    "Hmm, déjame pensar un segundo en eso.",
    "¡Vaya, eso sí que no me lo esperaba!",
    "La verdad es que tienes toda la razón.",
    "¿De verdad? Cuéntame más sobre eso, me interesa.",
    "Estoy un poco ocupado ahora mismo mirando por la ventana, pero te escucho.",
    "¡Jaja, qué ocurrencia la tuya!",
    "Bueno, las cosas como son, hay que admitirlo.",
    "Eso suena genial, ¿cuándo lo hacemos?",
    "A veces me pongo a pensar en mis cosas y se me va el santo al cielo.",
    "¡Claro que sí! Cuenta conmigo para lo que sea."
]
contador_frase = 0

def obtener_nombre_actual():
    return "Cocodrila" if GENERO_ACTUAL == "cocodrila" else "Cocodrily"

def actualizar_genero_segun_voz():
    global GENERO_ACTUAL
    GENERO_ACTUAL = "cocodrily"

def preparar_audio_mp3(buffer_bytes):
    global ULTIMO_AUDIO_BYTES, BOT_HABLANDO
    try:
        if not buffer_bytes or not isinstance(buffer_bytes, io.BytesIO):
            BOT_HABLANDO = False
            return
        buffer_bytes.seek(0)
        ULTIMO_AUDIO_BYTES = buffer_bytes.read()
    except Exception as e:
        print(f"⚠️ [AUDIO MP3 ERROR]: {e}")
    finally:
        BOT_HABLANDO = False

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
            preparar_audio_mp3(io.BytesIO(response.content))
        else:
            global BOT_HABLANDO
            BOT_HABLANDO = False
    except Exception:
        BOT_HABLANDO = False

def _hablar_edge_bytes(texto):
    global BOT_HABLANDO
    elementos = voces_por_servicio["edge"]["mujeres"]
    if not elementos:
        voz_nombre = "es-ES-ElviraNeural"
    else:
        item = elementos[INDICE_VOZ_ACTUAL % len(elementos)]
        voz_nombre = item["nombre"]

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
            preparar_audio_mp3(io.BytesIO(ab))
        else:
            BOT_HABLANDO = False
    except Exception as e:
        print(f"⚠️ [EDGE TTS ERROR]: {e}")
        BOT_HABLANDO = False

def generar_audio_web(texto):
    global BOT_HABLANDO
    BOT_HABLANDO = True
    if not texto or not texto.strip():
        BOT_HABLANDO = False
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
            preparar_audio_mp3(buffer_ram)
        else:
            _hablar_edge_bytes(texto)
    except Exception:
        BOT_HABLANDO = False

def procesar_inteligencia(prompt):
    global modo_actual, contador_frase, COCODRILO_VIVO
    texto_lower = prompt.lower()

    if "muérete" in texto_lower or "muere" in texto_lower:
        COCODRILO_VIVO = False
        return "..."

    if "pedo" in texto_lower or "gas" in texto_lower:
        return "¡Prrrt! Ups, ¡fue el cocodrilo!"

    if not COCODRILO_VIVO:
        return ""

    rol_genero = "Eres Cocodrila, alegre y divertida." if GENERO_ACTUAL == "cocodrila" else "Eres Cocodrily, amigable y curioso."
    
    if modo_actual in [0, 1]:
        modelo_nom = "qwen2:1.5b" if modo_actual == 0 else "llama3.1"
        try:
            client = ollama.Client()
            res = client.chat(
                model=modelo_nom, 
                messages=[{"role": "user", "content": f"{prompt}\n({rol_genero} Responde súper corto y directo en español.)"}], 
                options={"num_predict": 30, "temperature": 0.7}
            )
            return res["message"]["content"]
        except Exception:
            return FRASES_REALISTAS[0]
    else:
        res = FRASES_REALISTAS[contador_frase % len(FRASES_REALISTAS)]
        contador_frase += 1
        return res

class ServidorSistema(BaseHTTPRequestHandler):
    def do_GET(self):
        global modo_actual, MOTOR_ACTUAL, INDICE_VOZ_ACTUAL, EFECTO_VOZ_ACTUAL, COCODRILO_VIVO, ULTIMO_AUDIO_BYTES
        try:
            parsed_path = urllib.parse.urlparse(self.path)
            path = parsed_path.path

            if path == "/audio-actual":
                self.send_response(200)
                self.send_header("Content-type", "audio/mpeg")
                if ULTIMO_AUDIO_BYTES:
                    self.send_header("Content-Length", str(len(ULTIMO_AUDIO_BYTES)))
                    self.end_headers()
                    self.wfile.write(ULTIMO_AUDIO_BYTES)
                else:
                    self.end_headers()
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

            elif path.startswith("/efecto/"):
                EFECTO_VOZ_ACTUAL = path.split("/")[-1]
                self.send_response(303)
                self.send_header("Location", "/")
                self.end_headers()
                return

            elif path.startswith("/motor/"):
                MOTOR_ACTUAL = path.split("/")[-1]
                INDICE_VOZ_ACTUAL = 0
                actualizar_genero_segun_voz()
                self.send_response(303)
                self.send_header("Location", "/")
                self.end_headers()
                return

            elif path.startswith("/modo/"):
                try:
                    modo_actual = int(path.split("/")[-1])
                except Exception:
                    pass
                self.send_response(303)
                self.send_header("Location", "/")
                self.end_headers()
                return

            elif path.startswith("/voz/"):
                try:
                    INDICE_VOZ_ACTUAL = int(path.split("/")[-1])
                    actualizar_genero_segun_voz()
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
            estado_texto = "🟢 VIVO Y ACTIVO" if COCODRILO_VIVO else "💀 MODO MUERTO"
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
        .btn-fx { background: #0284c7; font-size: 10px; margin: 2px; padding: 6px 10px; text-decoration: none; color: white; border-radius: 4px; display: inline-block; }
        .btn-ai { background: #2563eb; margin: 3px; padding: 8px 14px; text-decoration: none; color: white; border-radius: 4px; display: inline-block; font-weight: bold; }
        .btn-dead { background: #dc2626; margin: 3px; padding: 8px 14px; text-decoration: none; color: white; border-radius: 4px; display: inline-block; font-weight: bold; }
        .btn-activo { background: #16a34a !important; font-weight: bold; border: 1px solid #86efac; }
        .section { margin-top: 15px; border-top: 1px solid #334155; padding-top: 10px; text-align: left; }
        .voice-group { background: #111827; padding: 10px; border-radius: 6px; margin-bottom: 8px; max-height: 250px; overflow-y: auto; }
    </style>
    <script>
        let recognition = null;

        function reproducirVozHTML() {
            const audioElem = document.getElementById("reproductorAudio");
            const srcActual = "/audio-actual?t=" + new Date().getTime();
            if (audioElem.src !== srcActual) {
                audioElem.src = srcActual;
                audioElem.play().catch(e => console.log("Esperando interacción para audio"));
            }
        }

        setInterval(async () => {
            try {
                let res = await fetch("/audio-actual");
                if (res.ok) {
                    let blob = await res.blob();
                    if (blob.size > 100) {
                        reproducirVozHTML();
                    }
                }
            } catch(e) {}
        }, 1500);

        function iniciarEscuchaContinua() {
            const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
            if (!SpeechRecognition) {
                document.getElementById("estadoMic").innerText = "⚠️ Navegador sin soporte de micrófono.";
                return;
            }
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
            recognition.onerror = function(event) {
                if (event.error === 'not-allowed') {
                    document.getElementById("estadoMic").innerText = "🔴 Micrófono Bloqueado (Permisos)";
                    document.getElementById("estadoMic").style.background = "#dc2626";
                }
            };
            recognition.onend = function() { 
                setTimeout(() => { try { recognition.start(); } catch(e) {} }, 500); 
            };
            try { recognition.start(); } catch(e) {}
        }

        window.onload = function() { iniciarEscuchaContinua(); };
    </script>
</head>
<body>
    <div class="container">
        <h1>🐊 PANEL WEB: """ + nombre_activo.upper() + """</h1>
        <p>Estado: <strong style="color: """ + color_estado + """;">""" + estado_texto + """</strong></p>

        <audio id="reproductorAudio" autoplay></audio>

        <div style="margin: 10px 0;">
            <a href="/revivir" class="btn-ai">🟢 Revivir / Activar</a>
            <a href="/matar" class="btn-dead">💀 Matar</a>
        </div>

        <div>
            <h3>Modo de Respuesta:</h3>
            <a href="/modo/0" class="btn-ai """ + ("btn-activo" if modo_actual == 0 else "") + """">⚡ Qwen 2 (1.5B)</a>
            <a href="/modo/1" class="btn-ai """ + ("btn-activo" if modo_actual == 1 else "") + """">🧠 Llama 3.1</a>
            <a href="/modo/2" class="btn-ai """ + ("btn-activo" if modo_actual == 2 else "") + """">💬 Modo Realista</a>
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
    actualizar_genero_segun_voz()
    threading.Thread(target=iniciar_servidor, daemon=True).start()
    while Time.sleep(1):
        pass
