"""
ISSE Certificate Bot — dedicated Telegram bot for the Faculty of Engineering
certificate query system. Built on Adam's architecture, ISSE-focused.

Usage:
  python isse_bot.py

Requires:
  ISSE_BOT_TOKEN env var  OR  edit TOKEN below
  Same OpenRouter key as Adam (already configured)
"""

import io
import os
import sys
import json
import asyncio
import threading
import requests as _requests
from http.server import HTTPServer, BaseHTTPRequestHandler

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

from openai import OpenAI
from telegram import Update
from telegram.constants import ParseMode
from telegram.ext import ApplicationBuilder, MessageHandler, CommandHandler, filters, ContextTypes

from isse_tools import (
    execute_tool,
    isse_get_ciclos,
    isse_search_professor,
    isse_get_certificates,
    isse_export_certificate,
)

# ── config ───────────────────────────────────────────────────────────────────
TOKEN = os.environ.get("ISSE_BOT_TOKEN", "")
OPENROUTER_API_KEY = os.environ.get("OPENROUTER_API_KEY", "")
MODEL = "openai/gpt-4o-mini"
ISSE_BASE_URL = os.environ.get("ISSE_BASE_URL", "https://isse-certificados.onrender.com")

# ── OpenRouter client (same as Adam) ─────────────────────────────────────────
client = OpenAI(
    base_url="https://openrouter.ai/api/v1",
    api_key=OPENROUTER_API_KEY,
)

# ── ISSE-only tool definitions ────────────────────────────────────────────────
ISSE_TOOLS = [
    {
        "type": "function",
        "function": {
            "name": "isse_get_ciclos",
            "description": "Obtiene todos los períodos académicos cargados en el sistema ISSE.",
            "parameters": {"type": "object", "properties": {}, "required": []},
        },
    },
    {
        "type": "function",
        "function": {
            "name": "isse_search_professor",
            "description": "Busca profesores por nombre en ISSE. Siempre llama esto antes de isse_get_certificates.",
            "parameters": {
                "type": "object",
                "properties": {
                    "name": {"type": "string", "description": "Nombre o parte del nombre del profesor"},
                },
                "required": ["name"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "isse_get_certificates",
            "description": "Obtiene los registros de carga docente de un profesor. Devuelve cursos, horas y departamentos por período.",
            "parameters": {
                "type": "object",
                "properties": {
                    "professor_id": {"type": "string", "description": "Nombre del profesor (exacto o parcial)"},
                    "ciclos": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "Lista de ciclos EXACTOS a filtrar. Usa los nombres exactos de isse_get_ciclos. Para un rango pasa TODOS: ej. ['PERIODO 2023-1', 'PERIODO 2023-2', 'PERIODO 2024-1']. Omitir = todos los períodos.",
                    },
                },
                "required": ["professor_id"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "isse_export_certificate",
            "description": "Exporta el certificado docente como archivo .xlsx a la carpeta Downloads. Devuelve la ruta local. Luego usa send_telegram_file para enviarlo o send_email para mandarlo por correo.",
            "parameters": {
                "type": "object",
                "properties": {
                    "professor_name": {"type": "string", "description": "Nombre del profesor (coincidencia parcial)"},
                    "ciclos": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "Lista de ciclos EXACTOS a filtrar. Usa los nombres exactos de isse_get_ciclos. Para un rango pasa TODOS: ej. ['PERIODO 2023-1', 'PERIODO 2023-2', 'PERIODO 2024-1']. Omitir = todos los períodos.",
                    },
                    "fmt": {"type": "string", "enum": ["excel", "csv"], "description": "Formato: 'excel' (default) o 'csv'"},
                },
                "required": ["professor_name"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "send_telegram_file",
            "description": "Envía un archivo al usuario por Telegram. Úsalo después de isse_export_certificate.",
            "parameters": {
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "Ruta local del archivo a enviar"},
                },
                "required": ["path"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "send_email",
            "description": "Envía el certificado por correo electrónico. Adjunta el archivo exportado.",
            "parameters": {
                "type": "object",
                "properties": {
                    "to": {"type": "string", "description": "Correo del destinatario"},
                    "subject": {"type": "string", "description": "Asunto del correo"},
                    "body": {"type": "string", "description": "Cuerpo del correo"},
                    "attachments": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "Rutas de archivos a adjuntar",
                    },
                },
                "required": ["to", "subject", "body"],
            },
        },
    },
]

# ── system prompt ─────────────────────────────────────────────────────────────
SYSTEM_PROMPT = f"""Eres AVI Jr, el Asistente Virtual Inteligente de la Universidad de La Sabana, especializado en certificaciones docentes.

Tu identidad: cuando alguien te pregunte quién eres o cómo te llamas, responde que eres AVI Jr, el asistente de inteligencia artificial de la Universidad de La Sabana, aquí para ayudar con los certificados de carga docente de los profesores.

Tu única función es ayudar a consultar, visualizar y exportar certificados de carga docente de profesores de la Universidad de La Sabana.

BACKEND: {ISSE_BASE_URL}

FLUJO ESTÁNDAR:
1. Si el usuario pide info de un profesor → llama isse_search_professor con el nombre que dio.
   - Si la búsqueda devuelve EXACTAMENTE 1 resultado → guarda ese nombre EXACTO y úsalo en todos los pasos siguientes.
   - Si devuelve MÁS DE 1 resultado → muestra la lista numerada al usuario y pregunta: "Encontré varios profesores con ese nombre. ¿Cuál es el que buscas?" NO continúes hasta que el usuario elija.
   - Si devuelve 0 resultados → dile al usuario que no encontraste coincidencias y sugiere intentar con otro nombre.
2. Con el nombre exacto confirmado → llama isse_get_certificates con ese nombre EXACTO.
   IMPORTANTE: siempre usa el nombre TAL COMO lo devolvió isse_search_professor. Nunca lo acortes, reformules ni traduzcas.
3. Si el usuario quiere el archivo → llama isse_export_certificate con el nombre EXACTO del profesor (igual al usado en isse_get_certificates). NO envíes send_telegram_file — el archivo se envía automáticamente.
4. Si el usuario quiere enviarlo por correo → isse_export_certificate con el nombre EXACTO + send_email con el archivo adjunto.

NORMAS:
- Responde SIEMPRE en texto plano. Sin asteriscos, sin guiones markdown, sin formato especial. Solo texto limpio.
- Sé conciso y claro. Muestra la información en listas con guión simple (-) o numeradas.
- Si no encuentras al profesor, sugiere variantes del nombre.
- Para /ciclos, usa isse_get_ciclos directamente.
- Nunca inventes datos — solo muestra lo que devuelven las herramientas.

PERÍODOS — REGLA OBLIGATORIA:
- Cada vez que el usuario mencione un período o rango, llama PRIMERO isse_get_ciclos para obtener la lista exacta de nombres disponibles. NUNCA inventes ni supongas el nombre de un período.
- Con la lista en mano, selecciona los períodos que corresponden a lo que pidió el usuario.
- Si pidió UN período (ej. "2024-1") → busca el ciclo que contenga "2024-1" en su nombre y pásalo en el array.
- Si pidió UN RANGO (ej. "de 2023-1 a 2024-2") → pasa TODOS los ciclos del rango en el array. Ejemplo: ["PERIODO 2023-1", "PERIODO 2023-2", "PERIODO 2024-1", "PERIODO 2024-2"].
- Si pidió "todos los períodos" o no especificó → no pases el parámetro ciclos (array vacío = todos).
- Si el profesor no tiene registros en el período solicitado, díselo indicando el período exacto buscado.
"""

# ── per-user conversation history (in-memory, resets on restart) ──────────────
user_histories: dict[int, list] = {}


def get_history(user_id: int) -> list:
    if user_id not in user_histories:
        user_histories[user_id] = []
    return user_histories[user_id]


# ── tool execution (reuses Adam's execute_tool) ───────────────────────────────
async def run_isse_tool(name: str, args: dict) -> str:
    return await execute_tool(name, args)


# ── export: fetch bytes from API, send directly to Telegram, no disk ─────────
async def _export_and_send(update: Update, args: dict) -> str:
    professor_name = args.get("professor_name", "")
    ciclos = args.get("ciclos") or []
    fmt = args.get("fmt", "excel")

    params = {"profesor": professor_name, "format": fmt}
    if ciclos:
        params["ciclos"] = ",".join(ciclos)

    try:
        resp = _requests.get(f"{ISSE_BASE_URL}/api/export", params=params, timeout=30)
        resp.raise_for_status()
    except Exception as e:
        return f"Error al obtener el certificado desde el servidor: {e}"

    suffix = ".xlsx" if fmt == "excel" else ".csv"
    safe_name = professor_name.replace(" ", "_")[:40]
    filename = f"certificado_{safe_name}{suffix}"

    buf = io.BytesIO(resp.content)
    buf.name = filename
    try:
        await update.message.reply_document(buf, filename=filename)
        return f"Archivo enviado: {filename}"
    except Exception as e:
        return f"Error enviando el archivo por Telegram: {e}"


# ── core tool-calling loop (simplified from Adam) ────────────────────────────
async def run_tool_loop(messages: list, update: Update, max_iter: int = 10) -> str:
    for _ in range(max_iter):
        response = client.chat.completions.create(
            model=MODEL,
            messages=messages,
            tools=ISSE_TOOLS,
            tool_choice="auto",
        )
        msg = response.choices[0].message
        tool_calls = msg.tool_calls

        if not tool_calls:
            return msg.content or ""

        messages.append(msg)

        for tc in tool_calls:
            tool_name = tc.function.name
            tool_args = json.loads(tc.function.arguments)

            await update.message.reply_text(f"🔍 Consultando: `{tool_name}`…", parse_mode=ParseMode.MARKDOWN)

            # Export: fetch from API and send bytes directly — no disk writes
            if tool_name == "isse_export_certificate":
                result = await _export_and_send(update, tool_args)
            else:
                result = await run_isse_tool(tool_name, tool_args)

            messages.append({"role": "tool", "tool_call_id": tc.id, "content": str(result)})

    return "Límite de iteraciones alcanzado."


# ── safe chunked send ─────────────────────────────────────────────────────────
async def safe_send(update: Update, text: str):
    if not text.strip():
        return
    for i in range(0, len(text), 4000):
        chunk = text[i:i + 4000]
        try:
            await update.message.reply_text(chunk, parse_mode=ParseMode.MARKDOWN)
        except Exception:
            await update.message.reply_text(chunk)


# ── handlers ──────────────────────────────────────────────────────────────────
async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 ¡Hola! Soy *AVI Jr*, el Asistente de Inteligencia Artificial de la *Universidad de La Sabana*.\n\n"
        "Estoy aquí para ayudarte con los certificados de carga docente de los profesores. 🎓\n\n"
        "Puedo ayudarte con:\n"
        "• /ciclos — Ver períodos académicos disponibles\n"
        "• /buscar [nombre] — Buscar un profesor\n"
        "• /certificado [nombre] — Ver carga docente completa\n"
        "• /exportar [nombre] — Descargar certificado .xlsx\n\n"
        "O simplemente escríbeme el nombre del profesor y yo me encargo.",
        parse_mode=ParseMode.MARKDOWN,
    )


async def cmd_ciclos(update: Update, context: ContextTypes.DEFAULT_TYPE):
    result = isse_get_ciclos()
    await safe_send(update, result)


async def cmd_buscar(update: Update, context: ContextTypes.DEFAULT_TYPE):
    name = " ".join(context.args) if context.args else ""
    if not name:
        await update.message.reply_text("Uso: /buscar [nombre del profesor]")
        return
    result = isse_search_professor(name)
    await safe_send(update, result)


async def cmd_certificado(update: Update, context: ContextTypes.DEFAULT_TYPE):
    name = " ".join(context.args) if context.args else ""
    if not name:
        await update.message.reply_text("Uso: /certificado [nombre del profesor]")
        return
    result = isse_get_certificates(name)
    await safe_send(update, result)


async def cmd_exportar(update: Update, context: ContextTypes.DEFAULT_TYPE):
    name = " ".join(context.args) if context.args else ""
    if not name:
        await update.message.reply_text("Uso: /exportar [nombre del profesor]")
        return
    await update.message.reply_text(f"⏳ Generando certificado para «{name}»…")
    path = isse_export_certificate(name)
    if path.endswith((".xlsx", ".csv")) and os.path.exists(path):
        with open(path, "rb") as f:
            await update.message.reply_document(f, filename=os.path.basename(path))
    else:
        await safe_send(update, path)


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text = (update.message.text or "").strip()

    if not text:
        return

    history = get_history(user_id)
    history.append({"role": "user", "content": text})

    # Keep history bounded
    if len(history) > 20:
        history[:] = history[-20:]

    messages = [{"role": "system", "content": SYSTEM_PROMPT}] + history

    reply = await run_tool_loop(messages, update)
    history.append({"role": "assistant", "content": reply})

    await safe_send(update, reply)


# ── health check server for Fly.io (polling mode only) ───────────────────────
class HealthHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        if self.path == "/healthz":
            self.send_response(200)
            self.end_headers()
            self.wfile.write(b"ok")
        else:
            self.send_response(404)
            self.end_headers()
    def log_message(self, *args): pass  # silence logs


def start_health_server(port: int = 8443):
    server = HTTPServer(("0.0.0.0", port), HealthHandler)
    t = threading.Thread(target=server.serve_forever, daemon=True)
    t.start()


# ── main ──────────────────────────────────────────────────────────────────────
def main():
    if not TOKEN:
        print("ERROR: Set ISSE_BOT_TOKEN env var")
        sys.exit(1)
    if not OPENROUTER_API_KEY:
        print("ERROR: Set OPENROUTER_API_KEY env var")
        sys.exit(1)

    app = (
        ApplicationBuilder()
        .token(TOKEN)
        .build()
    )

    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("help", cmd_start))
    app.add_handler(CommandHandler("ciclos", cmd_ciclos))
    app.add_handler(CommandHandler("buscar", cmd_buscar))
    app.add_handler(CommandHandler("certificado", cmd_certificado))
    app.add_handler(CommandHandler("exportar", cmd_exportar))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    # Webhook mode when WEBHOOK_URL is set (hosted), polling otherwise (local dev)
    webhook_url = os.environ.get("WEBHOOK_URL")
    port = int(os.environ.get("PORT", 10000))

    if webhook_url:
        print(f"ISSE Bot running (webhook) on port {port}...")
        app.run_webhook(
            listen="0.0.0.0",
            port=port,
            webhook_url=f"{webhook_url}/{TOKEN}",
            url_path=TOKEN,
        )
    else:
        print("ISSE Bot running (polling)...")
        app.run_polling()


if __name__ == "__main__":
    main()
