"""
Adam Voice Call — real-time voice conversation with Adam.
Run: python call.py
Press SPACE to talk, release to send. Ctrl+C to end call.
"""

import os
import sys
import json
import tempfile
import threading
import subprocess
import asyncio
import time

import sounddevice as sd
import numpy as np
import scipy.io.wavfile as wav
import pygame
from openai import OpenAI

# ── Config ──────────────────────────────────────────────────────────────────
CREDENTIALS_FILE = os.path.join(os.path.dirname(__file__), "credentials.json")
SAMPLE_RATE = 16000
CHANNELS = 1
SILENCE_THRESHOLD = 0.01   # amplitude below this = silence
SILENCE_DURATION = 1.2     # seconds of silence before auto-cut
MAX_RECORD_SECS = 30       # safety cap

def load_creds():
    try:
        with open(CREDENTIALS_FILE, "r") as f:
            return json.load(f)
    except Exception:
        return {}

creds = load_creds()
GROQ_KEY = creds.get("groq", {}).get("api_key", "") or os.environ.get("GROQ_API_KEY", "")
OPENROUTER_KEY = creds.get("openrouter", {}).get("api_key", "") or os.environ.get("OPENROUTER_API_KEY", "")

groq_client  = OpenAI(api_key=GROQ_KEY, base_url="https://api.groq.com/openai/v1") if GROQ_KEY else None
chat_client  = OpenAI(api_key=OPENROUTER_KEY, base_url="https://openrouter.ai/api/v1")

# Load Adam's identity
WORKSPACE = "C:/Users/dani1/.openclaw/workspace/ollama"
OBSIDIAN  = "C:/Users/dani1/OneDrive/Documentos/Obsidian Vault"
MEMORY_FILE = f"{OBSIDIAN}/Adam/MEMORY.md"

def load_identity():
    parts = []
    for fname in ["IDENTITY.md", "USER.md", "SOUL.md"]:
        try:
            parts.append(open(f"{WORKSPACE}/{fname}").read())
        except Exception:
            pass
    try:
        parts.append(open(MEMORY_FILE).read())
    except Exception:
        pass
    return "\n\n".join(parts)

SYSTEM = f"""You are Adam, Daniel's personal AI assistant.
This is a VOICE CALL — respond naturally and conversationally, like talking to a person.
Keep replies SHORT (2-4 sentences max unless Daniel asks for detail).
Do not use bullet points, markdown, or lists — speak in natural sentences.
Address Daniel by name occasionally.

{load_identity()}"""

conversation = []  # running call history

# ── Audio ────────────────────────────────────────────────────────────────────

def record_until_silence() -> np.ndarray:
    """Record from mic, stop after SILENCE_DURATION seconds of silence or MAX_RECORD_SECS."""
    print("  🎤 Listening...", end="", flush=True)
    chunks = []
    silent_chunks = 0
    silence_limit = int(SILENCE_DURATION * SAMPLE_RATE / 512)

    def callback(indata, frames, time_info, status):
        chunks.append(indata.copy())

    with sd.InputStream(samplerate=SAMPLE_RATE, channels=CHANNELS,
                        dtype="float32", blocksize=512, callback=callback):
        start = time.time()
        last_sound = time.time()
        while True:
            time.sleep(0.05)
            if len(chunks) < 5:
                continue
            # check last chunk amplitude
            recent = chunks[-1]
            amplitude = np.abs(recent).mean()
            if amplitude > SILENCE_THRESHOLD:
                last_sound = time.time()
            if time.time() - last_sound > SILENCE_DURATION:
                break
            if time.time() - start > MAX_RECORD_SECS:
                break

    print(" done.")
    return np.concatenate(chunks, axis=0)

def save_wav(audio: np.ndarray) -> str:
    path = os.path.join(tempfile.gettempdir(), "adam_call_input.wav")
    wav.write(path, SAMPLE_RATE, (audio * 32767).astype(np.int16))
    return path

def transcribe(wav_path: str) -> str:
    if not groq_client:
        print("  ⚠ No Groq key — transcription unavailable.")
        return ""
    try:
        with open(wav_path, "rb") as f:
            result = groq_client.audio.transcriptions.create(
                model="whisper-large-v3",
                file=f,
                response_format="text",
            )
        return result.strip()
    except Exception as e:
        print(f"  ⚠ Transcription error: {e}")
        return ""

# ── TTS ───────────────────────────────────────────────────────────────────────

async def speak(text: str):
    """Convert text to speech and play it."""
    try:
        import edge_tts
        mp3_path = os.path.join(tempfile.gettempdir(), "adam_call_reply.mp3")
        # Clean markdown
        clean = text.replace("**","").replace("__","").replace("`","").replace("*","")
        communicate = edge_tts.Communicate(clean, voice="es-MX-JorgeNeural")
        await communicate.save(mp3_path)

        pygame.mixer.init()
        pygame.mixer.music.load(mp3_path)
        pygame.mixer.music.play()
        while pygame.mixer.music.get_busy():
            time.sleep(0.1)
    except Exception as e:
        print(f"  ⚠ TTS error: {e}")
        print(f"  Adam: {text}")

# ── Chat ──────────────────────────────────────────────────────────────────────

def chat(user_text: str) -> str:
    conversation.append({"role": "user", "content": user_text})
    messages = [{"role": "system", "content": SYSTEM}] + conversation[-12:]
    try:
        resp = chat_client.chat.completions.create(
            model="anthropic/claude-sonnet-4-5",
            messages=messages,
        )
        reply = (resp.choices[0].message.content or "").strip()
        conversation.append({"role": "assistant", "content": reply})
        return reply
    except Exception as e:
        return f"Error: {e}"

# ── Main loop ─────────────────────────────────────────────────────────────────

async def call_loop():
    print("\n" + "="*50)
    print("  📞 ADAM — LLAMADA EN VIVO")
    print("  Habla cuando veas 🎤. Silencio = enviar.")
    print("  Ctrl+C para colgar.")
    print("="*50 + "\n")

    # Greeting
    greeting = "Hola Daniel, Adam al habla. ¿En qué te ayudo?"
    print(f"  Adam: {greeting}")
    await speak(greeting)

    while True:
        try:
            # Record
            audio = record_until_silence()
            wav_path = save_wav(audio)

            # Check not just silence
            amplitude = np.abs(audio).mean()
            if amplitude < SILENCE_THRESHOLD * 0.5:
                print("  (silencio — esperando...)")
                continue

            # Transcribe
            text = transcribe(wav_path)
            if not text:
                continue
            print(f"  Tú: {text}")

            # Exit phrases
            if any(w in text.lower() for w in ["adiós", "adios", "hasta luego", "colgar", "bye"]):
                farewell = "Hasta luego Daniel, aquí estaré cuando me necesites."
                print(f"  Adam: {farewell}")
                await speak(farewell)
                break

            # Get response
            print("  Adam: ", end="", flush=True)
            reply = chat(text)
            print(reply)
            await speak(reply)

        except KeyboardInterrupt:
            print("\n\n  📵 Llamada terminada.")
            break

if __name__ == "__main__":
    asyncio.run(call_loop())
