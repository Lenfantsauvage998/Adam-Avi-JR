<div align="center">

# 🤖 AVI Jr

**Asistente Virtual de Certificación Docente**  
**Facultad de Ingeniería · Universidad de La Sabana**

Consulta certificados de carga docente directamente desde Telegram.  
Sin formularios, sin esperas — solo escríbele el nombre del profesor.

[![Telegram](https://img.shields.io/badge/Telegram-Bot-26A5E4?style=for-the-badge&logo=telegram&logoColor=white)](https://t.me/)

![Python](https://img.shields.io/badge/Python-3776AB?style=flat-square&logo=python&logoColor=white)
![OpenAI](https://img.shields.io/badge/GPT--4o--mini-412991?style=flat-square&logo=openai&logoColor=white)
![Render](https://img.shields.io/badge/Render-46E3B7?style=flat-square&logo=render&logoColor=black)

</div>

---

## ¿Qué es AVI Jr?

AVI Jr es un bot de Telegram especializado en la consulta de certificados de carga docente. Se conecta al sistema ISSE en tiempo real y puede buscar profesores, mostrar su historial de cursos y exportar certificados directamente en el chat — todo mediante lenguaje natural en español.

---

## 🚀 Primeros pasos

```
1. Abre Telegram
2. Busca el bot por su usuario
3. Envía  /start
4. ¡Listo! Ya puedes hacer preguntas en lenguaje natural
```

---

## 💬 Cómo usarlo

AVI Jr entiende lenguaje natural. No necesitas comandos exactos:

```
"Busca el certificado de Mojica"
"¿Qué materias dictó García en 2024-1?"
"Exporta el certificado de López para 2023-2 y 2024-1"
"¿Qué períodos están disponibles?"
"Muéstrame la carga de Rodríguez en los últimos dos semestres"
```

---

## 📌 Comandos disponibles

| Comando | Descripción |
|---------|-------------|
| `/start` | Bienvenida e instrucciones |
| `/ciclos` | Lista todos los períodos académicos disponibles |
| `/buscar [nombre]` | Busca un profesor por nombre |
| `/certificado [nombre]` | Muestra la carga docente completa |
| `/exportar [nombre]` | Descarga el certificado como archivo `.xlsx` |

---

## 🔄 Flujo típico

```
Tú:     "Busca a Mojica"

AVI Jr: Encontré 2 profesores:
        1. MOJICA MACIAS JUAN PABLO
        2. MOJICA OSSA NICOLAS
        ¿Cuál buscas?

Tú:     "Juan Pablo"

AVI Jr: Certificados de MOJICA MACIAS JUAN PABLO — 6 registros en 3 períodos:

        PERIODO 2024-1  (2 cursos · 128 hrs):
          - Cálculo Diferencial — 64 hrs — Matemáticas
          - Álgebra Lineal — 64 hrs — Matemáticas

        PERIODO 2023-2  (2 cursos · 128 hrs):
          ...

Tú:     "Exporta el certificado"

AVI Jr: [envía archivo .xlsx directamente al chat]
```

---

## 🏗️ Arquitectura

```
Usuario (Telegram)
       │  mensaje
       ▼
  Render.com — isse-bot-1
  ┌──────────────────────────────────────────┐
  │  python-telegram-bot  (polling)          │
  │           │                              │
  │     OpenRouter API                       │
  │     GPT-4o-mini  →  decide qué tool usar │
  │           │                              │
  │     isse_tools.py  →  consulta al backend│
  └──────────────────────────────────────────┘
               │
               ▼
  proyecto-indes-challenge.onrender.com
               │
               ▼
        Neon PostgreSQL
```

---

## ⚙️ Stack técnico

| Componente | Tecnología |
|------------|------------|
| Bot framework | python-telegram-bot 21.x |
| LLM | GPT-4o-mini vía OpenRouter |
| Backend ISSE | FastAPI en Render |
| Base de datos | Neon PostgreSQL |
| Deploy | Render.com (Docker, free tier) |

---

## 📁 Archivos principales

```
pc-agent/
├── isse_bot.py          # Bot principal — handlers, tool loop, LLM
├── isse_tools.py        # Herramientas: buscar, certificados, exportar
├── Dockerfile.bot       # Imagen Docker del bot
├── requirements-bot.txt # Dependencias Python
└── render.yaml          # Config de despliegue
```

---

## 🔑 Variables de entorno (Render)

| Variable | Descripción |
|----------|-------------|
| `ISSE_BOT_TOKEN` | Token del bot de Telegram (BotFather) |
| `OPENROUTER_API_KEY` | API key de OpenRouter |
| `ISSE_BASE_URL` | URL del backend ISSE |
| `PORT` | Puerto del health server (`10000`) |

---

## 🧠 Cómo funciona el LLM

AVI Jr usa un loop de herramientas (tool-calling loop) con GPT-4o-mini:

```
mensaje usuario
      │
      ▼
   LLM decide
      │
   ┌──┴─────────────────────────────────────┐
   │  isse_search_professor(name)           │  ← busca profesor
   │  isse_get_certificates(id, ciclos)     │  ← obtiene carga
   │  isse_export_certificate(name, ciclos) │  ← exporta .xlsx
   │  isse_get_ciclos()                     │  ← lista períodos
   └──┬─────────────────────────────────────┘
      │  resultado de la herramienta
      ▼
   LLM redacta respuesta en español
      │
      ▼
   respuesta al usuario
```

El bot mantiene historial de conversación por usuario (últimos 20 mensajes) para entender contexto como "el mismo profesor" o "ese período".

---

<div align="center">
  <sub>Facultad de Ingeniería · Universidad de La Sabana · AVI Jr</sub>
</div>
