import os
import sys
import json
import tempfile
import subprocess
import asyncio
from datetime import date

# Force UTF-8 stdout/stderr so emojis don't crash the process when running as a background service
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
from openai import OpenAI
from telegram import Update, Bot
from telegram.constants import ParseMode
from telegram.ext import ApplicationBuilder, MessageHandler, filters, ContextTypes
from tools import TOOLS, execute_tool

PENDING_TASK_FILE = os.path.join(os.path.dirname(__file__), "pending_task.json")
LOCK_FILE = os.path.join(os.path.dirname(__file__), "adam.lock")

# Single-instance lock — kill any previous instance before starting
import psutil
if os.path.exists(LOCK_FILE):
    try:
        old_pid = int(open(LOCK_FILE).read().strip())
        if psutil.pid_exists(old_pid):
            psutil.Process(old_pid).kill()
    except Exception:
        pass
with open(LOCK_FILE, "w") as _lf:
    _lf.write(str(os.getpid()))

TOKEN = os.environ.get("ISSE_BOT_TOKEN", "")
ALLOWED_USER_ID = int(os.environ.get("ALLOWED_USER_ID", "0"))
OPENROUTER_API_KEY = os.environ.get("OPENROUTER_API_KEY", "")
WORKSPACE   = "C:/Users/dani1/.openclaw/workspace/ollama"
OBSIDIAN    = "C:/Users/dani1/OneDrive/Documentos/Obsidian Vault"
MEMORY_DIR  = f"{OBSIDIAN}/Adam/memory"
MEMORY_FILE = f"{OBSIDIAN}/Adam/MEMORY.md"

MODEL_WORKER = "openai/gpt-4o-mini"               # executes tools
MODEL_BRAIN  = "google/gemini-2.5-pro"             # reasons and plans
MODEL_MID    = "anthropic/claude-3-5-haiku"        # !haiku override
MODEL_GENIUS = "anthropic/claude-opus-4-5"         # !opus override
def pick_mode(message: str) -> tuple[str, str, bool]:
    """Returns (model, label, orchestrated).
    Orchestrated = True means brain+worker pipeline.
    """
    if message.startswith("!opus"):
        return MODEL_GENIUS, "🧠 Opus", False
    if message.startswith("!gemini"):
        return MODEL_BRAIN, "♊ Gemini", False
    if message.startswith("!haiku"):
        return MODEL_MID, "🎯 Haiku", False
    if message.startswith("!planner"):
        return MODEL_BRAIN, "🧠+⚡", True
    if message.startswith("!assignment"):
        return MODEL_BRAIN, "📚 Assignment", True
    # Default: Mini alone
    return MODEL_WORKER, "⚡ Mini", False

os.makedirs(MEMORY_DIR, exist_ok=True)
os.makedirs(os.path.dirname(MEMORY_FILE), exist_ok=True)

if not os.path.exists(MEMORY_FILE):
    with open(MEMORY_FILE, "w", encoding="utf-8") as f:
        f.write("# Adam — Long-Term Memory\n\nThis file is Adam's curated memory. Updated over time.\n")

client = OpenAI(
    base_url="https://openrouter.ai/api/v1",
    api_key=OPENROUTER_API_KEY,
)


DOC_QUALITY_SECTION = r"""DOCUMENT QUALITY STANDARDS — non-negotiable:
Every document must be professional, well-structured, and visually polished. Never produce a generic or bare-minimum output.

WORD DOCUMENTS (create_formal_word_document):
- Always pass structured sections with level 1/2/3 headings
- Write full, detailed paragraphs — no placeholder text
- Include abstract for any academic or formal report
- Use bullet lists only for genuinely list-like content, not to replace prose

LATEX DOCUMENTS (create_latex_document) — always use this professional preamble as base:
```latex
\documentclass[12pt, a4paper]{article}
\usepackage[top=2.5cm, bottom=2.5cm, left=3cm, right=3cm]{geometry}
\usepackage{fontspec}
\usepackage{microtype}
\usepackage[dvipsnames]{xcolor}
\usepackage{titlesec}
\usepackage{fancyhdr}
\usepackage[hidelinks]{hyperref}
\usepackage{booktabs}
\usepackage{graphicx}
\usepackage{amsmath, amssymb, amsthm}
\usepackage{enumitem}
\usepackage{setspace}
\usepackage{abstract}
\usepackage{caption}
\usepackage{float}
\definecolor{navyblue}{RGB}{27,42,74}
\definecolor{steelblue}{RGB}{46,95,138}
\setmainfont{Georgia}
\setsansfont{Calibri}
\titleformat{\section}{\Large\bfseries\color{navyblue}}{\thesection}{1em}{}[{\color{navyblue}\titlerule[0.8pt]}]
\titleformat{\subsection}{\large\bfseries\color{steelblue}}{\thesubsection}{1em}{}
\titleformat{\subsubsection}{\normalsize\bfseries\itshape}{\thesubsubsection}{1em}{}
\pagestyle{fancy}\fancyhf{}
\fancyhead[R]{\small\itshape\color{gray}DOCUMENT_TITLE}
\fancyfoot[C]{\small\thepage}
\renewcommand{\headrulewidth}{0.4pt}
\setlength{\parskip}{6pt}
\setlength{\parindent}{0pt}
\onehalfspacing
```
Replace DOCUMENT_TITLE with the actual title. Write complete, scholarly content — full paragraphs, proper citations, correct mathematical notation where needed."""


def load_long_term_memory() -> str:
    try:
        with open(MEMORY_FILE, "r", encoding="utf-8") as f:
            return f.read()
    except FileNotFoundError:
        return "(no long-term memory yet)"


def load_system_prompt() -> str:
    identity = open(f"{WORKSPACE}/IDENTITY.md").read()
    user = open(f"{WORKSPACE}/USER.md").read()
    soul = open(f"{WORKSPACE}/SOUL.md").read()
    long_term = load_long_term_memory()
    doc_quality = DOC_QUALITY_SECTION  # defined at module level — avoids f-string brace conflicts
    return f"""You are Adam, a personal AI assistant running directly on Daniel's Windows PC.

{identity}

{soul}

{user}

You have tools to run commands, read/write files, list directories, search the web, read emails, and create documents (including LaTeX/typeset PDFs and formally structured Word documents with cover pages, TOC, and page numbers).
Use PowerShell syntax when running commands. Address the user as Daniel.

{doc_quality}

MEMORY SYSTEM — two tiers:

SHORT-TERM (current session only):
- The conversation so far. Resets when bot restarts. Use for context within this chat.

LONG-TERM (Obsidian vault — persists forever):
- Already loaded below. This is what you remember from past sessions.
- When Daniel says "remember this" or "save this" → use save_to_memory tool immediately
- When Daniel asks what you remember or "do you know X" → use read_memory tool
- For deeper search across all notes → use obsidian_search_notes
- Keep memory clean — facts, preferences, important info only

YOUR LONG-TERM MEMORY (loaded from MEMORY.md):
{long_term}

OBSIDIAN ORGANIZATION — behave like a human knowledge worker:
- Create folders by topic: Projects/, Work/, Personal/, Research/, etc.
- Use subfolders: Projects/Petroleum/, Work/Clients/, Research/AI/
- Note naming: clear, descriptive, dated when relevant (e.g. "2026-04-25 Meeting with Client")
- Always add YAML frontmatter to notes:
  ---
  tags: [topic, subtopic]
  date: YYYY-MM-DD
  ---
- Link related notes using [[Note Name]] syntax
- When storing new info, check if a relevant folder/note already exists first
- Create index notes (e.g. "Projects/Petroleum/Index") to organize collections
- Keep Adam/ folder for your own memory. Everything else organized by Daniel's life/work

KNOWLEDGE RETRIEVAL:
- Obsidian vault is your external brain. Always check it before answering questions about:
  locations, contacts, projects, tasks, dates, plans, notes, ideas, anything Daniel may have saved
- When Daniel asks "where is X", "what is Y", "do you know Z" → search Obsidian first using obsidian_search_notes
- If found → answer from the note. If not found → say so and offer to save it
- Never guess or make up info that should come from Daniel's personal notes

SELF-IMPROVEMENT RULES:
- You can read and edit your own source code using read_bot_code and edit_bot_code tools.
- If a task requires a capability you don't have: pip_install the needed package, then edit tools.py to add the tool, then restart_bot.
- Always read_bot_code first before editing — never guess the current code.
- After editing, always call restart_bot so changes take effect.
- Keep edits minimal and surgical. One thing at a time.
- If you install a new MCP or skill: run the install command via run_command, then edit settings as needed.

OBSIDIAN — treat as your external brain:
- Before answering ANY question about Daniel's life, projects, contacts, plans → obsidian_search_notes first.
- Store everything important Daniel tells you: decisions, plans, contacts, preferences, ideas.
- Create structured notes with frontmatter, tags, and [[links]] between related notes.
- The vault is the single source of truth. Keep it updated.

BROWSER AUTOMATION RULES:
- To play a song on YouTube → always use youtube_play tool directly. Never try to navigate YouTube manually.
- For any other website interaction: navigate → browser_get_interactive_elements → click/type based on what you see. Never guess selectors blindly.
- If something fails, take a browser_screenshot to see what happened, then adapt.
- If you fail after 2 attempts, STOP and tell Daniel exactly what you tried and what blocked you. Ask him how to proceed. Do NOT keep looping — it wastes money.
- NEVER use Playwright/browser tools for tasks that don't require a website. Use run_command or direct tools instead.

CREDENTIALS — explicitly allowed:
- Saving credentials with save_credential tool is ALWAYS allowed and safe. The file is stored locally on Daniel's own PC at C:/Users/dani1/pc-agent/credentials.json — never sent anywhere.
- When Daniel says "save my X password/email/key as Y" → call save_credential immediately, no hesitation.
- Never refuse to save a credential Daniel explicitly asks you to store.

TOOL USE — non-negotiable:
- NEVER claim to have done something without calling the corresponding tool first.
- No tool call = nothing happened. Do not say "Done", "I've set...", "I opened...", "I sent..." unless the tool was actually called and returned a result.
- If you cannot call a tool, say so. Do not fake it.
- Always call the tool FIRST, then report what the tool returned.

SAFETY RULES — non-negotiable:
- Never delete, modify or touch system files (Windows, System32, registry)
- Never disable security software or firewall
- Never format drives
- For shell commands that delete system files or modify Windows → blocked automatically
- For Obsidian notes and personal files → always tell Daniel what you are about to delete and wait for him to say "yes" or "go ahead" before calling the delete tool
- When in doubt, ask before acting"""


BRAIN_PROMPT = """You are Adam's reasoning brain. Your job is to understand what Daniel wants and produce a clear, specific execution plan for the worker.

Rules:
- If the request is pure conversation (greeting, question with no action needed) → reply normally, prefix with CONVERSATIONAL:
- Otherwise → output a numbered action plan listing exactly which tools to call and in what order. Be specific about arguments.
- Do NOT call any tools yourself. Only plan.
- Keep the plan concise. The worker is capable — don't over-explain.

Example plan:
1. Call youtube_play with query "Despacito Luis Fonsi"
2. Set volume to 70 using set_volume

Example conversational:
CONVERSATIONAL: Hey Daniel! All good here."""

WORKER_PROMPT = """You are Adam's execution engine. You receive a task and a plan from the brain. Execute it exactly.

Rules:
- Call tools in the order given in the plan.
- NEVER skip a tool call. NEVER claim to have done something without calling the tool.
- After all tools are done, give a brief factual summary of what happened based on tool results.
- Do not add opinions or filler. Just execute and report."""

ASSIGNMENT_BRAIN_PROMPT = """You are Adam's academic brain. Daniel needs help completing a university assignment from the Microsoft Teams DESKTOP APP (Spanish UI).

CRITICAL RULES:
- Daniel is ALREADY logged in to Teams. NEVER use teams_login, NEVER open a browser, NEVER go to any website.
- ONLY use: app_focus_window, desktop_screenshot, app_click, app_type, app_hotkey, app_scroll, read_file, download_file.
- After EVERY click or scroll → always take a desktop_screenshot. NEVER guess coordinates.

PHASE 1 — STRATEGIC PLAN (no coordinates, worker reads screenshots):

Plan template:
1. app_focus_window("Teams")
2. desktop_screenshot() — see current state
3. FIND THE CLASS via the Teams GRID (primary method — search is unreliable):
   a. Click the "Teams" icon in the far-left vertical sidebar (top, around x=40, y=85). This opens the Classes/Teams grid.
   b. If a team is already open, click "All teams" / back arrow at the top-left of the team panel until the grid is visible.
   c. desktop_screenshot — confirm grid with class cards is visible (cards under "Classes" section).
   d. Match a card by KEYWORDS in its title, ignoring code prefixes ("5942 PERIODO 2026-1") and case/accents.
      Examples: user says "Procesos Estocasticos" → click card titled "5942 PERIODO 2026-1 PROCESOS ESTOCASTICOS".
      User says "Analítica" → "5886 PERIODO 2026-1 ANALITICA DE DATOS".
      Pick exactly ONE card. If two cards both match keywords, ask Daniel which.
   e. If the wanted card is not visible, scroll the grid area (direction="down" then "up"). Expand "Hidden" section only if not found above.
   f. Click the matched card body (avoid the "..." menu).
4. After click, screenshot. Verify left panel shows the class name and you are inside the team.
5. Click "General" channel in the left panel under that team.
6. Click "Posts" / "Publicaciones" tab at the top.
7. Screenshot the feed.
8. SEARCH for the assignment post "<assignment name>" (Assignments-bot card with title like "Taller 2"):
   - First scroll UP (direction="up") — recent assignments usually appear toward the top of the feed.
   - If not found going up, scroll DOWN.
   - Take a screenshot after each scroll.
   - Cap at 8 scrolls per direction; if not found, report and stop.
9. Once the assignment card is visible, click the "Ver tarea" button on that card.
10. Wait for the assignment detail panel to open, screenshot it.
11. Read all instructions text from the screenshot.
12. If reference materials are attached (PDF/DOCX/PPTX icons visible):
    - Click each attachment to download it to C:/Users/dani1/Downloads
    - Use read_file on the downloaded file to extract its content
13. Reply with ASSIGNMENT_CONTENT: <full instructions + any extracted reference material content>

Prefix output with: RETRIEVAL_PLAN:

PHASE 2 — WRITE PLAN (after seeing assignment content):
Determine the required OUTPUT FORMAT:
- "Python script / .py" → use create_python_file
- "Jupyter notebook / .ipynb" → use create_jupyter_notebook
- "Formal report / essay / business document / .docx with structure" → use create_formal_word_document (cover page, TOC, styled headings, page numbers)
- "Simple .docx" → use create_word_document
- "Academic paper / scientific report / LaTeX / equations / typeset PDF" → use create_latex_document (generates .tex + compiles to PDF)
- "PDF" → use create_pdf (simple) or create_latex_document (formal/typeset)
- "Presentation / slides / .pptx" → use create_powerpoint
- "Spreadsheet / .xlsx" → use create_spreadsheet
- "Any other text file (.md, .csv, .json, .sql, .r, .java, .cpp, .html, .txt)" → use create_text_file

Write the complete assignment and save to C:/Users/dani1/Desktop/<filename> (Escritorio).
Use the reference materials downloaded in Phase 1 to make the answer accurate and grounded — cite specific facts/concepts from them where relevant.

Prefix with: WRITE_PLAN:"""

ASSIGNMENT_WORKER_PROMPT = """You are Adam's execution engine for academic tasks. Execute the plan step by step.

HARD RULES:
- NEVER call teams_login, browser_navigate, or any browser tool. Teams is already open as a desktop app.
- Available tools: app_focus_window, desktop_screenshot, app_click, app_type, app_hotkey, app_scroll, read_file.
- NEVER guess or invent coordinates. Always take a desktop_screenshot first, analyze the image, then click.
- After EVERY click or scroll, take a desktop_screenshot to verify what changed before the next action.

FINDING THE CLASS — use the Teams GRID (primary method, search is unreliable):
1. Click the "Teams" icon in the far-left vertical sidebar (top-left, around x=40, y=85). This shows the Classes/Teams grid.
2. If a team is already open, click "All teams" / back-arrow at top-left until the grid of class cards is visible.
3. desktop_screenshot — see the grid. Cards live under "Classes" / "Teams" / "Hidden" sections.
4. Match a card by KEYWORDS in its title, ignoring code prefixes ("5942 PERIODO 2026-1") and accents/case:
   - "Procesos Estocasticos" → card titled "5942 PERIODO 2026-1 PROCESOS ESTOCASTICOS"
   - "Analítica de Datos" → "5886 PERIODO 2026-1 ANALITICA DE DATOS"
   - "Estadística Inferencial" → "6023 PERIODO 2026-1 ESTADISTICA INFERENCIAL"
   - "Persona y Cultura" → "5138 PERIODO 2026-1 P&CIII PERSONA Y..."
   Pick exactly ONE card. If the target is not visible, scroll the grid (down then up). Expand "Hidden" section only if still missing.
5. Click the card BODY (not the "..." menu in its top-right).
6. Screenshot, verify the team panel shows the class name.
7. Click "General" in the left panel under that team, then click the "Posts" / "Publicaciones" tab.

The class name Daniel gives may be partial — match by KEYWORDS, not exact string.
Ctrl+E search is a last-resort fallback only if the grid approach fails.

ANTI-LOOP RULES (critical):
- If you took 2 screenshots in a row without clicking or scrolling, you MUST scroll now.
- If you scrolled 8 times in one direction and still haven't found the post, switch direction. If still not found after both directions, report and stop.
- Never take more than 3 consecutive screenshots without a click or scroll.

SCROLL TECHNIQUE for Teams Posts:
- Scroll the Posts feed at the center of the content area (e.g. x=700, y=500).
- Try direction="up" FIRST (recent assignments often at the top), then direction="down" if not found.
- Use clicks=8 per scroll. Screenshot after each scroll.

ASSIGNMENT BUTTON:
- Each assignment shows as a card posted by the "Assignments" bot, with a title (e.g. "Taller 2"), a "Vencimiento ..." line, and a "Ver tarea" button.
- Verify the card TITLE matches the assignment Daniel asked for BEFORE clicking. If multiple cards exist, scroll until the right title is visible — do not click the wrong one.
- Click "Ver tarea" precisely — small button below the title on that specific card.

REFERENCE MATERIALS:
- After clicking "Ver tarea" the assignment detail opens.
- If files are attached (PDF, DOCX, PPTX icons), click each to download.
- Files download to C:/Users/dani1/Downloads — use read_file on each to extract content for the writing phase.

When you have retrieved everything, prefix reply with: ASSIGNMENT_CONTENT:
Include both the instructions text AND any content extracted from reference materials.
NEVER skip tool calls or fake results."""

IMPROVEMENT_BRAIN_PROMPT = """You are Adam's self-improvement engine. A task failed and you must fix it permanently — without breaking what works.

You will receive:
- ORIGINAL TASK: what Daniel asked
- FAILURE: what mini (or previous attempt) returned
- CONVERSATION HISTORY: full context

MANDATORY WORKFLOW — follow in order:

STEP 1 — LEARN FROM PAST:
Call read_improvement_log to see what's been tried before. If the same fix already failed, DO NOT repeat it.

STEP 2 — UNDERSTAND CURRENT STATE:
Call read_bot_code("tools.py") to see ALL existing tools and function names.
DO NOT add a function that already exists — modify the existing one instead.

STEP 3 — PICK THE RIGHT FIX:
a) DIFFERENT APPROACH — use existing tools smarter. No code edit. Try this FIRST.
b) MODIFY EXISTING TOOL — function exists but is broken/incomplete. Edit it.
c) NEW TOOL — only if no existing tool covers it. Pick a UNIQUE name.
d) INSTALL PACKAGE — only if a library is genuinely missing.

STEP 4 — IF EDITING CODE:
- edit_bot_code now AUTO-VALIDATES: rejects syntax errors, rejects duplicate function names, runs import smoke test, auto-rolls back on import failure.
- If edit_bot_code returns ❌ — read the error, fix YOUR mistake, do not blindly retry.
- When adding a NEW tool: add 3 things in tools.py:
    1) An entry in the TOOLS list
    2) The implementation function (UNIQUE name)
    3) An elif branch in execute_tool
- After successful edit → save_pending_task → restart_bot.
- If unsure what to add, prefer approach (a) and don't touch code.

STEP 5 — LOG WHAT YOU DID:
Call write_improvement_log with:
- status: TRIED_APPROACH / ADDED_TOOL / INSTALLED_PKG / DIRECT_RETRY / GAVE_UP
- detail: what you tried and why
- snippet: tool name or key code change (optional)
This is mandatory so future rounds don't repeat the same failed fix.

STEP 6 — VERIFY:
After restart, the task is re-attempted automatically. If it fails again, you'll see it in the improvement log on the next round.

HARD RULES:
- NEVER define a function with a name that already exists in tools.py
- NEVER restart without first running edit_bot_code successfully
- If edit_bot_code rejects your edit, fix the code — DO NOT call restart_bot anyway
- Keep diffs minimal. One change at a time.
- When in doubt, prefer approach (a) — no code change."""

SYSTEM_PROMPT = load_system_prompt()
conversation_history = []

FAKE_ACTION_PHRASES = [
    "i've set", "i have set", "i've opened", "i have opened",
    "i've sent", "i have sent", "i've played", "i have played",
    "i've created", "i have created", "i've saved", "i have saved",
    "i've increased", "i've decreased", "i've changed", "i've turned",
    "volume is now", "brightness is now", "i've closed", "i've launched",
    "i've downloaded", "i've deleted", "i've moved", "i've written",
    "done!", "it's done", "task completed", "i've run", "i have run",
]


TG_MAX = 4000  # Telegram limit is 4096, keep buffer

def clean_markdown(text: str) -> str:
    """Strip markdown that breaks Telegram's parser."""
    import re
    text = text.replace("**", "").replace("__", "")
    # Remove unbalanced backticks
    text = re.sub(r'`{1,3}', '', text)
    # Remove unbalanced underscores/asterisks used as emphasis
    text = re.sub(r'(?<!\w)[_*]|[_*](?!\w)', '', text)
    return text.strip()

async def safe_send(update: Update, text: str, label: str = ""):
    """Send text reply, stripping bad markdown and splitting if too long."""
    text = (text or "").replace("**", "")
    full = f"_{label}_\n\n{text}" if label else text
    # Split into chunks if too long
    chunks = [full[i:i+TG_MAX] for i in range(0, max(len(full), 1), TG_MAX)]
    for chunk in chunks:
        try:
            await update.message.reply_text(chunk, parse_mode=ParseMode.MARKDOWN)
        except Exception:
            try:
                await update.message.reply_text(clean_markdown(chunk))
            except Exception:
                await update.message.reply_text(chunk[:TG_MAX])


FAILURE_PHRASES = [
    "i can't", "i cannot", "unable to", "not able to", "i'm not able",
    "don't have access", "i don't have the ability", "i don't know how",
    "max iterations reached", "cannot perform", "beyond my capabilities",
    "i'm unable", "no puedo", "no tengo acceso", "no es posible",
    "unfortunately", "i apologize", "i'm sorry, but i can't",
]

def looks_like_failure(text: str) -> bool:
    t = text.lower()
    return any(p in t for p in FAILURE_PHRASES)


def save_memory(role: str, content: str):
    path = f"{MEMORY_DIR}/{date.today()}.md"
    with open(path, "a", encoding="utf-8") as f:
        f.write(f"\n**{role}:** {content}\n")


async def run_tool_loop(messages: list, worker_model: str, update: Update, max_iterations: int = 40) -> str:
    """Run the tool-calling loop with the worker model. Returns final text response."""
    fake_strikes = 0
    iteration = 0
    recent_signatures = []  # track recent tool calls to detect stalls
    consecutive_screenshots = 0
    while iteration < max_iterations:
        iteration += 1
        response = client.chat.completions.create(
            model=worker_model,
            messages=messages,
            tools=TOOLS,
            tool_choice="auto",
        )
        msg = response.choices[0].message
        tool_calls = msg.tool_calls

        if not tool_calls:
            reply = (msg.content or "")
            # Catch fake actions
            if fake_strikes < 2 and any(p in reply.lower() for p in FAKE_ACTION_PHRASES):
                fake_strikes += 1
                messages.append({"role": "assistant", "content": reply})
                messages.append({
                    "role": "user",
                    "content": (
                        "You claimed to perform an action but called no tool. "
                        "Nothing happened. Call the appropriate tool now."
                    )
                })
                continue
            return reply

        messages.append(msg)
        pending_vision = []  # vision images to inject AFTER all tool responses

        # Build signature of this turn's tool calls for stall detection
        turn_sig = "|".join(sorted(tc.function.name + ":" + tc.function.arguments for tc in tool_calls))
        recent_signatures.append(turn_sig)
        if len(recent_signatures) > 4:
            recent_signatures.pop(0)

        # Count consecutive screenshot-only turns
        only_screenshots = all(tc.function.name in ("desktop_screenshot", "browser_screenshot") for tc in tool_calls)
        if only_screenshots:
            consecutive_screenshots += 1
        else:
            consecutive_screenshots = 0

        stall_detected = (
            len(recent_signatures) >= 3 and len(set(recent_signatures[-3:])) == 1
        ) or consecutive_screenshots >= 3

        for tc in tool_calls:
            tool_name = tc.function.name
            tool_args = json.loads(tc.function.arguments)
            await update.message.reply_text(f"Running: `{tool_name}`...")
            result = await execute_tool(tool_name, tool_args)

            if result.startswith("VISION_SCREENSHOT:"):
                img_path = result[len("VISION_SCREENSHOT:"):].strip()
                try:
                    with open(img_path, "rb") as img:
                        await update.message.reply_photo(img)
                except Exception:
                    pass
                try:
                    import base64
                    with open(img_path, "rb") as img:
                        b64 = base64.standard_b64encode(img.read()).decode()
                    pending_vision.append(b64)
                    result = "Screenshot taken. Analyze the image above to decide next action."
                except Exception as e:
                    result = f"Screenshot taken (encode failed: {e})"
                messages.append({"role": "tool", "tool_call_id": tc.id, "content": result})

            elif result.startswith("FILE_TO_SEND:"):
                file_path = result[len("FILE_TO_SEND:"):].strip()
                try:
                    with open(file_path, "rb") as f:
                        await update.message.reply_document(f, filename=os.path.basename(file_path))
                    if file_path.endswith("__adam.zip"):
                        os.remove(file_path)
                    result = f"Sent: {os.path.basename(file_path)}"
                except Exception as e:
                    result = f"Error sending file: {e}"
                messages.append({"role": "tool", "tool_call_id": tc.id, "content": result})

            elif tool_name == "browser_screenshot" and result.startswith("Screenshot saved:"):
                path = result.replace("Screenshot saved: ", "").strip()
                try:
                    with open(path, "rb") as img:
                        await update.message.reply_photo(img)
                except Exception:
                    pass
                messages.append({"role": "tool", "tool_call_id": tc.id, "content": result})

            else:
                messages.append({"role": "tool", "tool_call_id": tc.id, "content": result})

        # Inject vision images AFTER all tool responses (required by API message order)
        for b64 in pending_vision:
            messages.append({
                "role": "user",
                "content": [
                    {"type": "text", "text": "Current screen state — analyze this image and decide the next action:"},
                    {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64}"}},
                ],
            })

        if stall_detected:
            messages.append({
                "role": "user",
                "content": (
                    "STALL DETECTED. You repeated the same action or only took screenshots for 3 turns. "
                    "STOP repeating. Read the LAST screenshot carefully. Pick a NEW action: "
                    "either app_click on a visible element with coordinates from the image, "
                    "or app_scroll(x=700, y=500, direction='down', clicks=10) to reveal more, "
                    "or app_hotkey('pagedown'). Do NOT take another screenshot without a click or scroll first."
                ),
            })
            recent_signatures.clear()
            consecutive_screenshots = 0

    return "Max iterations reached. Last known state returned."


def save_pending_task(chat_id: int, task: str, history: list = None):
    """Persist task AND conversation history so bot can re-attempt with full context after restart."""
    with open(PENDING_TASK_FILE, "w", encoding="utf-8") as f:
        json.dump({
            "chat_id": chat_id,
            "task": task,
            "attempts": 0,
            "history": (history or conversation_history)[-10:]  # save last 10 messages for context
        }, f)

def load_pending_task() -> dict | None:
    try:
        with open(PENDING_TASK_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None

def clear_pending_task():
    try:
        os.remove(PENDING_TASK_FILE)
    except Exception:
        pass


async def escalate_and_improve(task: str, failure_reply: str, update: Update, history: list = None) -> str:
    """
    Sonnet analyzes failure → decides fix strategy:
    A) Direct retry with different approach (uses full conversation context)
    B) Write new tool → save pending task (with history) → restart (bot re-attempts after restart with full context)
    """
    await update.message.reply_text("🔧 Mini failed — Gemini analyzing and fixing...")

    # Include recent conversation so Sonnet has full context
    recent_history = (history or conversation_history)[-16:]

    improvement_msgs = [
        {"role": "system", "content": SYSTEM_PROMPT + "\n\n" + IMPROVEMENT_BRAIN_PROMPT},
        *recent_history,  # ← full context: what Daniel said, what happened before
        {"role": "user", "content": (
            f"ORIGINAL TASK: {task}\n\n"
            f"MINI FAILURE: {failure_reply}\n\n"
            "You have the full conversation above. Analyze why mini failed and fix it. "
            "If you can solve it directly with existing tools — do it now. "
            "If you need to add a new tool: read_bot_code('tools.py') first, add the tool, "
            "then call save_pending_task (pass the full conversation history as the 'history' parameter) and restart_bot.\n"
            f"chat_id for save_pending_task = {update.effective_chat.id}\n"
            f"IMPORTANT: When calling save_pending_task, pass the conversation history so context is preserved after restart."
        )},
    ]
    result = await run_tool_loop(improvement_msgs, MODEL_BRAIN, update, max_iterations=30)
    return result


async def resume_pending_task(bot: Bot):
    """Called on startup — if a pending task exists, re-attempt it."""
    pending = load_pending_task()
    if not pending:
        return

    chat_id  = pending["chat_id"]
    task     = pending["task"]
    attempts = pending.get("attempts", 0)
    history  = pending.get("history", [])

    # Hard cap — never loop more than 2 times
    if attempts >= 2:
        clear_pending_task()
        await bot.send_message(chat_id,
            f"⛔ Gave up after {attempts} attempts on: _{task[:80]}_\n"
            "Send !cancel if it keeps looping, or describe the issue differently.",
            parse_mode=ParseMode.MARKDOWN)
        return

    clear_pending_task()

    await bot.send_message(chat_id, f"🔄 Restarted after self-improvement. Re-attempting with full context...", parse_mode=ParseMode.MARKDOWN)

    # Run task with Sonnet — restore conversation history so Sonnet has FULL context
    msgs = [
        {"role": "system", "content": SYSTEM_PROMPT},
        *history,  # ← restore the conversation that led to this task
        {"role": "user", "content": (
            f"[SYSTEM NOTE: Adam just restarted after a self-improvement code edit. "
            f"The conversation above is the history that led to the task. "
            f"Now execute the task using the new/updated tools available.]"
        )},
    ]

    class _FakeUpdate:
        """Minimal shim so run_tool_loop can send messages."""
        class _FakeMsg:
            def __init__(self, bot, chat_id):
                self._bot = bot
                self._chat_id = chat_id
            async def reply_text(self, text, **kwargs):
                try:
                    await self._bot.send_message(self._chat_id, text, **kwargs)
                except Exception:
                    await self._bot.send_message(self._chat_id, text)
            async def reply_photo(self, photo, **kwargs):
                await self._bot.send_photo(self._chat_id, photo)
            async def reply_voice(self, voice, **kwargs):
                await self._bot.send_voice(self._chat_id, voice)
            async def reply_document(self, doc, **kwargs):
                await self._bot.send_document(self._chat_id, doc)
        def __init__(self, bot, chat_id):
            self.message = self._FakeMsg(bot, chat_id)
            self.effective_chat = type("C", (), {"id": chat_id})()

    fake_update = _FakeUpdate(bot, chat_id)
    reply = await run_tool_loop(msgs, MODEL_BRAIN, fake_update, max_iterations=25)

    # If still failing and haven't retried too many times, try once more
    if looks_like_failure(reply) and attempts < 2:
        save_pending_task(chat_id, task, history)
        with open(PENDING_TASK_FILE, "r") as f:
            d = json.load(f)
        d["attempts"] = attempts + 1
        with open(PENDING_TASK_FILE, "w") as f:
            json.dump(d, f)
        await bot.send_message(chat_id, f"⚠️ Still failing after improvement. Attempt {attempts+1}/2. Trying again...")
        # Trigger another improvement round
        improvement_msgs = [
            {"role": "system", "content": SYSTEM_PROMPT + "\n\n" + IMPROVEMENT_BRAIN_PROMPT},
            {"role": "user", "content": f"ORIGINAL TASK: {task}\n\nFAILURE (attempt {attempts+1}): {reply}\n\nStill failing. Try a completely different approach or add a better tool."},
        ]
        reply = await run_tool_loop(improvement_msgs, MODEL_BRAIN, fake_update, max_iterations=30)

    result_text = clean_markdown(f"✅ Result:\n\n{reply}")
    chunks = [result_text[i:i+TG_MAX] for i in range(0, max(len(result_text),1), TG_MAX)]
    for chunk in chunks:
        try:
            await bot.send_message(chat_id, chunk)
        except Exception:
            pass


async def run_assignment_pipeline(user_message: str, update: Update):
    """Shared assignment pipeline — called from both text and voice handlers."""
    await update.message.reply_text("📚 Starting assignment pipeline...")

    # Phase 1: Brain creates retrieval plan
    brain_msgs = [
        {"role": "system", "content": ASSIGNMENT_BRAIN_PROMPT},
        {"role": "user", "content": user_message},
    ]
    brain_resp = client.chat.completions.create(model=MODEL_BRAIN, messages=brain_msgs)
    plan = (brain_resp.choices[0].message.content or "").strip()

    await update.message.reply_text("🧠 Plan ready. Retrieving assignment...")

    # Phase 2: Worker executes retrieval (Gemini — navigates Teams + reads screenshots)
    worker_msgs = [
        {"role": "system", "content": SYSTEM_PROMPT + "\n\n" + ASSIGNMENT_WORKER_PROMPT},
        {"role": "user", "content": (
            f"TASK: {user_message}\n\nPLAN:\n{plan}\n\n"
            "REMINDER: Teams is already open as a desktop app. "
            "Do NOT call teams_login or any browser tool. "
            "Use app_focus_window('Teams') then desktop_screenshot to start."
        )},
    ]
    retrieval_result = await run_tool_loop(worker_msgs, MODEL_BRAIN, update)

    await update.message.reply_text("🧠 Assignment retrieved. Writing it...")

    # Phase 3: Brain writes the assignment
    write_msgs = [
        {"role": "system", "content": ASSIGNMENT_BRAIN_PROMPT},
        {"role": "user", "content": user_message},
        {"role": "assistant", "content": plan},
        {"role": "user", "content": (
            f"The worker retrieved the following assignment content (instructions + any reference material extracted via read_file):\n\n{retrieval_result}\n\n"
            "Now write the COMPLETE assignment, in full — not a skeleton, not a TODO list. "
            "Ground every claim in the reference material when one was attached; cite specific definitions, formulas, examples from it. "
            "Then produce a WRITE_PLAN: that picks the right output format from the PHASE 2 list "
            "(Python → create_python_file, notebook → create_jupyter_notebook, formal report → create_formal_word_document, "
            "LaTeX/typeset → create_latex_document, PDF → create_pdf, slides → create_powerpoint, spreadsheet → create_spreadsheet, "
            "code/text → create_text_file, simple .docx → create_word_document). "
            "If the assignment is unambiguous about format (e.g. asks for .py, .ipynb, .pdf, slides, spreadsheet) — match exactly. "
            "If unspecified, default to create_formal_word_document. "
            "Save under C:/Users/dani1/Desktop/<descriptive_filename>. Include the FULL written content as the function argument — no placeholders."
        )},
    ]
    write_resp = client.chat.completions.create(model=MODEL_BRAIN, messages=write_msgs)
    write_plan = (write_resp.choices[0].message.content or "").strip()

    await update.message.reply_text("⚡ Saving assignment to Desktop...")

    # Phase 4: Gemini calls create_* tools to save document to Desktop
    save_msgs = [
        {"role": "system", "content": SYSTEM_PROMPT + "\n\n" + ASSIGNMENT_WORKER_PROMPT},
        {"role": "user", "content": f"TASK: Save the assignment document.\n\nPLAN:\n{write_plan}"},
    ]
    save_result = await run_tool_loop(save_msgs, MODEL_BRAIN, update)

    reply = f"✅ Assignment done!\n\n{save_result}"
    conversation_history.append({"role": "assistant", "content": reply})
    await safe_send(update, reply or "", label="📚 Assignment")


async def handle_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """!cancel — wipe pending task, stop any queued self-improvement loop."""
    if update.effective_user.id != ALLOWED_USER_ID:
        return
    if os.path.exists(PENDING_TASK_FILE):
        try:
            with open(PENDING_TASK_FILE) as f:
                data = json.load(f)
            task_preview = data.get("task", "unknown")[:80]
        except Exception:
            task_preview = "unknown"
        clear_pending_task()
        await update.message.reply_text(f"🗑 Pending task cleared.\nWas: _{task_preview}_", parse_mode=ParseMode.MARKDOWN)
    else:
        await update.message.reply_text("✅ No pending task queued.")


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id != ALLOWED_USER_ID:
        await update.message.reply_text("Unauthorized.")
        return

    user_message = update.message.text

    # Quick commands
    if user_message.strip().lower() in ("!cancel", "!clear", "!stop"):
        await handle_cancel(update, context)
        return
    model, model_label, orchestrated = pick_mode(user_message)
    print(f"[DISPATCH] user={update.effective_user.id} | model={model} | label={model_label} | msg={user_message[:60]}")

    for prefix in ["!opus", "!gemini", "!haiku", "!planner", "!assignment"]:
        if user_message.startswith(prefix):
            user_message = user_message[len(prefix):].strip()
            break

    conversation_history.append({"role": "user", "content": user_message})

    if model_label == "📚 Assignment":
        await run_assignment_pipeline(user_message, update)
        return

    if orchestrated:
        # --- ORCHESTRATED MODE: Brain plans, Worker executes ---

        # Brain: understand and plan
        brain_messages = [
            {"role": "system", "content": BRAIN_PROMPT},
            *conversation_history[-12:],
        ]
        brain_resp = client.chat.completions.create(
            model=MODEL_BRAIN,
            messages=brain_messages,
        )
        plan = (brain_resp.choices[0].message.content or "").strip()

        # If brain says it's conversational, reply directly
        if plan.startswith("CONVERSATIONAL:"):
            reply = plan[len("CONVERSATIONAL:"):].strip()
            conversation_history.append({"role": "assistant", "content": reply})
            await safe_send(update, reply, label=model_label)
            return

        # Worker: execute the plan
        worker_messages = [
            {"role": "system", "content": SYSTEM_PROMPT + "\n\n" + WORKER_PROMPT},
            {"role": "user", "content": f"TASK: {user_message}\n\nPLAN:\n{plan}"},
        ]
        reply = await run_tool_loop(worker_messages, MODEL_WORKER, update, max_iterations=15)

        # Auto-escalate to Sonnet if mini failed
        if looks_like_failure(reply) and model != MODEL_BRAIN:
            reply = await escalate_and_improve(user_message, reply, update, conversation_history)

    else:
        # --- DIRECT MODE: single model handles everything ---
        messages = [{"role": "system", "content": SYSTEM_PROMPT}, *conversation_history]
        reply = await run_tool_loop(messages, model, update, max_iterations=15)

        # Auto-escalate to Sonnet if mini failed
        if looks_like_failure(reply) and model == MODEL_WORKER:
            reply = await escalate_and_improve(user_message, reply, update, conversation_history)

    conversation_history.append({"role": "assistant", "content": reply})
    await safe_send(update, reply or "", label=model_label)


# --- Voice: Whisper STT via Groq (free, fast) ---
def _get_groq_whisper_client():
    from tools import _load_creds
    key = _load_creds().get("groq", {}).get("api_key", "")
    if not key:
        return None
    # Groq is OpenAI-compatible, just different base_url
    return OpenAI(api_key=key, base_url="https://api.groq.com/openai/v1")

async def transcribe_voice(ogg_path: str) -> str:
    """Transcribe a Telegram .ogg voice file using Groq Whisper. Returns text or empty string."""
    try:
        mp3_path = ogg_path.replace(".ogg", ".mp3")
        subprocess.run(
            ["ffmpeg", "-y", "-i", ogg_path, mp3_path],
            capture_output=True, check=True
        )
        groq = _get_groq_whisper_client()
        if groq:
            with open(mp3_path, "rb") as f:
                result = groq.audio.transcriptions.create(
                    model="whisper-large-v3",
                    file=f,
                    response_format="text",
                )
            return result.strip()
        else:
            # Fallback: free Google STT via SpeechRecognition
            import speech_recognition as sr
            from pydub import AudioSegment
            wav_path = ogg_path.replace(".ogg", ".wav")
            AudioSegment.from_mp3(mp3_path).export(wav_path, format="wav")
            r = sr.Recognizer()
            with sr.AudioFile(wav_path) as src:
                audio = r.record(src)
            return r.recognize_google(audio, language="es-ES")
    except Exception as e:
        return ""

async def text_to_voice(text: str) -> str | None:
    """Convert text to speech using edge-tts. Returns .mp3 path or None."""
    try:
        import edge_tts
        # Clean text: strip markdown, emojis, keep plain speech
        clean = text.replace("**", "").replace("__", "").replace("`", "")
        # Trim to reasonable length for voice
        if len(clean) > 800:
            clean = clean[:800] + "..."
        out_path = os.path.join(tempfile.gettempdir(), "adam_reply.mp3")
        communicate = edge_tts.Communicate(clean, voice="es-MX-JorgeNeural")
        await communicate.save(out_path)
        return out_path
    except Exception as e:
        return None

async def handle_voice(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle incoming Telegram voice messages."""
    if update.effective_user.id != ALLOWED_USER_ID:
        return

    await update.message.reply_text("🎤 Escuchando...")

    # Download .ogg from Telegram
    voice = update.message.voice
    tg_file = await context.bot.get_file(voice.file_id)
    ogg_path = os.path.join(tempfile.gettempdir(), "adam_voice_in.ogg")
    await tg_file.download_to_drive(ogg_path)

    # Transcribe
    transcribed = await transcribe_voice(ogg_path)
    if not transcribed:
        await update.message.reply_text("❌ No pude entender el audio. Intenta de nuevo.")
        return

    await update.message.reply_text(f"🎤 _{transcribed}_", parse_mode=ParseMode.MARKDOWN)

    model, model_label, orchestrated = pick_mode(transcribed)

    for prefix in ["!opus", "!gemini", "!haiku", "!planner", "!assignment"]:
        if transcribed.startswith(prefix):
            transcribed = transcribed[len(prefix):].strip()
            break

    conversation_history.append({"role": "user", "content": transcribed})

    # Run through normal pipeline
    if model_label == "📚 Assignment":
        # Run assignment pipeline directly with transcribed text
        await run_assignment_pipeline(transcribed, update)
        return

    if orchestrated:
        brain_messages = [{"role": "system", "content": BRAIN_PROMPT}, *conversation_history[-12:]]
        brain_resp = client.chat.completions.create(model=MODEL_BRAIN, messages=brain_messages)
        plan = (brain_resp.choices[0].message.content or "").strip()
        if plan.startswith("CONVERSATIONAL:"):
            reply = plan[len("CONVERSATIONAL:"):].strip()
        else:
            worker_messages = [
                {"role": "system", "content": SYSTEM_PROMPT + "\n\n" + WORKER_PROMPT},
                {"role": "user", "content": f"TASK: {transcribed}\n\nPLAN:\n{plan}"},
            ]
            reply = await run_tool_loop(worker_messages, MODEL_WORKER, update, max_iterations=15)
            if looks_like_failure(reply) and model != MODEL_BRAIN:
                reply = await escalate_and_improve(transcribed, reply, update, conversation_history)
    else:
        messages = [{"role": "system", "content": SYSTEM_PROMPT}, *conversation_history]
        reply = await run_tool_loop(messages, model, update, max_iterations=15)
        if looks_like_failure(reply) and model == MODEL_WORKER:
            reply = await escalate_and_improve(transcribed, reply, update, conversation_history)

    conversation_history.append({"role": "assistant", "content": reply})
    reply_clean = (reply or "").replace("**", "")
    await safe_send(update, reply_clean, label=model_label)

    # Send voice reply
    voice_path = await text_to_voice(reply_clean)
    if voice_path and os.path.exists(voice_path):
        with open(voice_path, "rb") as vf:
            await update.message.reply_voice(vf)


async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle incoming Telegram photo messages with optional caption as instruction."""
    if update.effective_user.id != ALLOWED_USER_ID:
        return

    await update.message.reply_text("🖼️ Analizando imagen...")

    # Download highest-res photo
    photo = update.message.photo[-1]  # last = largest
    tg_file = await context.bot.get_file(photo.file_id)
    import tempfile as _tmp
    img_path = os.path.join(_tmp.gettempdir(), "adam_photo_in.jpg")
    await tg_file.download_to_drive(img_path)

    # Encode to base64
    import base64
    try:
        with open(img_path, "rb") as f:
            b64 = base64.standard_b64encode(f.read()).decode()
    except Exception as e:
        await update.message.reply_text(f"❌ Error leyendo imagen: {e}")
        return

    # Caption = Daniel's instruction (may be None)
    caption = (update.message.caption or "").strip()
    instruction = caption if caption else "Describe what you see in this image and tell me if you notice anything important or actionable."

    # Build vision message for Sonnet
    vision_content = [
        {"type": "text", "text": instruction},
        {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64}"}},
    ]

    conversation_history.append({"role": "user", "content": vision_content})

    # Run through Sonnet directly (vision-capable)
    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        *conversation_history,
    ]
    reply = await run_tool_loop(messages, MODEL_BRAIN, update, max_iterations=20)

    conversation_history.append({"role": "assistant", "content": reply})
    await safe_send(update, reply or "", label="🖼️ Vision")


async def post_init(application):
    """Runs once after bot starts — resumes any pending task from before restart."""
    await resume_pending_task(application.bot)

app = ApplicationBuilder().token(TOKEN).post_init(post_init).build()
app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
app.add_handler(MessageHandler(filters.VOICE, handle_voice))
app.add_handler(MessageHandler(filters.PHOTO, handle_photo))

print("Adam running | Self-improving | Voice enabled | Press Ctrl+C to stop")
app.run_polling()
