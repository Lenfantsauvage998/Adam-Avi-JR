"""
Lightweight ISSE tools — no Adam/pc-agent dependencies.
Only needs: requests
"""
import os
from collections import defaultdict

import requests

ISSE_BASE_URL = os.environ.get("ISSE_BASE_URL", "https://isse-certificados.fly.dev")


def isse_get_ciclos() -> str:
    try:
        resp = requests.get(f"{ISSE_BASE_URL}/api/ciclos", timeout=10)
        resp.raise_for_status()
        ciclos = resp.json().get("ciclos", [])
        if not ciclos:
            return "No hay ciclos cargados en la base de datos ISSE."
        return "Ciclos disponibles en ISSE:\n" + "\n".join(f"  - {c}" for c in ciclos)
    except Exception as e:
        return f"Error consultando ISSE: {e}"


def isse_search_professor(name: str) -> str:
    try:
        resp = requests.get(f"{ISSE_BASE_URL}/api/professors", params={"q": name}, timeout=10)
        resp.raise_for_status()
        profs = resp.json().get("professors", [])
        if not profs:
            return f"No se encontró ningún profesor con el nombre '{name}' en ISSE."
        lines = [f"  - {p}" for p in profs[:15]]
        return f"Profesores encontrados ({len(profs)}):\n" + "\n".join(lines)
    except Exception as e:
        return f"Error buscando profesor en ISSE: {e}"


def isse_get_certificates(professor_id: str, ciclos: list = None) -> str:
    try:
        params = {"profesor": professor_id}
        if ciclos:
            params["ciclos"] = ",".join(ciclos)
        resp = requests.get(f"{ISSE_BASE_URL}/api/certificates", params=params, timeout=15)
        resp.raise_for_status()
        data = resp.json().get("results", [])
        if not data:
            if ciclos:
                ciclos_str = ", ".join(ciclos)
                return f"El profesor '{professor_id}' no tiene registros para los períodos: {ciclos_str}."
            return f"No se encontraron certificados para '{professor_id}'."
        by_ciclo = defaultdict(list)
        for row in data:
            by_ciclo[row.get("ciclo_lectivo", "")].append(row)
        lines = [f"Certificados de '{professor_id}' — {len(data)} registros en {len(by_ciclo)} período(s):\n"]
        for ciclo in sorted(by_ciclo.keys()):
            rows_c = by_ciclo[ciclo]
            total_hrs = sum(int(r.get("horas_semestre", 0) or 0) for r in rows_c)
            lines.append(f"{ciclo} ({len(rows_c)} cursos, {total_hrs} hrs totales):")
            for r in rows_c:
                curso = r.get("nombre_curso") or r.get("materia", "")
                hrs = int(r.get("horas_semestre", 0) or 0)
                dpto = r.get("departamento", "")
                lines.append(f"  - {curso} — {hrs} hrs — {dpto}")
        return "\n".join(lines)
    except Exception as e:
        return f"Error obteniendo certificados ISSE: {e}"


def isse_export_certificate(professor_name: str, ciclos: list = None, fmt: str = "excel") -> str:
    """Returns raw bytes of the file (for in-memory Telegram send)."""
    try:
        params = {"profesor": professor_name, "format": fmt}
        if ciclos:
            params["ciclos"] = ",".join(ciclos)
        resp = requests.get(f"{ISSE_BASE_URL}/api/export", params=params, timeout=30)
        resp.raise_for_status()
        return resp.content  # bytes — caller handles sending
    except Exception as e:
        return f"Error exportando certificado ISSE: {e}"


async def execute_tool(name: str, args: dict) -> str:
    if name == "isse_get_ciclos":
        return isse_get_ciclos()
    elif name == "isse_search_professor":
        return isse_search_professor(args["name"])
    elif name == "isse_get_certificates":
        return isse_get_certificates(args["professor_id"], ciclos=args.get("ciclos"))
    elif name == "isse_export_certificate":
        result = isse_export_certificate(
            args["professor_name"], ciclos=args.get("ciclos"), fmt=args.get("fmt", "excel")
        )
        if isinstance(result, bytes):
            return "__EXPORT_BYTES__"  # signal to bot to handle file send
        return result
    return f"Unknown tool: {name}"
