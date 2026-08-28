#!/usr/bin/env python
"""
Genera un HTML jugable y autocontenido a partir de un .docx de crucigrama.

Uso:
    python generate_crossword.py "Crucigrama_001_tablas_Word.docx" [salida.html]

Si no se indica archivo de salida, se genera "<nombre_del_crucigrama>.html"
en el mismo directorio que este script.
"""
from __future__ import annotations

import html
import json
import re
import sys
from pathlib import Path

from parse_docx import parse_crossword_docx, CrosswordParseError

SCRIPT_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = SCRIPT_DIR / "template.html"


def slugify(name: str) -> str:
    slug = re.sub(r"[^A-Za-z0-9]+", "_", name).strip("_")
    return slug or "crucigrama"


def build_html(docx_path: Path) -> tuple[str, str]:
    crossword = parse_crossword_docx(str(docx_path))
    template = TEMPLATE_PATH.read_text(encoding="utf-8")

    data_json = json.dumps(crossword.to_json_data(), ensure_ascii=False)
    # evita que un "</script>" dentro de una definicion corte el <script> del HTML
    data_json = data_json.replace("</", "<\\/")

    safe_title = html.escape(crossword.name, quote=True)

    output = template.replace("__DATA_JSON__", data_json)
    output = output.replace("__TITLE__", safe_title)

    return output, crossword.name


def main(argv: list[str]) -> int:
    if len(argv) < 2:
        print("Uso: python generate_crossword.py <archivo.docx> [salida.html]", file=sys.stderr)
        return 1

    docx_path = Path(argv[1])
    if not docx_path.exists():
        print(f"No existe el archivo: {docx_path}", file=sys.stderr)
        return 1

    try:
        html_out, name = build_html(docx_path)
    except CrosswordParseError as e:
        print(f"Error al interpretar el documento: {e}", file=sys.stderr)
        return 2

    if len(argv) >= 3:
        out_path = Path(argv[2])
    else:
        out_path = SCRIPT_DIR / f"{slugify(name)}.html"

    out_path.write_text(html_out, encoding="utf-8")
    print(f"Crucigrama '{name}' generado en: {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
