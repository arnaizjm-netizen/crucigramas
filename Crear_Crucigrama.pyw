"""
Doble clic para generar un crucigrama jugable a partir de un .docx.
No requiere terminal: abre un dialogo para elegir el archivo Word,
genera el HTML al lado de este script y ofrece abrirlo directamente.
"""
import os
import sys
import webbrowser
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)

from generate_crossword import build_html, slugify  # noqa: E402
from parse_docx import CrosswordParseError  # noqa: E402


def main():
    root = tk.Tk()
    root.withdraw()

    docx_path = filedialog.askopenfilename(
        title="Elige el archivo Word del crucigrama",
        filetypes=[("Documentos Word", "*.docx")],
    )
    if not docx_path:
        return  # el usuario ha cancelado

    try:
        html_out, name = build_html(Path(docx_path))
    except CrosswordParseError as e:
        messagebox.showerror("No se pudo generar", f"Hubo un problema con el documento:\n\n{e}")
        return
    except Exception as e:  # noqa: BLE001
        messagebox.showerror("No se pudo generar", f"Ha ocurrido un error inesperado:\n\n{e}")
        return

    out_path = os.path.join(SCRIPT_DIR, f"{slugify(name)}.html")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html_out)

    abrir = messagebox.askyesno(
        "Crucigrama generado",
        f"'{name}' se ha generado correctamente.\n\n¿Quieres abrirlo ahora para jugar?",
    )
    if abrir:
        webbrowser.open(f"file:///{out_path}")


if __name__ == "__main__":
    main()
