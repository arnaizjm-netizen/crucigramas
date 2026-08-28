"""
Parser de crucigramas en formato .docx (tablas de Word) -> estructura de datos Python.

Formato esperado del documento:
  1. Un parrafo con el nombre del crucigrama (puede llevar un sufijo tipo
     "-- DEFINICIONES" que se descarta).
  2. Una tabla NxN con fila y columna de cabecera numeradas (1..N) y celdas
     interiores en blanco (FFFFFF) o negro (000000).
  3. Parrafo "HORIZONTALES", seguido de N parrafos "num.- clue | clue | ...".
  4. Parrafo "VERTICALES", seguido de N parrafos "num.- clue | clue | ...".
  5. Un parrafo con "SOLUCION" en el texto (opcional, solo como separador).
  6. Una segunda tabla NxN, incidentes celdas con una letra mayuscula (blancas)
     o vacias (negras, 000000), con el mismo patron de negras que la tabla 1.

Solo depende de la libreria estandar (zipfile + xml.etree.ElementTree).
"""
from __future__ import annotations

import hashlib
import json
import re
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
FILL_TAG = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill"


class CrosswordParseError(ValueError):
    pass


def _get_text(elem: ET.Element) -> str:
    return "".join(t.text or "" for t in elem.findall(".//w:t", NS))


def _get_fill(tc: ET.Element) -> str | None:
    tcPr = tc.find("w:tcPr", NS)
    if tcPr is None:
        return None
    shd = tcPr.find("w:shd", NS)
    if shd is None:
        return None
    return shd.get(FILL_TAG)


def _load_body(docx_path: str) -> ET.Element:
    with zipfile.ZipFile(docx_path) as z:
        xml_bytes = z.read("word/document.xml")
    root = ET.fromstring(xml_bytes)
    body = root.find("w:body", NS)
    if body is None:
        raise CrosswordParseError("document.xml sin <w:body>")
    return body


def _parse_table(tbl: ET.Element) -> list[list[tuple[str, str | None]]]:
    grid = []
    for tr in tbl.findall("w:tr", NS):
        row = []
        for tc in tr.findall("w:tc", NS):
            row.append((_get_text(tc).strip(), _get_fill(tc)))
        grid.append(row)
    return grid


def _clean_name(raw: str) -> str:
    raw = raw.strip()
    for sep in ("—", "–", " - "):  # em dash, en dash, " - "
        idx = raw.find(sep)
        if idx != -1:
            return raw[:idx].strip()
    return raw


_CLUE_RE = re.compile(r"^(\d+)\D+?\s(.*)$")


def _parse_clue_line(line: str) -> tuple[int, list[str]]:
    m = _CLUE_RE.match(line.strip())
    if not m:
        raise CrosswordParseError(f"Linea de definicion no reconocida: {line!r}")
    num = int(m.group(1))
    clues = [c.strip() for c in m.group(2).split("|")]
    return num, clues


def _segments_in_line(is_black: list[bool]) -> list[tuple[int, int]]:
    segs: list[tuple[int, int]] = []
    start = None
    for i, black in enumerate(is_black):
        if not black:
            if start is None:
                start = i
        else:
            if start is not None:
                segs.append((start, i - 1))
                start = None
    if start is not None:
        segs.append((start, len(is_black) - 1))
    return segs


@dataclass
class Crossword:
    id: str
    name: str
    size: int
    black: list[list[bool]]
    solution: list[list[str]]
    cell_seg_h: list[list[str | None]]
    cell_seg_v: list[list[str | None]]
    segments: dict = field(default_factory=dict)

    def to_json_data(self) -> dict:
        return {
            "id": self.id,
            "name": self.name,
            "size": self.size,
            "black": self.black,
            "solution": self.solution,
            "cellSegH": self.cell_seg_h,
            "cellSegV": self.cell_seg_v,
            "segments": self.segments,
        }


def parse_crossword_docx(docx_path: str) -> Crossword:
    body = _load_body(docx_path)

    paragraphs: list[str] = []
    tables: list[list[list[tuple[str, str | None]]]] = []
    for child in body:
        tag = child.tag.split("}")[-1]
        if tag == "p":
            text = _get_text(child).strip()
            if text:
                paragraphs.append(text)
        elif tag == "tbl":
            tables.append(_parse_table(child))

    if len(tables) < 2:
        raise CrosswordParseError(
            f"Se esperaban 2 tablas (cuadrante y solucion), encontradas {len(tables)}"
        )
    if not paragraphs:
        raise CrosswordParseError("No se ha encontrado ningun parrafo de texto")

    name = _clean_name(paragraphs[0])

    # --- tabla 1: cuadrante vacio (con cabecera de fila/columna) ---
    grid_table = tables[0]
    n = len(grid_table) - 1
    if n <= 0:
        raise CrosswordParseError("La tabla del cuadrante esta vacia")
    black = []
    for row in grid_table[1:]:
        cells = row[1:]
        if len(cells) != n:
            raise CrosswordParseError("La tabla del cuadrante no es cuadrada")
        black.append([fill == "000000" for _text, fill in cells])

    # --- tabla 2: solucion (sin cabecera) ---
    sol_table = tables[1]
    if len(sol_table) != n or any(len(r) != n for r in sol_table):
        raise CrosswordParseError(
            f"La tabla de solucion no mide {n}x{n} como el cuadrante"
        )
    solution: list[list[str]] = []
    for r in range(n):
        row_letters = []
        for c in range(n):
            text, fill = sol_table[r][c]
            is_black = fill == "000000"
            if is_black != black[r][c]:
                raise CrosswordParseError(
                    f"Patron de negras distinto entre cuadrante y solucion en ({r + 1},{c + 1})"
                )
            if is_black:
                row_letters.append("")
            else:
                letter = text.strip().upper()
                if len(letter) != 1 or not letter.isalpha():
                    raise CrosswordParseError(
                        f"Casilla de solucion invalida en ({r + 1},{c + 1}): {text!r}"
                    )
                row_letters.append(letter)
        solution.append(row_letters)

    # --- definiciones ---
    horiz: dict[int, list[str]] = {}
    vert: dict[int, list[str]] = {}
    mode = None
    for p in paragraphs[1:]:
        upper = p.upper()
        if upper.startswith("HORIZONTALES"):
            mode = "H"
            continue
        if upper.startswith("VERTICALES"):
            mode = "V"
            continue
        if "SOLUC" in upper:
            break
        if mode is None:
            continue
        num, clues = _parse_clue_line(p)
        (horiz if mode == "H" else vert)[num] = clues

    # --- segmentos por fila ---
    cell_seg_h: list[list[str | None]] = [[None] * n for _ in range(n)]
    cell_seg_v: list[list[str | None]] = [[None] * n for _ in range(n)]
    segments: dict[str, dict] = {}

    for r in range(n):
        row_num = r + 1
        segs = _segments_in_line(black[r])
        clues = horiz.get(row_num)
        if clues is None:
            raise CrosswordParseError(f"Faltan definiciones horizontales para la fila {row_num}")
        if len(clues) != len(segs):
            raise CrosswordParseError(
                f"Fila {row_num}: {len(segs)} tramos de casillas pero {len(clues)} definiciones"
            )
        for k, (c0, c1) in enumerate(segs):
            seg_id = f"H{row_num}_{k}"
            cells = [[r, c] for c in range(c0, c1 + 1)]
            segments[seg_id] = {
                "dir": "H",
                "label": str(row_num),
                "cells": cells,
                "clue": clues[k],
            }
            for c in range(c0, c1 + 1):
                cell_seg_h[r][c] = seg_id

    for c in range(n):
        col_num = c + 1
        col_black = [black[r][c] for r in range(n)]
        segs = _segments_in_line(col_black)
        clues = vert.get(col_num)
        if clues is None:
            raise CrosswordParseError(f"Faltan definiciones verticales para la columna {col_num}")
        if len(clues) != len(segs):
            raise CrosswordParseError(
                f"Columna {col_num}: {len(segs)} tramos de casillas pero {len(clues)} definiciones"
            )
        for k, (r0, r1) in enumerate(segs):
            seg_id = f"V{col_num}_{k}"
            cells = [[r, c] for r in range(r0, r1 + 1)]
            segments[seg_id] = {
                "dir": "V",
                "label": str(col_num),
                "cells": cells,
                "clue": clues[k],
            }
            for r in range(r0, r1 + 1):
                cell_seg_v[r][c] = seg_id

    # id estable: depende del contenido, no de la fecha, para no perder partidas guardadas
    digest_src = json.dumps(
        {"name": name, "black": black, "solution": solution}, sort_keys=True, ensure_ascii=False
    )
    cw_id = "cw_" + hashlib.sha256(digest_src.encode("utf-8")).hexdigest()[:12]

    return Crossword(
        id=cw_id,
        name=name,
        size=n,
        black=black,
        solution=solution,
        cell_seg_h=cell_seg_h,
        cell_seg_v=cell_seg_v,
        segments=segments,
    )
