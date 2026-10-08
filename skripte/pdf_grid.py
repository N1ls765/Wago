"""Rasterbasierte Tabellenextraktion fuer das WAGO-PEP-PDF.

Warum nicht pdfplumber.extract_table()?
  Die Zeichen in den eingebetteten Excel-Objekten haben eine Textmatrix mit
  winzigen Nebendiagonal-Termen (~1e-08). pdfplumber stuft sie deshalb als
  upright=False ein und zerlegt jedes Zeichen in ein eigenes Wort
  ("Im p a c t c a t e g o r y"). Wir umgehen die Wortlogik komplett und
  setzen die Zeichen anhand ihrer Koordinaten selbst zu Zellen zusammen.
"""

import pdfplumber


def _snap(values, tol=1.5):
    """Nahe beieinanderliegende Koordinaten zu einer zusammenfassen."""
    out = []
    for v in sorted(values):
        if not out or v - out[-1] > tol:
            out.append(v)
    return out


def grid_lines(page, tol=1.5, min_len=20):
    """Spalten- und Zeilengrenzen aus gezeichneten Linien und Rechtecken."""
    xs, ys = [], []
    for ln in page.lines:
        if abs(ln["x0"] - ln["x1"]) < 0.6 and abs(ln["bottom"] - ln["top"]) >= min_len:
            xs.append(ln["x0"])
        if abs(ln["top"] - ln["bottom"]) < 0.6 and abs(ln["x1"] - ln["x0"]) >= min_len:
            ys.append(ln["top"])
    for r in page.rects:
        if r["width"] < 2 and r["height"] >= min_len:
            xs.append(r["x0"])
        elif r["height"] < 2 and r["width"] >= min_len:
            ys.append(r["top"])
        elif r["width"] >= min_len and r["height"] >= min_len:
            xs.extend([r["x0"], r["x1"]])
            ys.extend([r["top"], r["bottom"]])
    return _snap(xs, tol), _snap(ys, tol)


def cell_text(chars, y_tol=2.5, join=" "):
    """Zeichen einer Zelle zu Text zusammensetzen.

    Zeichen werden nach Zeile (top) gruppiert, innerhalb der Zeile nach x0
    sortiert und direkt aneinandergehaengt - die Leerzeichen stehen als
    echte Zeichen im PDF, es muss also nichts anhand von Luecken geraten
    werden.
    """
    if not chars:
        return ""
    rows, current, last_top = [], [], None
    for ch in sorted(chars, key=lambda c: (c["top"], c["x0"])):
        if last_top is None or abs(ch["top"] - last_top) <= y_tol:
            current.append(ch)
        else:
            rows.append(current)
            current = [ch]
        last_top = ch["top"]
    rows.append(current)
    lines = ["".join(c["text"] for c in sorted(r, key=lambda c: c["x0"])).strip()
             for r in rows]
    return join.join(l for l in lines if l).strip()


def extract_grid(page, xs=None, ys=None, y_tol=2.5):
    """Seite anhand des Rasters in eine Liste von Zeilen zerlegen."""
    if xs is None or ys is None:
        gx, gy = grid_lines(page)
        xs = xs if xs is not None else gx
        ys = ys if ys is not None else gy
    if len(xs) < 2 or len(ys) < 2:
        return []

    buckets = {}
    for ch in page.chars:
        cx = (ch["x0"] + ch["x1"]) / 2
        cy = (ch["top"] + ch["bottom"]) / 2
        col = row = None
        for i in range(len(xs) - 1):
            if xs[i] <= cx < xs[i + 1]:
                col = i
                break
        for j in range(len(ys) - 1):
            if ys[j] <= cy < ys[j + 1]:
                row = j
                break
        if col is not None and row is not None:
            buckets.setdefault((row, col), []).append(ch)

    return [[cell_text(buckets.get((r, c), []), y_tol) for c in range(len(xs) - 1)]
            for r in range(len(ys) - 1)]


def open_pdf(path):
    return pdfplumber.open(path)
