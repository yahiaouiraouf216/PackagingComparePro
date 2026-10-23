#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Packaging Compare Pro  -  version améliorée
===========================================
Dépendances :
    pip install PySide6 pymupdf pikepdf opencv-python numpy openpyxl
    (optionnel, décodage des codes-barres : pip install pyzbar pylibdmtx)

Architecture conservée : Viewer (QGraphicsView) + MainWindow avec
build_ui / load_original / load_modified / toggle_lock / reset_view /
toggle_blink / blink / auto_align / show_difference / show_colors /
set_zoom / set_opacity / pdf_to_png.

Principe général
----------------
* Les PDF sont lus nativement avec PyMuPDF (pas de PNG intermédiaire pour
  l'analyse des couleurs, des objets ou des séparations).
* Chaque page est rendue une fois en RGB (aperçu réel) et en "plaques"
  (C, M, Y, K + une plaque par encre spot) puis mise en cache.
* L'image du PDF modifié est un item déplaçable. Sa transformation de scène
  (déplacement manuel + auto-align) est la seule source de vérité : la vue
  Difference, Side by Side et la comparaison d'objets l'utilisent toutes.

Limites connues (honnêtes) :
* Plaques spot : on isole les /Separation en réécrivant leurs espaces colorimétriques
  dans une copie mémoire du PDF. Les /DeviceN multi-colorants et les images
  contenant des spots ne sont pas isolés individuellement.
* Les couleurs d'affichage des spots sont des approximations d'écran.
"""

import sys
import re
import math
import zlib
import hashlib
import datetime
import csv
from dataclasses import dataclass, field

import cv2
import fitz
import numpy as np

try:                      # pikepdf : lecture native des flux de contenu
    import pikepdf
except ImportError:       # l'application démarre quand même (inspecteurs désactivés)
    pikepdf = None

from PySide6.QtCore import Qt, QTimer, QPointF, QRectF, QLineF, QUrl, QMarginsF, Signal, QEvent
from PySide6.QtGui import (
    QPixmap, QImage, QTransform, QColor, QPen, QPainter, QPdfWriter,
    QTextDocument, QPageSize, QPageLayout, QIcon, QBrush, QShortcut, QKeySequence, QCursor
)
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QFileDialog, QPushButton, QLabel,
    QSlider, QVBoxLayout, QHBoxLayout, QGridLayout, QGraphicsView,
    QGraphicsScene, QGraphicsPixmapItem, QGraphicsItem, QMessageBox, QDockWidget,
    QListWidget, QListWidgetItem, QTableWidget, QTableWidgetItem, QHeaderView,
    QSpinBox, QComboBox, QSplitter, QFrame, QTabWidget, QTextEdit, QCheckBox
)

# ============================================================================
# CONSTANTES
# ============================================================================
RENDER_SCALE = 2.0          # 2.0 = 144 dpi (1 pt = 2 px)
MAX_SIDE_PX = 5000          # plafond de taille de rendu (gros formats)
DIFF_THRESHOLD = 40         # seuil de différence pixel (0-255)
TOL_OK_MM = 0.3             # <= 0.3 mm  -> OK
TOL_WARN_MM = 1.0           # <= 1.0 mm  -> WARNING, au-delà CRITICAL
IDENTICAL_MM = 0.05         # en dessous : objet considéré identique
PT_PER_MM = 72.0 / 25.4
MAX_DRAWINGS = 40000        # garde-fou performance (objets vectoriels)

# Couleurs d'affichage (approximations écran) des encres process
PROCESS = {
    "Cyan": (0, 174, 239),
    "Magenta": (236, 0, 140),
    "Yellow": (255, 242, 0),
    "Black": (35, 31, 32),
}
# Couleurs d'affichage (repli) : PMS connus + encres spéciales (le blanc/vernis
# sont invisibles sur fond blanc, on les teinte pour les voir dans l'aperçu)
KNOWN_PMS = {
    "pantone186c": (200, 16, 46), "pantone485c": (218, 41, 28),
    "pantonewarmred": (249, 66, 58), "pantonewarmredc": (249, 66, 58),
}
SPECIAL_RGB = {"White Ink": (170, 205, 235), "Varnish": (120, 180, 255),
               "Primer": (190, 170, 90), "Emboss": (130, 130, 130), "Braille": (120, 80, 200)}
CODE_KINDS = ("code-barres", "code-2D", "QR", "DataMatrix")
STATUS_COLORS = {"OK": "#2e9e44", "WARNING": "#e69500", "CRITICAL": "#d62828"}


# ============================================================================
# OUTILS GÉNÉRAUX
# ============================================================================
def norm(s):
    """Normalise un nom (minuscules, alphanumérique)."""
    return re.sub(r"[^a-z0-9]", "", s.lower())


def pdf_name_decode(n):
    """Décode un nom PDF (#20 -> espace)."""
    return re.sub(r"#([0-9A-Fa-f]{2})", lambda m: chr(int(m.group(1), 16)), n.lstrip("/"))


def classify_spot(name):
    """Détermine le type d'une encre spot d'après son nom."""
    n = norm(name)
    if "white" in n or "blanc" in n:
        return "White Ink"
    if "varnish" in n or "vernis" in n or "coating" in n:
        return "Varnish"
    if "primer" in n:
        return "Primer"
    if "emboss" in n or "gaufrage" in n or "relief" in n:
        return "Emboss"
    if "braille" in n:
        return "Braille"
    if "pantone" in n or "pms" in n:
        return "PMS"
    return "Spot"


def guess_rgb(name):
    """Couleur d'affichage déterministe pour un spot inconnu."""
    h = zlib.crc32(name.encode()) % 360
    c = QColor.fromHsv(h, 200, 220)
    return (c.red(), c.green(), c.blue())


def to_pixmap(arr):
    """Tableau numpy RGB uint8 -> QPixmap."""
    arr = np.ascontiguousarray(arr)
    h, w = arr.shape[:2]
    img = QImage(arr.data, w, h, 3 * w, QImage.Format_RGB888)
    return QPixmap.fromImage(img)


def grade(dx, dy):
    m = max(abs(dx), abs(dy))
    return "OK" if m <= TOL_OK_MM else ("WARNING" if m <= TOL_WARN_MM else "CRITICAL")


def center(b):
    return ((b[0] + b[2]) / 2.0, (b[1] + b[3]) / 2.0)


# ============================================================================
# LECTURE NATIVE DU PDF
# ============================================================================
# /Separation /Nom <alternate> <fonction>
SEP_RE = re.compile(
    r"\[\s*/Separation\s*(/[^\s/\[\]<>()]+)\s*"
    r"(?:/\w+|\[[^\]]*\]|\d+\s+\d+\s+R)\s*"
    r"(?:\d+\s+\d+\s+R|<<.*?>>)\s*\]", re.S)
# /DeviceN [noms] alternate fonction [attributs]
DEVN_RE = re.compile(
    r"\[\s*/DeviceN\s*\[([^\]]*)\]\s*(?:/\w+|\[[^\]]*\]|\d+\s+\d+\s+R)\s*"
    r"(?:\d+\s+\d+\s+R|<<.*?>>)(?:\s*(?:\d+\s+\d+\s+R|<<.*?>>))?\s*\]", re.S)
_NUM = rb"[-+]?(?:\d+\.?\d*|\.\d+)"
_K_RE = re.compile(rb"(?<![\w/])(?:" + _NUM + rb"\s+){4}([kK])(?![\w*])")
_RGB_RE = re.compile(rb"(?<![\w/])(?:" + _NUM + rb"\s+){3}(rg|RG)(?![\w*])")
_G_RE = re.compile(rb"(?<![\w/])" + _NUM + rb"\s+([gG])(?![\w*])")


def neutralize_process_colors(stream):
    """Passe en blanc toutes les couleurs process d'un flux de contenu,
    pour ne garder que l'encre spot ciblée lors du rendu d'une plaque."""
    stream = _K_RE.sub(lambda m: b"0 0 0 0 " + m.group(1), stream)
    stream = _RGB_RE.sub(lambda m: b"1 1 1 " + m.group(1), stream)
    stream = _G_RE.sub(lambda m: b"1 " + m.group(1), stream)
    return stream


@dataclass
class Obj:
    kind: str
    bbox: tuple
    key: str
    label: str
    attrs: dict = field(default_factory=dict)   # police, style, corps, couleur


@dataclass
class Finding:
    kind: str
    change: str      # déplacé / identique / ajouté / supprimé / modifié
    status: str      # OK / WARNING / CRITICAL
    label: str
    dx: float = 0.0  # mm (+ = droite)
    dy: float = 0.0  # mm (+ = bas)
    bbox_o: tuple = None
    bbox_m: tuple = None
    details: str = ""

    def xy(self, which):
        """Position (X, Y) en mm du coin haut-gauche : 'o' original, 'm' modifié (recalé)."""
        b = self.bbox_o if which == "o" else self.bbox_m
        return None if b is None else (b[0] / PT_PER_MM, b[1] / PT_PER_MM)


def cluster_rects(rects, gap=1.5):
    """Regroupe les rectangles qui se touchent (union-find + grille)."""
    n = len(rects)
    parent = list(range(n))

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    cell, grid = 40.0, {}
    for i, r in enumerate(rects):
        x0, y0, x1, y1 = r[0] - gap, r[1] - gap, r[2] + gap, r[3] + gap
        cx0, cx1 = int(x0 // cell), int(x1 // cell)
        cy0, cy1 = int(y0 // cell), int(y1 // cell)
        if (cx1 - cx0 + 1) * (cy1 - cy0 + 1) > 400:
            continue  # très grand objet : reste seul
        for cx in range(cx0, cx1 + 1):
            for cy in range(cy0, cy1 + 1):
                lst = grid.setdefault((cx, cy), [])
                for j in lst:
                    q = rects[j]
                    if not (x1 < q[0] - gap or x0 > q[2] + gap or y1 < q[1] - gap or y0 > q[3] + gap):
                        parent[find(i)] = find(j)
                lst.append(i)
    groups = {}
    for i in range(n):
        groups.setdefault(find(i), []).append(i)
    return list(groups.values())


def extract_objects(page):
    """Extrait les objets de packaging : texte, images, vecteurs/logos, codes-barres."""
    objs = []
    try:
        d = page.get_text("dict", flags=fitz.TEXT_PRESERVE_IMAGES |
                          fitz.TEXT_PRESERVE_LIGATURES | fitz.TEXT_PRESERVE_WHITESPACE)
    except Exception:
        d = {"blocks": []}
    for b in d["blocks"]:
        bb = tuple(b["bbox"])
        if b["type"] == 0:
            txt = " ".join(s["text"] for l in b["lines"] for s in l["spans"]).strip()
            if txt:
                sp = max((sx for l in b["lines"] for sx in l["spans"]), key=lambda sx: len(sx["text"]))
                fn = re.sub(r"^[A-Z]{6}\+", "", sp.get("font", ""))
                low, fl = fn.lower(), sp.get("flags", 0)
                style = ("Bold" if (fl & 16 or "bold" in low) else "") + \
                        ("Italic" if (fl & 2 or "italic" in low or "oblique" in low) else "")
                attrs = {"font": fn, "style": style or "Regular", "size": round(sp.get("size", 0), 1),
                         "family": re.sub(r"(psmt|mt|ps)$", "", re.split(r"[-,]", low)[0]),
                         "color": int(sp.get("color", 0))}
                objs.append(Obj("texte", bb, " ".join(txt.lower().split()), txt[:40], attrs))
        elif b["type"] == 1:
            img = b.get("image", b"")
            key = hashlib.md5(img).hexdigest()[:12] if img else f"{b.get('width')}x{b.get('height')}"
            objs.append(Obj("image", bb, key, f"Image {b.get('width')}x{b.get('height')} px"))
    # --- vecteurs : regroupement des tracés voisins ---
    try:
        drawings = page.get_drawings()[:MAX_DRAWINGS]
    except Exception:
        drawings = []
    pr = page.rect
    page_area = max(pr.width * pr.height, 1)
    rects = []
    for dr in drawings:
        r = dr["rect"]
        if r.width * r.height > 0.4 * page_area:
            continue  # fond / filet de coupe
        rects.append((r.x0, r.y0, r.x1, r.y1))
    for g in cluster_rects(rects):
        x0 = min(rects[i][0] for i in g); y0 = min(rects[i][1] for i in g)
        x1 = max(rects[i][2] for i in g); y1 = max(rects[i][3] for i in g)
        w, h, n = x1 - x0, y1 - y0, len(g)
        if max(w, h) < 2:
            continue
        # code-barres : beaucoup de barres fines et hautes
        bars = sum(1 for i in g if (rects[i][2] - rects[i][0]) <= 3 and
                   (rects[i][3] - rects[i][1]) >= 3 * max(rects[i][2] - rects[i][0], 0.1))
        smalls = sum(1 for i in g if max(rects[i][2] - rects[i][0], rects[i][3] - rects[i][1]) <= max(w, h) / 6)
        if n >= 30 and 0.6 <= w / max(h, 0.1) <= 1.6 and smalls >= 0.8 * n:
            kind, label = "code-2D", f"Code 2D ({n} modules)"   # QR / DataMatrix (modules vectoriels)
        elif n >= 12 and bars >= 0.7 * n:
            kind, label = "code-barres", f"Code-barres ({n} barres)"
        elif n >= 3 and w * h / PT_PER_MM ** 2 >= 64:
            kind, label = "logo", f"Logo/vecteur {w / PT_PER_MM:.0f}x{h / PT_PER_MM:.0f} mm"
        else:
            kind, label = "vecteur", f"Vecteur {w / PT_PER_MM:.1f}x{h / PT_PER_MM:.1f} mm"
        key = f"{round(w * 2) / 2}x{round(h * 2) / 2}:{n}"
        objs.append(Obj(kind, (x0, y0, x1, y1), key, label))
    return objs


# ============================================================================
# ANALYSE DU CONTENU PDF (pikepdf) : recettes couleur, overprint, knockout, trapping
# ============================================================================
MAX_ITEMS = 300000
I6 = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)
CS_N = {"DeviceGray": 1, "CalGray": 1, "DeviceRGB": 3, "CalRGB": 3, "Lab": 3, "DeviceCMYK": 4}
PROC_MAP = {"cyan": "C", "magenta": "M", "yellow": "Y", "black": "K"}
DEV_GRAY = {"t": "DeviceGray", "n": 1, "names": []}
DEV_RGB = {"t": "DeviceRGB", "n": 3, "names": []}
DEV_CMYK = {"t": "DeviceCMYK", "n": 4, "names": []}


def n2(t, k):
    """Formate une coordonnée (mm) ou vide."""
    return "" if t is None else f"{t[k]:.2f}"


def mmul(m, n):
    """Produit de matrices PDF : applique m puis n."""
    a, b, c, d, e, f = m
    A, B, C, D, E, F = n
    return (a * A + b * C, a * B + b * D, c * A + d * C, c * B + d * D, e * A + f * C + E, e * B + f * D + F)


def mpt(m, x, y):
    return (m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5])


def cmyk_rgb(c, m, y, k):
    return (int(255 * (1 - c) * (1 - k)), int(255 * (1 - m) * (1 - k)), int(255 * (1 - y) * (1 - k)))


def parse_cs(cs):
    """Descripteur d'espace colorimétrique : {t, n, names[, rgb]}."""
    try:
        if isinstance(cs, pikepdf.Name):
            t = str(cs)[1:]
            return {"t": t, "n": CS_N.get(t, 0), "names": []}
        a = list(cs)
        t = str(a[0])[1:]
        if t == "Separation":
            d = {"t": t, "n": 1, "names": [pdf_name_decode(str(a[1]))]}
            try:   # couleur écran : valeur alternate à 100 % (fonction type 2)
                fn = a[3]
                if int(fn.get("/FunctionType", -1)) == 2:
                    c1 = [float(x) for x in fn.get("/C1", [])]
                    if len(c1) == 4:
                        d["rgb"] = cmyk_rgb(*c1)
                    elif len(c1) == 3:
                        d["rgb"] = tuple(int(255 * x) for x in c1)
            except Exception:
                pass
            return d
        if t == "DeviceN":
            names = [pdf_name_decode(str(x)) for x in a[1]]
            return {"t": t, "n": len(names), "names": names}
        if t == "ICCBased":
            return {"t": t, "n": int(a[1].get("/N", 3)), "names": []}
        return {"t": t, "n": 1 if t == "Indexed" else CS_N.get(t, 0), "names": []}
    except Exception:
        return {"t": "Unknown", "n": 0, "names": []}


def overprintable(cs):
    """L'overprint n'agit que sur CMYK / Separation / DeviceN (ignoré en RGB et Gray)."""
    return cs["t"] in ("DeviceCMYK", "Separation", "DeviceN") or (cs["t"] == "ICCBased" and cs["n"] == 4)


@dataclass
class Recipe:
    cs: str
    comps: dict          # colorant -> % (C,M,Y,K ou nom de spot) ; R,G,B en 0-255

    def text(self):
        c = self.comps
        if "R" in c:
            return f"RGB({c['R']:.0f},{c['G']:.0f},{c['B']:.0f})"
        proc = [f"{k}={v:.0f}" for k, v in c.items() if k in ("C", "M", "Y", "K")]
        spots = [f"{k} {v:.0f}%" for k, v in c.items() if k not in ("C", "M", "Y", "K")]
        if self.cs == "DeviceGray" and proc:
            return f"Gray {c['K']:.0f}% (K)"
        parts = spots + ([" ".join(proc)] if proc else [])
        return " + ".join(parts) if parts else self.cs

    def is_white(self):
        c = self.comps
        if "R" in c:
            return min(c.values()) >= 254
        return bool(c) and max(c.values()) < 0.5


def make_recipe(cs, comps):
    t, n = cs["t"], cs["n"]
    c = [float(x) for x in comps]
    if t in ("DeviceGray", "CalGray") or (t == "ICCBased" and n == 1):
        return Recipe(t, {"K": (1 - (c[0] if c else 0.0)) * 100})
    if t in ("DeviceRGB", "CalRGB") or (t == "ICCBased" and n == 3):
        d = (c + [0, 0, 0])[:3]
        return Recipe(t, {"R": d[0] * 255, "G": d[1] * 255, "B": d[2] * 255})
    if t == "DeviceCMYK" or (t == "ICCBased" and n == 4):
        d = (c + [0] * 4)[:4]
        return Recipe(t, dict(zip("CMYK", [x * 100 for x in d])))
    if t == "Separation":
        return Recipe(t, {cs["names"][0]: (c[0] if c else 1.0) * 100})
    if t == "DeviceN":
        return Recipe(t, {PROC_MAP.get(nm.lower(), nm): v * 100 for nm, v in zip(cs["names"], c)})
    return Recipe(t, {})


@dataclass
class GItem:
    """Un objet peint du PDF (texte, tracé, image, dégradé) avec son état graphique."""
    idx: int
    kind: str
    bbox: tuple                    # pt, repère fitz (origine en haut à gauche)
    fill: object = None
    stroke: object = None
    op_fill: bool = False          # overprint effectif (rempli)
    op_stroke: bool = False
    op_flag: bool = False          # drapeau brut du PDF (peut être sans effet en RGB/Gray)
    opm: int = 0
    blend: str = "Normal"
    ko_group: bool = False         # groupe de transparence knockout
    font: str = ""
    size: float = 0.0
    lw: float = 0.0                # largeur de contour (pt)
    label: str = ""
    category: str = "objet"        # texte / fond / objet

    def knock_fill(self):
        return self.fill is not None and not self.op_fill and self.blend in ("Normal", "Compatible")

    def name(self):
        if self.kind == "text":
            return f"Texte « {self.label} »" if self.label else f"Texte ({self.font} {self.size:.1f} pt)"
        w = (self.bbox[2] - self.bbox[0]) / PT_PER_MM
        h = (self.bbox[3] - self.bbox[1]) / PT_PER_MM
        base = {"image": "Image", "shading": "Dégradé"}.get(self.kind, "Tracé")
        return ("Fond " if self.category == "fond" else base + " ") + f"{w:.1f}x{h:.1f} mm #{self.idx}"


class ContentAnalyzer:
    """Interprète les flux de contenu (pikepdf) en suivant l'état graphique :
    couleurs, overprint (ExtGState /OP /op /OPM), groupes knockout, contours (trapping)."""

    def __init__(self, pdf, fpage):
        self.pdf = pdf
        self.items, self.trap_candidates = [], []
        self.used_spots, self.spot_rgb, self.info = set(), {}, {}
        self.spot_kinds = {}      # nom -> Separation / DeviceN (spots réellement peints)
        r = fpage.rect
        self.page_area = max(r.width * r.height, 1.0)
        self.base = tuple(fpage.transformation_matrix)   # PDF -> repère fitz (rotation incluse)

    def run(self, page):
        try:
            res = page.resources
        except Exception:
            res = page.obj.get("/Resources")
        st = dict(ctm=self.base, fcs=DEV_GRAY, fc=[0.0], scs=DEV_GRAY, sc=[0.0], op=False, OP=False,
                  opm=0, lw=1.0, bm="Normal", font="", fsize=0.0, tc=0.0, tw=0.0, tz=1.0, tl=0.0,
                  ts=0.0, tr=0, tm=I6, tlm=I6, ko=False)
        self._stream(page, res, st, 0)
        try:
            self.info["Trapped"] = str(self.pdf.docinfo.get("/Trapped", "non défini"))
        except Exception:
            self.info["Trapped"] = "non défini"
        try:
            self.info["TrapNet"] = sum(1 for a in page.obj.get("/Annots", []) if str(a.get("/Subtype")) == "/TrapNet")
            self.info["TrapRegions"] = "/TrapRegions" in page.obj
        except Exception:
            self.info["TrapNet"], self.info["TrapRegions"] = 0, False

    # ---- helpers ----
    def _cs(self, res, name):
        if name in CS_N or name == "Pattern":
            return {"t": name, "n": CS_N.get(name, 0), "names": []}
        try:
            cs = parse_cs(res.get("/ColorSpace").get("/" + name))
        except Exception:
            return {"t": "Unknown", "n": 0, "names": []}
        if "rgb" in cs:
            self.spot_rgb[cs["names"][0]] = cs["rgb"]
        return cs

    @staticmethod
    def _default(cs):
        if cs["t"] in ("Separation", "DeviceN"):
            return [1.0] * cs["n"]
        if cs["t"] == "DeviceCMYK" or (cs["t"] == "ICCBased" and cs["n"] == 4):
            return [0.0, 0.0, 0.0, 1.0]
        return [0.0] * max(cs["n"], 1)

    @staticmethod
    def _gs(st, g):
        if "/OP" in g:
            st["OP"] = bool(g.get("/OP"))
            if "/op" not in g:
                st["op"] = st["OP"]
        if "/op" in g:
            st["op"] = bool(g.get("/op"))
        if "/OPM" in g:
            st["opm"] = int(g.get("/OPM"))
        if "/LW" in g:
            st["lw"] = float(g.get("/LW"))
        if "/BM" in g:
            bm = g.get("/BM")
            if isinstance(bm, pikepdf.Array):
                bm = bm[0]
            st["bm"] = str(bm)[1:]

    @staticmethod
    def _bb(pts, pad=0.0):
        xs, ys = [p[0] for p in pts], [p[1] for p in pts]
        return (min(xs) - pad, min(ys) - pad, max(xs) + pad, max(ys) + pad)

    def _emit(self, kind, bbox, st, fill, stroke, fcs=None, fixed=None, label="", size=0.0, lw=0.0):
        if len(self.items) >= MAX_ITEMS:
            return
        fcs = fcs or st["fcs"]
        fr = (fixed or make_recipe(fcs, st["fc"])) if fill else None
        sr = make_recipe(st["scs"], st["sc"]) if stroke else None
        for cs, on in ((fcs, fill), (st["scs"], stroke)):
            if on and cs["t"] in ("Separation", "DeviceN"):
                for n in cs["names"]:
                    if n.lower() not in PROC_MAP and n.lower() not in ("none", "all"):
                        self.used_spots.add(n)
                        self.spot_kinds[n] = cs["t"]
        w, h = bbox[2] - bbox[0], bbox[3] - bbox[1]
        cat = "texte" if kind == "text" else ("fond" if kind in ("path", "image") and w * h >= 0.5 * self.page_area else "objet")
        it = GItem(len(self.items), kind, bbox, fr, sr,
                   bool(fill and st["op"] and overprintable(fcs)),
                   bool(stroke and st["OP"] and overprintable(st["scs"])),
                   bool(st["op"] if fill else st["OP"]), st["opm"], st["bm"], st["ko"],
                   st["font"], size, lw, label, cat)
        self.items.append(it)
        if stroke and it.op_stroke and 0 < lw <= 1.5 * PT_PER_MM:   # trap probable
            self.trap_candidates.append(it)

    def _lw(self, st):
        c = st["ctm"]
        return st["lw"] * math.sqrt(abs(c[0] * c[3] - c[1] * c[2]))

    def _paint(self, st, path, fill, stroke):
        if path:
            lw = self._lw(st) if stroke else 0.0
            self._emit("path", self._bb(path, lw / 2), st, fill, stroke, lw=lw)

    def _image(self, st, xo):
        pts = [mpt(st["ctm"], x, y) for x, y in ((0, 0), (1, 0), (1, 1), (0, 1))]
        cs = DEV_GRAY
        try:
            if bool(xo.get("/ImageMask", False)):
                cs = st["fcs"]
            elif "/ColorSpace" in xo:
                cs = parse_cs(xo.get("/ColorSpace"))
        except Exception:
            pass
        self._emit("image", self._bb(pts), st, True, False, fcs=cs, fixed=Recipe("Image " + cs["t"], {}))

    def _show(self, st, nch, adj, label):
        size, tz = st["fsize"], st["tz"]
        adv = (nch * 0.5 * size + nch * st["tc"]) * tz - adj / 1000.0 * size * tz
        tr = st["tr"]
        if nch and size > 0 and tr not in (3, 7):
            m = mmul(st["tm"], st["ctm"])
            y0, y1 = -0.2 * size + st["ts"], 0.8 * size + st["ts"]
            pts = [mpt(m, x, y) for x, y in ((0, y0), (adv, y0), (adv, y1), (0, y1))]
            eff = size * math.sqrt(abs(m[0] * m[3] - m[1] * m[2]))
            self._emit("text", self._bb(pts), st, tr in (0, 2, 4, 6), tr in (1, 2, 5, 6),
                       label=label, size=eff, lw=self._lw(st))
        st["tm"] = mmul((1, 0, 0, 1, adv, 0), st["tm"])

    # ---- interpréteur ----
    def _stream(self, obj, res, st, depth):
        try:
            ins = pikepdf.parse_content_stream(obj)
        except Exception:
            return
        stack, path = [], []
        for i in ins:
            try:
                op, a = str(i.operator), i.operands
            except Exception:
                continue
            try:
                if op == "q":
                    stack.append(dict(st))
                elif op == "Q":
                    if stack:
                        st.clear(); st.update(stack.pop())
                elif op == "cm":
                    st["ctm"] = mmul(tuple(float(x) for x in a), st["ctm"])
                elif op == "re":
                    x, y, w, h = (float(v) for v in a)
                    path += [mpt(st["ctm"], x, y), mpt(st["ctm"], x + w, y),
                             mpt(st["ctm"], x + w, y + h), mpt(st["ctm"], x, y + h)]
                elif op in ("m", "l"):
                    path.append(mpt(st["ctm"], float(a[0]), float(a[1])))
                elif op == "c":
                    path += [mpt(st["ctm"], float(a[k]), float(a[k + 1])) for k in (0, 2, 4)]
                elif op in ("v", "y"):
                    path += [mpt(st["ctm"], float(a[k]), float(a[k + 1])) for k in (0, 2)]
                elif op in ("f", "F", "f*"):
                    self._paint(st, path, True, False); path = []
                elif op in ("S", "s"):
                    self._paint(st, path, False, True); path = []
                elif op in ("B", "B*", "b", "b*"):
                    self._paint(st, path, True, True); path = []
                elif op == "n":
                    path = []
                elif op == "w":
                    st["lw"] = float(a[0])
                elif op == "gs":
                    g = res.get("/ExtGState").get("/" + str(a[0])[1:])
                    if g is not None:
                        self._gs(st, g)
                elif op in ("cs", "CS"):
                    k = "f" if op == "cs" else "s"
                    st[k + "cs"] = self._cs(res, str(a[0])[1:])
                    st[k + "c"] = self._default(st[k + "cs"])
                elif op in ("sc", "scn", "SC", "SCN"):
                    comps = []
                    for x in a:
                        try:
                            comps.append(float(x))
                        except Exception:
                            pass
                    st["fc" if op in ("sc", "scn") else "sc"] = comps
                elif op in ("g", "G"):
                    k = "f" if op == "g" else "s"
                    st[k + "cs"], st[k + "c"] = DEV_GRAY, [float(a[0])]
                elif op in ("rg", "RG"):
                    k = "f" if op == "rg" else "s"
                    st[k + "cs"], st[k + "c"] = DEV_RGB, [float(x) for x in a]
                elif op in ("k", "K"):
                    k = "f" if op == "k" else "s"
                    st[k + "cs"], st[k + "c"] = DEV_CMYK, [float(x) for x in a]
                elif op == "BT":
                    st["tm"] = st["tlm"] = I6
                elif op == "Tf":
                    st["fsize"] = float(a[1])
                    try:
                        fo = res.get("/Font").get("/" + str(a[0])[1:])
                        st["font"] = re.sub(r"^[A-Z]{6}\+", "", str(fo.get("/BaseFont", ""))[1:])
                    except Exception:
                        st["font"] = ""
                elif op in ("Td", "TD"):
                    if op == "TD":
                        st["tl"] = -float(a[1])
                    st["tlm"] = mmul((1, 0, 0, 1, float(a[0]), float(a[1])), st["tlm"]); st["tm"] = st["tlm"]
                elif op == "Tm":
                    st["tm"] = st["tlm"] = tuple(float(x) for x in a)
                elif op == "T*":
                    st["tlm"] = mmul((1, 0, 0, 1, 0, -st["tl"]), st["tlm"]); st["tm"] = st["tlm"]
                elif op == "Tc":
                    st["tc"] = float(a[0])
                elif op == "Tw":
                    st["tw"] = float(a[0])
                elif op == "Tz":
                    st["tz"] = float(a[0]) / 100.0
                elif op == "TL":
                    st["tl"] = float(a[0])
                elif op == "Ts":
                    st["ts"] = float(a[0])
                elif op == "Tr":
                    st["tr"] = int(a[0])
                elif op in ("Tj", "TJ", "'", '"'):
                    if op in ("'", '"'):
                        st["tlm"] = mmul((1, 0, 0, 1, 0, -st["tl"]), st["tlm"]); st["tm"] = st["tlm"]
                    items = list(a[0]) if op == "TJ" else [a[-1]]
                    nch, adj, lab = 0, 0.0, ""
                    for e in items:
                        try:
                            b = bytes(e)
                            nch += len(b)
                            lab += "".join(chr(c) for c in b if 32 <= c < 127)
                        except Exception:
                            try:
                                adj += float(e)
                            except Exception:
                                pass
                    self._show(st, nch, adj, lab.strip()[:40])
                elif op == "Do":
                    xo = res.get("/XObject").get("/" + str(a[0])[1:])
                    sub = str(xo.get("/Subtype"))
                    if sub == "/Image":
                        self._image(st, xo)
                    elif sub == "/Form" and depth < 8:
                        st2 = dict(st)
                        try:
                            st2["ctm"] = mmul(tuple(float(x) for x in xo.get("/Matrix", I6)), st["ctm"])
                        except Exception:
                            pass
                        grp = xo.get("/Group")
                        if grp is not None and "/K" in grp:
                            st2["ko"] = bool(grp.get("/K"))
                        self._stream(xo, xo.get("/Resources", res), st2, depth + 1)
                elif op == "INLINE IMAGE":
                    self._image(st, None)
                elif op == "sh":
                    self._emit("shading", self._bb([mpt(st["ctm"], 0, 0)]), st, True, False)
            except Exception:
                continue          # un opérateur illisible ne bloque jamais l'analyse


def inspector_rows(an):
    """Lignes (objet, catégorie, couleur, état, remarque) pour Overprint et Knockout."""
    op_rows, ko_rows = [], []
    for it in an.items:
        col = it.fill.text() if it.fill else (it.stroke.text() if it.stroke else "—")
        on_op = it.op_fill if it.fill else it.op_stroke
        rem = ""
        if it.op_flag and not on_op:
            rem = "flag ignoré (RGB/Gray)"
        elif on_op and it.fill is not None and it.fill.is_white():
            rem = "overprint blanc : invisible"
        op_rows.append((it.name(), it.category, col, "ON" if on_op else "OFF", rem))
        ko = it.knock_fill() if it.fill else (it.stroke is not None and not it.op_stroke)
        ko_rows.append((it.name(), it.category, col, "ON" if ko else "OFF", "groupe knockout" if it.ko_group else ""))
    return op_rows, ko_rows


def trap_report(an):
    i = an.info
    L = [f"Métadonnée /Trapped : {i.get('Trapped')}",
         f"Annotations TrapNet (page) : {i.get('TrapNet', 0)}",
         f"/TrapRegions (page) : {'oui' if i.get('TrapRegions') else 'non'}",
         f"Contours en overprint (traps probables) : {len(an.trap_candidates)}"]
    for it in an.trap_candidates[:200]:
        L.append(f"  • largeur {it.lw / PT_PER_MM:.3f} mm · couleur {it.stroke.text()} · {it.name()}")
    if not an.trap_candidates and not i.get("TrapNet") and "True" not in str(i.get("Trapped")):
        L.append("Aucun trapping détecté dans le fichier.")
    L.append("Note : largeur/couleur lues sur les contours en overprint (heuristique). "
             "Un trapping calculé par un RIP n'est pas lisible dans le PDF.")
    return "\n".join(L)


def items_at(an, x, y, tol=1.0):
    return [it for it in an.items if it.bbox[0] - tol <= x <= it.bbox[2] + tol and it.bbox[1] - tol <= y <= it.bbox[3] + tol]


def effective_recipe(stack):
    """Recette réellement imprimée : knockout efface le dessous, overprint s'y ajoute."""
    res = {}
    for it in stack:
        r = it.fill
        if r is None or not r.comps:
            continue
        comps = dict(r.comps)
        if it.knock_fill():
            res = {}
        elif it.opm == 1 and "C" in comps:       # OPM 1 : seules les composantes non nulles écrasent
            comps = {k: v for k, v in comps.items() if v > 0}
        res.update(comps)
    return res


def recipe_report(an, x, y):
    st = sorted(items_at(an, x, y), key=lambda i: i.idx)
    if not st:
        return "Color Recipe\nAucun objet à cet endroit."
    top = st[-1]
    onoff = lambda v: "ON" if v else "OFF"
    L = ["Color Recipe", f"Objet : {top.name()} ({top.kind})"]
    if top.fill:
        L.append(f"Remplissage : {top.fill.text()}   [Overprint {onoff(top.op_fill)} · Knockout {onoff(top.knock_fill())}]")
    if top.stroke:
        L.append(f"Contour : {top.stroke.text()} ({top.lw / PT_PER_MM:.2f} mm)   [Overprint {onoff(top.op_stroke)}]")
    if top.kind == "text":
        L.append(f"Police : {top.font or '?'} · corps {top.size:.1f} pt")
    eff = effective_recipe(st)
    if eff:
        L.append(f"Recette effective (empilement + overprint) : {Recipe('eff', eff).text()}")
    if len(st) > 1:
        L.append("Objets sous le curseur (haut → bas) :")
        for it in reversed(st[:-1][-8:]):
            L.append(f"  • {it.name()} — {it.fill.text() if it.fill else '—'}")
    L.append("(zones de sélection = boîtes englobantes ; largeur du texte estimée)")
    return "\n".join(L)


# ============================================================================
# CODES-BARRES : décodage et comparaison
# ============================================================================
def decode_code(page, bbox, scale=6.0):
    """Rend la zone du code à haute résolution (depuis le PDF) et tente de le décoder."""
    r = fitz.Rect(*bbox) + (-6, -6, 6, 6)
    pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), clip=r, colorspace=fitz.csGRAY, alpha=False)
    g = np.frombuffer(pix.samples, np.uint8).reshape(pix.h, pix.w).copy()
    g = cv2.copyMakeBorder(g, 40, 40, 40, 40, cv2.BORDER_CONSTANT, value=255)
    typ = text = None
    try:
        from pyzbar.pyzbar import decode as zdec          # EAN, Code128, QR...
        res = zdec(g)
        if res:
            typ, text = res[0].type, res[0].data.decode("utf-8", "replace")
    except Exception:
        pass
    if text is None:
        try:
            from pylibdmtx.pylibdmtx import decode as ddec   # DataMatrix
            res = ddec(g, timeout=500)
            if res:
                typ, text = "DATAMATRIX", res[0].data.decode("utf-8", "replace")
        except Exception:
            pass
    if text is None:
        try:
            t, _, _ = cv2.QRCodeDetector().detectAndDecode(g)
            if t:
                typ, text = "QRCODE", t
        except Exception:
            pass
    if text is None:
        try:
            ok, infos, types, _ = cv2.barcode.BarcodeDetector().detectAndDecodeWithType(g)
            if ok and infos and infos[0]:
                typ, text = str(types[0]), infos[0]
        except Exception:
            pass
    return {"type": typ, "text": text, "img": g}


def crops_differ(a, b):
    """Comparaison visuelle quand le décodage est impossible."""
    a = cv2.resize(a, (200, 200), interpolation=cv2.INTER_AREA)
    b = cv2.resize(b, (200, 200), interpolation=cv2.INTER_AREA)
    return float(np.mean(cv2.absdiff(a, b) > 100)) > 0.03


def code_kind(c, default):
    t = ((c or {}).get("type") or "").upper()
    if "QR" in t:
        return "QR"
    if "DATAMATRIX" in t or "DMTX" in t:
        return "DataMatrix"
    return default


def compare_codes(lo, lm, to_orig, dec_o, dec_m, out):
    """Codes-barres / QR / DataMatrix : appariement par proximité puis contrôle du contenu."""
    free = set(range(len(lm)))
    for o in lo:
        oc = center(o.bbox)
        best = min(free, key=lambda j: math.dist(oc, center(to_orig(lm[j].bbox))), default=None)
        if best is None or math.dist(oc, center(to_orig(lm[best].bbox))) > 40 * PT_PER_MM:
            out.append(Finding(o.kind, "supprimé", "CRITICAL", o.label, bbox_o=o.bbox))
            continue
        free.discard(best)
        m = lm[best]
        mb = to_orig(m.bbox)
        co = dec_o(o.bbox) if dec_o else None
        cm = dec_m(m.bbox) if dec_m else None
        kind = code_kind(co, o.kind)
        label = f"{kind} {co['text']}" if co and co["text"] else o.label
        dx, dy = (center(mb)[0] - oc[0]) / PT_PER_MM, (center(mb)[1] - oc[1]) / PT_PER_MM
        changed, det = False, ""
        if co and cm and co["text"] and cm["text"]:
            if co["text"] != cm["text"]:
                changed, det = True, f"Contenu : {co['text']} → {cm['text']}"
        elif co and cm and crops_differ(co["img"], cm["img"]):
            changed, det = True, "Contenu visuel différent (décodage impossible)"
        if changed:
            out.append(Finding(kind, "modifié", "CRITICAL", label, dx, dy, o.bbox, mb, det))
        dw = ((mb[2] - mb[0]) - (o.bbox[2] - o.bbox[0])) / PT_PER_MM
        dh = ((mb[3] - mb[1]) - (o.bbox[3] - o.bbox[1])) / PT_PER_MM
        if max(abs(dw), abs(dh)) >= IDENTICAL_MM and grade(dw, dh) != "OK":
            out.append(Finding(kind, "modifié", grade(dw, dh), label, dx, dy, o.bbox, mb,
                               f"Taille : Δl {dw:+.2f} mm, Δh {dh:+.2f} mm"))
        if max(abs(dx), abs(dy)) >= IDENTICAL_MM:
            out.append(Finding(kind, "déplacé", grade(dx, dy), label, dx, dy, o.bbox, mb))
        elif not changed:
            out.append(Finding(kind, "identique", "OK", label, dx, dy, o.bbox, mb))
    for j in free:
        out.append(Finding(lm[j].kind, "ajouté", "CRITICAL", lm[j].label, bbox_m=to_orig(lm[j].bbox)))


def style_diff(a, b):
    """Changements de police / style / corps / couleur entre deux textes identiques."""
    out = []
    if not a or not b:
        return out
    if a["family"] != b["family"]:
        out.append(f"Police : {a['font']} → {b['font']}")
    elif a["style"] != b["style"]:
        out.append(f"Style : {a['style']} → {b['style']}")
    if abs(a["size"] - b["size"]) > 0.05:
        out.append(f"Corps : {a['size']} → {b['size']} pt")
    if a["color"] != b["color"]:
        out.append(f"Couleur : #{a['color']:06X} → #{b['color']:06X}")
    return out


def compare_objects(objs_o, objs_m, to_orig, dec_o=None, dec_m=None):
    """Apparie les objets original/modifié : déplacés (ΔX/ΔY mm), ajoutés, supprimés,
    textes modifiés (contenu, police, style, corps, couleur), codes-barres.
    to_orig : ramène une bbox du modifié dans le repère de l'original
    (tient compte du déplacement manuel + auto-align)."""
    out = []
    for kind in sorted({o.kind for o in objs_o} | {m.kind for m in objs_m}):
        lo = [o for o in objs_o if o.kind == kind]
        lm = [m for m in objs_m if m.kind == kind]
        if kind in ("code-barres", "code-2D"):
            compare_codes(lo, lm, to_orig, dec_o, dec_m, out)
            continue
        buckets = {}
        for j, m in enumerate(lm):
            buckets.setdefault(m.key, []).append(j)
        used, un_o = set(), []
        for o in lo:
            cands = [j for j in buckets.get(o.key, []) if j not in used]
            if not cands:
                un_o.append(o)
                continue
            oc = center(o.bbox)
            j = min(cands, key=lambda j: math.dist(oc, center(to_orig(lm[j].bbox))))
            used.add(j)
            mb = to_orig(lm[j].bbox)
            dx, dy = (center(mb)[0] - oc[0]) / PT_PER_MM, (center(mb)[1] - oc[1]) / PT_PER_MM
            if max(abs(dx), abs(dy)) < IDENTICAL_MM:
                out.append(Finding(kind, "identique", "OK", o.label, dx, dy, o.bbox, mb))
            else:
                out.append(Finding(kind, "déplacé", grade(dx, dy), o.label, dx, dy, o.bbox, mb))
            if kind == "texte":
                for msg in style_diff(o.attrs, lm[j].attrs):
                    out.append(Finding(kind, "modifié", "CRITICAL", o.label, dx, dy, o.bbox, mb, msg))
        un_m = [j for j in range(len(lm)) if j not in used]
        if kind == "texte":  # même zone, contenu différent -> texte modifié
            for o in list(un_o):
                for j in list(un_m):
                    mb = to_orig(lm[j].bbox)
                    ix = max(0, min(o.bbox[2], mb[2]) - max(o.bbox[0], mb[0]))
                    iy = max(0, min(o.bbox[3], mb[3]) - max(o.bbox[1], mb[1]))
                    area = max((o.bbox[2] - o.bbox[0]) * (o.bbox[3] - o.bbox[1]), 1e-6)
                    if ix * iy / area > 0.5:
                        det = "; ".join(["Contenu modifié"] + style_diff(o.attrs, lm[j].attrs))
                        out.append(Finding(kind, "modifié", "CRITICAL", f"{o.label} -> {lm[j].label}",
                                           0, 0, o.bbox, mb, det))
                        un_o.remove(o); un_m.remove(j)
                        break
        for o in un_o:
            out.append(Finding(kind, "supprimé", "CRITICAL", o.label, 0, 0, o.bbox, None))
        for j in un_m:
            out.append(Finding(kind, "ajouté", "CRITICAL", lm[j].label, 0, 0, None, to_orig(lm[j].bbox)))
    return out


class PdfDoc:
    """Un PDF ouvert : rendu RGB, plaques de séparation, détection des couleurs."""

    def __init__(self, path):
        self.path = path
        self.doc = fitz.open(path)
        self.page_count = len(self.doc)
        self.spots = {}        # nom -> type (PMS, White Ink, Varnish...)
        self.spot_cs = {}      # nom -> "Separation" / "DeviceN"
        self.sep_xrefs = []    # objets contenant un espace Separation
        self.ocg_off = set()   # calques masqués
        self._rgb, self._planes, self._objects = {}, {}, {}
        self._present, self._an, self._pk = {}, {}, None   # encres présentes, analyses pikepdf
        self.spot_rgb = {}
        self._scan_colorspaces()
        self.layers = self._read_layers()

    # ---------- détection native : /DeviceCMYK /Separation /DeviceN ----------
    def _scan_colorspaces(self):
        skip = {"cyan", "magenta", "yellow", "black", "none", "all"}
        for x in range(1, self.doc.xref_length()):
            try:
                s = self.doc.xref_object(x, compressed=True)
            except Exception:
                continue
            if "/Separation" in s or "/DeviceN" in s:
                self.sep_xrefs.append(x)
            if "/Separation" in s:
                for m in SEP_RE.finditer(s):
                    n = pdf_name_decode(m.group(1))
                    if n.lower() not in skip:
                        self.spots[n] = classify_spot(n)
                        self.spot_cs[n] = "Separation"
            if "/DeviceN" in s:
                for m in re.finditer(r"/DeviceN\s*\[([^\]]*)\]", s):
                    for n in re.findall(r"/([^\s/\[\]<>()]+)", m.group(1)):
                        n = pdf_name_decode(n)
                        if n.lower() not in skip:
                            self.spots.setdefault(n, classify_spot(n))
                            self.spot_cs.setdefault(n, "DeviceN")

    # ---------- calques (OCG) ----------
    def _read_layers(self):
        out = {}
        try:
            for xref, info in self.doc.get_ocgs().items():
                out.setdefault(info.get("name", f"Layer {xref}"), []).append(xref)
        except Exception:
            pass
        return out

    def set_layer(self, name, on):
        for x in self.layers.get(name, []):
            (self.ocg_off.discard if on else self.ocg_off.add)(x)
        self._rgb.clear(); self._planes.clear()

    def _apply_layers(self, d):
        if self.ocg_off:
            try:
                d.set_layer(-1, off=list(self.ocg_off))
            except Exception:
                pass

    # ---------- rendu ----------
    def scale_for(self, pno):
        r = self.doc[pno].rect
        return min(RENDER_SCALE, MAX_SIDE_PX / max(r.width, r.height, 1))

    @staticmethod
    def _put(cache, key, val):
        if len(cache) >= 3:             # mémoire bornée : 3 pages max
            cache.pop(next(iter(cache)))
        cache[key] = val

    def render_rgb(self, pno, scale):
        key = (pno, round(scale, 4))
        if key not in self._rgb:
            d = self.doc
            self._apply_layers(d)
            pix = d[pno].get_pixmap(matrix=fitz.Matrix(scale, scale),
                                    colorspace=fitz.csRGB, alpha=False)
            arr = np.frombuffer(pix.samples, np.uint8).reshape(pix.h, pix.w, 3).copy()
            self._put(self._rgb, key, arr)
        return self._rgb[key]

    def _variant(self, pno, target):
        """Copie mémoire du PDF où chaque /Separation est blanc, sauf `target` (noir).
        target=None -> toutes les spots en blanc (plaques CMYK pures)."""
        d = fitz.open(self.path)
        self._apply_layers(d)

        def repl(m):
            on = target is not None and pdf_name_decode(m.group(1)) == target
            return (f"[/Separation {m.group(1)} /DeviceGray << /FunctionType 2 "
                    f"/Domain [0 1] /C0 [1] /C1 [{0 if on else 1}] /N 1 >>]")

        fns = {}

        def mk_fn(n, idx):
            """Fonction PostScript : sélectionne le colorant `idx` d'un DeviceN (ou blanc si None)."""
            if (n, idx) not in fns:
                body = ("pop " * (n - 1 - idx) + "exch pop " * idx + "1 exch sub ") if idx is not None \
                    else "pop " * n + "1 "
                xr = d.get_new_xref()
                d.update_object(xr, "<< /FunctionType 4 /Domain [" + "0 1 " * n + "] /Range [0 1] >>")
                d.update_stream(xr, ("{ " + body + "}").encode(), new=True)
                fns[(n, idx)] = xr
            return fns[(n, idx)]

        def repl_n(m):
            names = [pdf_name_decode(t) for t in re.findall(r"/([^\s/\[\]<>()]+)", m.group(1))]
            idx = names.index(target) if target in names else None
            return f"[/DeviceN [{m.group(1)}] /DeviceGray {mk_fn(len(names), idx)} 0 R]"

        for x in self.sep_xrefs:
            try:
                s = d.xref_object(x, compressed=True)
                n = DEVN_RE.sub(repl_n, SEP_RE.sub(repl, s))
                if n != s:
                    d.update_object(x, n)
            except Exception:
                pass
        if target is not None:  # on masque aussi le CMYK/RGB/Gray du flux
            xs = list(d[pno].get_contents())
            try:
                xs += [t[0] for t in d[pno].get_xobjects()]
            except Exception:
                pass
            for x in xs:
                try:
                    s = d.xref_stream(x)
                    if s:
                        d.update_stream(x, neutralize_process_colors(s))
                except Exception:
                    pass
        return d

    def planes(self, pno, scale):
        """Plaques d'encre uint8 (0 = pas d'encre, 255 = 100 %) : Cyan, Magenta, Yellow, Black + spots."""
        key = (pno, round(scale, 4))
        if key in self._planes:
            return self._planes[key]
        m = fitz.Matrix(scale, scale)
        d = self._variant(pno, None)
        pix = d[pno].get_pixmap(matrix=m, colorspace=fitz.csCMYK, alpha=False)
        a = np.frombuffer(pix.samples, np.uint8).reshape(pix.h, pix.w, pix.n)
        planes = {name: a[:, :, i].copy() for i, name in enumerate(PROCESS)}
        d.close()
        for name in self.page_spots(pno):        # uniquement les spots utilisés sur la page
            d = self._variant(pno, name)
            pix = d[pno].get_pixmap(matrix=m, colorspace=fitz.csGRAY, alpha=False)
            g = np.frombuffer(pix.samples, np.uint8).reshape(pix.h, pix.w)
            planes[name] = 255 - g
            d.close()
        self._put(self._planes, key, planes)
        self._present[key] = {k for k, v in planes.items() if v.max() >= 5}   # encres réellement imprimées
        return planes

    def coverage(self, pno, scale):
        """Couverture % = part de la surface de page portant de l'encre (>= 2 %)."""
        return {k: float((v >= 5).mean() * 100.0) for k, v in self.planes(pno, scale).items()}

    def analysis(self, pno):
        """Analyse native du contenu (pikepdf) : objets, couleurs, overprint, trapping."""
        if pikepdf is None:
            return None
        if pno not in self._an:
            a = None
            try:
                if self._pk is None:
                    self._pk = pikepdf.open(self.path)
                a = ContentAnalyzer(self._pk, self.doc[pno])
                a.run(self._pk.pages[pno])
                self.spot_rgb.update(a.spot_rgb)
                for n, k in a.spot_kinds.items():      # complète la détection par regex
                    self.spots.setdefault(n, classify_spot(n))
                    self.spot_cs.setdefault(n, k)
            except Exception:
                a = None
            self._an[pno] = a
        return self._an[pno]

    def page_spots(self, pno):
        a = self.analysis(pno)
        return list(self.spots) if a is None else [n for n in self.spots if n in a.used_spots]

    def present(self, pno, scale):
        """Noms des encres (process + spots) qui portent réellement de l'encre sur la page."""
        self.planes(pno, scale)
        return self._present[(pno, round(scale, 4))]

    def decode_code(self, pno, bbox):
        return decode_code(self.doc[pno], bbox)

    def close(self):
        try:
            self.doc.close()
        except Exception:
            pass
        if self._pk is not None:
            try:
                self._pk.close()
            except Exception:
                pass

    def objects(self, pno):
        if pno not in self._objects:
            self._objects[pno] = extract_objects(self.doc[pno])
        return self._objects[pno]


@dataclass
class Channel:
    name: str
    kind: str                     # "process" / "spot"
    rgb: tuple
    names: list = field(default_factory=list)   # noms réels dans le PDF
    present: bool = True
    visible: bool = True
    group: str = "CMYK"      # CMYK / PMS / type d'encre spéciale


# ============================================================================
# MOTEURS OPENCV : ALIGNEMENT ET DIFFÉRENCE
# ============================================================================
def estimate_alignment(orig_rgb, mod_rgb):
    """Retourne (M 2x3, infos) qui envoie les pixels du modifié vers l'original.
    Détecte translation + rotation + échelle (ORB + RANSAC, repli ECC affine)."""
    go = cv2.cvtColor(orig_rgb, cv2.COLOR_RGB2GRAY)
    gm = cv2.cvtColor(mod_rgb, cv2.COLOR_RGB2GRAY)
    f = min(1.0, 2000.0 / max(go.shape + gm.shape))
    so = cv2.resize(go, None, fx=f, fy=f, interpolation=cv2.INTER_AREA)
    sm = cv2.resize(gm, None, fx=f, fy=f, interpolation=cv2.INTER_AREA)
    M = None
    # 1) points caractéristiques ORB
    orb = cv2.ORB_create(5000)
    k1, d1 = orb.detectAndCompute(so, None)
    k2, d2 = orb.detectAndCompute(sm, None)
    if d1 is not None and d2 is not None and len(k1) >= 10 and len(k2) >= 10:
        matches = sorted(cv2.BFMatcher(cv2.NORM_HAMMING, crossCheck=True).match(d2, d1),
                         key=lambda m: m.distance)[:500]
        if len(matches) >= 8:
            src = np.float32([k2[m.queryIdx].pt for m in matches]).reshape(-1, 1, 2)
            dst = np.float32([k1[m.trainIdx].pt for m in matches]).reshape(-1, 1, 2)
            M, inl = cv2.estimateAffinePartial2D(src, dst, method=cv2.RANSAC,
                                                 ransacReprojThreshold=3.0)
            if M is not None and int(inl.sum()) < 8:
                M = None
    # 2) repli : corrélation ECC (affine)
    if M is None:
        try:
            H, W = max(so.shape[0], sm.shape[0]), max(so.shape[1], sm.shape[1])
            a = np.full((H, W), 255, np.uint8); a[:so.shape[0], :so.shape[1]] = so
            b = np.full((H, W), 255, np.uint8); b[:sm.shape[0], :sm.shape[1]] = sm
            a, b = cv2.GaussianBlur(a, (5, 5), 0), cv2.GaussianBlur(b, (5, 5), 0)
            warp = np.eye(2, 3, dtype=np.float32)
            crit = (cv2.TERM_CRITERIA_EPS | cv2.TERM_CRITERIA_COUNT, 200, 1e-6)
            _, warp = cv2.findTransformECC(a, b, warp, cv2.MOTION_AFFINE, crit, None, 5)
            M = cv2.invertAffineTransform(warp)   # ECC : original -> modifié
        except cv2.error:
            return None, None
    M = M.astype(np.float64)
    M[:, 2] /= f                                   # retour à l'échelle réelle
    angle = math.degrees(math.atan2(M[1, 0], M[0, 0]))
    scale = math.hypot(M[0, 0], M[1, 0])
    return M, (angle, scale, M[0, 2], M[1, 2])


def compute_difference(o_rgb, m_warp, moved_boxes_px):
    """Image de différence : rouge=supprimé, vert=ajouté, jaune=déplacé."""
    g1 = cv2.cvtColor(o_rgb, cv2.COLOR_RGB2GRAY)
    g2 = cv2.cvtColor(m_warp, cv2.COLOR_RGB2GRAY)
    diff = cv2.absdiff(o_rgb, m_warp).max(axis=2)
    mask = (diff > DIFF_THRESHOLD).astype(np.uint8)
    mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, np.ones((2, 2), np.uint8)) > 0
    removed, added = mask & (g1 < g2), mask & (g2 <= g1)
    moved_zone = np.zeros(mask.shape, bool)
    for (x0, y0, x1, y1) in moved_boxes_px:
        moved_zone[max(0, y0):max(0, y1), max(0, x0):max(0, x1)] = True
    moved = mask & moved_zone
    removed &= ~moved_zone
    added &= ~moved_zone
    out = (o_rgb.astype(np.float32) * 0.25 + 255 * 0.75).astype(np.uint8)  # fond fantôme
    out[removed] = (255, 0, 0)
    out[added] = (0, 170, 0)
    out[moved] = (255, 200, 0)
    for (x0, y0, x1, y1) in moved_boxes_px:
        cv2.rectangle(out, (x0, y0), (x1, y1), (255, 200, 0), 2)
    return out


# ============================================================================
# WIDGETS : ITEM DÉPLAÇABLE, VIEWER, RÈGLES, MINIMAP
# ============================================================================
class MovableItem(QGraphicsPixmapItem):
    """Pixmap déplaçable qui notifie la fenêtre à chaque déplacement."""

    def __init__(self, on_move=None):
        super().__init__()
        self.on_move = on_move
        self.setFlag(QGraphicsItem.ItemSendsGeometryChanges, True)

    def itemChange(self, change, value):
        if change == QGraphicsItem.ItemPositionHasChanged and self.on_move:
            self.on_move()
        return super().itemChange(change, value)


class Viewer(QGraphicsView):
    mouseMoved = Signal(QPointF)
    measured = Signal(QPointF, QPointF)
    picked = Signal(QPointF)
    wheelZoom = Signal(float)
    viewChanged = Signal()

    def __init__(self):
        super().__init__()

        self.setDragMode(
            QGraphicsView.ScrollHandDrag
        )
        self.setTransformationAnchor(QGraphicsView.AnchorUnderMouse)
        self.setFrameShape(QFrame.NoFrame)
        self.setMouseTracking(True)
        self.setBackgroundBrush(QColor(60, 60, 60))
        self.measure_mode = False
        self.pick_mode = False
        self._p0 = None
        self._line = None

    def set_measure_mode(self, on):
        self.measure_mode = on
        self.setDragMode(QGraphicsView.NoDrag if (on or self.pick_mode) else QGraphicsView.ScrollHandDrag)
        self.viewport().setCursor(Qt.CrossCursor if on else Qt.ArrowCursor)

    def set_pick_mode(self, on):
        self.pick_mode = on
        self.setDragMode(QGraphicsView.NoDrag if (on or self.measure_mode) else QGraphicsView.ScrollHandDrag)
        self.viewport().setCursor(Qt.PointingHandCursor if on else Qt.ArrowCursor)

    def mousePressEvent(self, e):
        if self.pick_mode and e.button() == Qt.LeftButton:
            self.picked.emit(self.mapToScene(e.position().toPoint()))
            return
        if self.measure_mode and e.button() == Qt.LeftButton:
            if self._line is not None:
                self.scene().removeItem(self._line)
            self._p0 = self.mapToScene(e.position().toPoint())
            pen = QPen(QColor(0, 220, 255), 2)
            pen.setCosmetic(True)
            self._line = self.scene().addLine(QLineF(self._p0, self._p0), pen)
            self._line.setZValue(100)
            return
        super().mousePressEvent(e)

    def mouseMoveEvent(self, e):
        p = self.mapToScene(e.position().toPoint())
        self.mouseMoved.emit(p)
        if self.measure_mode and self._p0 is not None:
            self._line.setLine(QLineF(self._p0, p))
            return
        super().mouseMoveEvent(e)

    def mouseReleaseEvent(self, e):
        if self.measure_mode and self._p0 is not None:
            p1 = self.mapToScene(e.position().toPoint())
            self.measured.emit(self._p0, p1)
            self._p0 = None
            return
        super().mouseReleaseEvent(e)

    def wheelEvent(self, e):
        if e.modifiers() & Qt.ControlModifier:      # Ctrl + molette = zoom
            self.wheelZoom.emit(1.15 if e.angleDelta().y() > 0 else 1 / 1.15)
        else:
            super().wheelEvent(e)

    def scrollContentsBy(self, dx, dy):
        super().scrollContentsBy(dx, dy)
        self.viewChanged.emit()

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self.viewChanged.emit()


class Ruler(QWidget):
    """Règle graduée (mm / pt / inch) synchronisée avec le Viewer."""

    def __init__(self, win, horizontal):
        super().__init__()
        self.win, self.h = win, horizontal
        if horizontal:
            self.setFixedHeight(22)
        else:
            self.setFixedWidth(40)

    def paintEvent(self, ev):
        p = QPainter(self)
        p.fillRect(self.rect(), QColor(40, 40, 40))
        p.setPen(QColor(200, 200, 200))
        v, ppu = self.win.viewer, self.win.px_per_unit()
        z = v.transform().m11()
        if ppu <= 0 or z <= 0:
            return
        a = v.mapToScene(0, 0)
        b = v.mapToScene(v.viewport().width(), v.viewport().height())
        s0, s1 = (a.x(), b.x()) if self.h else (a.y(), b.y())
        steps = [0.05, 0.1, 0.25, 0.5, 1, 2, 5, 10, 20, 50, 100, 200, 500, 1000]
        step = next((s for s in steps if s * ppu * z >= 60), steps[-1])
        k, kend = math.floor(s0 / ppu / step), math.ceil(s1 / ppu / step)
        for i in range(k, kend + 1):
            u = i * step
            sc = u * ppu
            if self.h:
                pos = v.mapFromScene(QPointF(sc, 0)).x()
                p.drawLine(pos, 12, pos, 22)
                p.drawText(pos + 2, 10, f"{u:g}")
            else:
                pos = v.mapFromScene(QPointF(0, sc)).y()
                p.drawLine(30, pos, 40, pos)
                p.drawText(2, pos - 2, f"{u:g}")
        # repère rouge = position du curseur
        c = self.win.cursor_scene
        if c is not None:
            p.setPen(QColor(255, 60, 60))
            if self.h:
                x = v.mapFromScene(c).x(); p.drawLine(x, 0, x, 22)
            else:
                y = v.mapFromScene(c).y(); p.drawLine(0, y, 40, y)


class MiniMap(QGraphicsView):
    """Vue miniature de navigation : clic / glisser = recentrer la vue principale."""

    def __init__(self, win):
        super().__init__(win.scene)
        self.win = win
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setInteractive(False)
        self.setMinimumSize(180, 140)

    def refit(self):
        r = self.win.scene.sceneRect()
        if r.isValid() and not r.isEmpty():
            self.fitInView(r, Qt.KeepAspectRatio)

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self.refit()

    def drawForeground(self, painter, rect):
        v = self.win.viewer
        r = v.mapToScene(v.viewport().rect()).boundingRect()
        pen = QPen(QColor(255, 60, 60), 2)
        pen.setCosmetic(True)
        painter.setPen(pen)
        painter.drawRect(r)

    def _go(self, e):
        self.win.viewer.centerOn(self.mapToScene(e.position().toPoint()))

    def mousePressEvent(self, e):
        self._go(e)

    def mouseMoveEvent(self, e):
        if e.buttons() & Qt.LeftButton:
            self._go(e)


class InspectorTable(QWidget):
    """Tableau d'inspection (Overprint / Knockout) avec filtre catégorie et état ON."""

    def __init__(self, state_title):
        super().__init__()
        v, h = QVBoxLayout(self), QHBoxLayout()
        self.lbl = QLabel("—")
        self.cb_cat = QComboBox()
        self.cb_cat.addItems(["Tous", "texte", "fond", "objet"])
        self.chk = QCheckBox("Seulement ON")
        h.addWidget(self.lbl, 1); h.addWidget(self.cb_cat); h.addWidget(self.chk)
        self.tbl = QTableWidget(0, 5)
        self.tbl.setHorizontalHeaderLabels(["Objet", "Catégorie", "Couleur", state_title, "Remarque"])
        self.tbl.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        v.addLayout(h); v.addWidget(self.tbl)
        self.rows = []
        self.cb_cat.currentIndexChanged.connect(self.refresh)
        self.chk.toggled.connect(self.refresh)

    def set_rows(self, rows):
        self.rows = rows
        self.refresh()

    def refresh(self):
        cat, only = self.cb_cat.currentText(), self.chk.isChecked()
        sel = [r for r in self.rows if (cat == "Tous" or r[1] == cat) and (not only or r[3] == "ON")]
        stats = " · ".join(f"{c}: {sum(1 for r in self.rows if r[1] == c and r[3] == 'ON')}/"
                           f"{sum(1 for r in self.rows if r[1] == c)} ON" for c in ("texte", "fond", "objet"))
        self.lbl.setText(f"{len(self.rows)} objets · {stats}" + (f" · {len(sel)} lignes (3000 affichées)" if len(sel) > 3000 else ""))
        sel = sel[:3000]
        self.tbl.setUpdatesEnabled(False)
        self.tbl.setRowCount(len(sel))
        for i, r in enumerate(sel):
            for j, v in enumerate(r):
                self.tbl.setItem(i, j, QTableWidgetItem(v))
        self.tbl.setUpdatesEnabled(True)


# ============================================================================
# FENÊTRE PRINCIPALE
# ============================================================================
class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()

        self.setWindowTitle(
            "Packaging Compare Pro"
        )

        self.resize(
            1700,
            900
        )

        # --- scène principale (overlay / diff / blink) + vue ---
        self.scene = QGraphicsScene()

        self.viewer = Viewer()
        self.viewer.setScene(self.scene)

        # --- vue secondaire (Side By Side) ---
        self.scene_b = QGraphicsScene()
        self.viewer2 = Viewer()
        self.viewer2.setScene(self.scene_b)
        self.viewer2.hide()
        self.sbs_item = self.scene_b.addPixmap(QPixmap())

        # items : original (fixe), modifié (déplaçable), différence
        self.original_item = MovableItem(self.on_modified_moved)   # l'original aussi est déplaçable
        self.scene.addItem(self.original_item)
        self.modified_item = MovableItem(self.on_modified_moved)
        self.scene.addItem(self.modified_item)
        self.modified_item.setZValue(1)
        self.diff_item = self.scene.addPixmap(QPixmap())
        self.diff_item.setZValue(2)
        self.diff_item.hide()

        self.original_png = None   # conservés pour compatibilité
        self.modified_png = None

        self.orig = None           # PdfDoc original
        self.mod = None            # PdfDoc modifié
        self.orig_rgb = None
        self.mod_rgb = None
        self.page = 0
        self.scale = RENDER_SCALE
        self.view_mode = "overlay"  # overlay / diff / sbs / blink
        self.unit = "mm"
        self.cursor_scene = None
        self.channels = []
        self.findings = []
        self._diff_rgb = None
        self.diff_dirty = True

        self.locked = False
        self.active = "modified"      # PDF actif : déplacé par souris et flèches
        self.apply_movable()

        self.blink_timer = QTimer()
        self.blink_timer.setInterval(500)
        self.blink_timer.timeout.connect(
            self.blink
        )

        self.minimap = MiniMap(self)
        self.ruler_h = Ruler(self, True)
        self.ruler_v = Ruler(self, False)

        self.build_ui()

    # ------------------------------------------------------------------
    # INTERFACE
    # ------------------------------------------------------------------
    def build_ui(self):

        self.btn_original = QPushButton(
            "ORIGINAL"
        )

        self.btn_modified = QPushButton(
            "MODIFIED"
        )

        self.btn_align = QPushButton(
            "AUTO ALIGN"
        )

        self.btn_blink = QPushButton(
            "BLINK"
        )

        self.btn_diff = QPushButton(
            "DIFFERENCE"
        )

        self.btn_colors = QPushButton(
            "COLORS"
        )

        self.btn_lock = QPushButton(
            "LOCK"
        )

        self.btn_reset = QPushButton(
            "RESET"
        )

        # nouveaux boutons
        self.btn_overlay = QPushButton("OVERLAY")
        self.btn_sbs = QPushButton("SIDE BY SIDE")
        self.btn_fit = QPushButton("FIT")
        self.btn_measure = QPushButton("MEASURE")
        self.btn_measure.setCheckable(True)
        self.btn_pick = QPushButton("PICK RECIPE")
        self.btn_pick.setCheckable(True)
        self.btn_active = QPushButton("ACTIVE: MODIFIED")
        for b in (self.btn_overlay, self.btn_sbs, self.btn_blink, self.btn_diff):
            b.setCheckable(True)

        self.zoom_slider = QSlider(
            Qt.Horizontal
        )

        self.zoom_slider.setRange(
            10,
            800
        )

        self.zoom_slider.setValue(
            100
        )

        self.opacity_slider = QSlider(
            Qt.Horizontal
        )

        self.opacity_slider.setRange(
            0,
            100
        )

        self.opacity_slider.setValue(
            50
        )

        self.lbl_zoom = QLabel(
            "Zoom 100%"
        )

        self.lbl_opacity = QLabel(
            "Opacity 50%"
        )

        self.lbl_offset = QLabel(
            "X=0  Y=0"
        )
        self.lbl_measure = QLabel("Mesure : —")
        self.lbl_legend = QLabel(
            "<span style='color:#f33'>■ Supprimé</span> &nbsp; "
            "<span style='color:#0a0'>■ Ajouté</span> &nbsp; "
            "<span style='color:#d90'>■ Déplacé</span>")

        self.page_spin = QSpinBox()
        self.page_spin.setRange(1, 1)
        self.page_spin.setPrefix("Page ")
        self.unit_combo = QComboBox()
        self.unit_combo.addItems(["mm", "pt", "inch"])

        # --- barre d'état : Zoom / Coordonnées / Page / Couleur active ---
        self.st_zoom = QLabel("Zoom 100%")
        self.st_coord = QLabel("X —  Y —")
        self.st_page = QLabel("Page —")
        self.st_color = QLabel("Couleur : —")
        for w in (self.st_zoom, self.st_coord, self.st_page, self.st_color):
            w.setMinimumWidth(150)
            self.statusBar().addPermanentWidget(w)

        # --- connexions existantes ---
        self.btn_original.clicked.connect(
            self.load_original
        )

        self.btn_modified.clicked.connect(
            self.load_modified
        )

        self.btn_lock.clicked.connect(
            self.toggle_lock
        )

        self.btn_reset.clicked.connect(
            self.reset_view
        )

        self.btn_blink.clicked.connect(
            self.toggle_blink
        )

        self.btn_align.clicked.connect(
            self.auto_align
        )

        self.btn_diff.clicked.connect(
            self.show_difference
        )

        self.btn_colors.clicked.connect(
            self.show_colors
        )

        self.zoom_slider.valueChanged.connect(
            self.set_zoom
        )

        self.opacity_slider.valueChanged.connect(
            self.set_opacity
        )

        # --- nouvelles connexions ---
        self.btn_overlay.clicked.connect(lambda: self.set_view("overlay"))
        self.btn_sbs.clicked.connect(lambda: self.set_view("sbs"))
        self.btn_fit.clicked.connect(self.fit_view)
        self.btn_measure.toggled.connect(self.on_measure_toggled)
        self.btn_pick.toggled.connect(self.on_pick_toggled)
        self.btn_active.clicked.connect(self.toggle_active)
        self.viewer.picked.connect(self.on_picked)
        self.page_spin.valueChanged.connect(self.change_page)
        self.unit_combo.currentTextChanged.connect(self.change_unit)
        self.viewer.mouseMoved.connect(self.on_mouse_moved)
        self.viewer.measured.connect(self.on_measured)
        self.viewer.wheelZoom.connect(lambda f: self.zoom_slider.setValue(int(self.zoom_slider.value() * f) or 1))
        self.viewer.viewChanged.connect(self.update_rulers)
        # défilement synchronisé pour Side By Side
        for sa, sb in ((self.viewer.horizontalScrollBar(), self.viewer2.horizontalScrollBar()),
                       (self.viewer.verticalScrollBar(), self.viewer2.verticalScrollBar())):
            sa.valueChanged.connect(sb.setValue)
            sb.valueChanged.connect(sa.setValue)

        top = QHBoxLayout()

        for w in [
            self.btn_original,
            self.btn_modified,
            self.btn_align,
            self.btn_blink,
            self.btn_diff,
            self.btn_colors,
            self.btn_lock,
            self.btn_reset
        ]:
            top.addWidget(w)

        top.addWidget(self.btn_overlay)
        top.addWidget(self.btn_sbs)

        top.addWidget(self.lbl_zoom)
        top.addWidget(self.zoom_slider)

        top.addWidget(self.lbl_opacity)
        top.addWidget(self.opacity_slider)

        row2 = QHBoxLayout()
        for w in (self.page_spin, self.unit_combo, self.btn_fit, self.btn_measure, self.btn_pick, self.btn_active):
            row2.addWidget(w)
        row2.addWidget(self.lbl_measure, 1)
        row2.addWidget(self.lbl_legend)

        # zone de vue : règles + viewer(s)
        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(self.viewer)
        splitter.addWidget(self.viewer2)
        grid = QGridLayout()
        grid.setSpacing(0)
        grid.addWidget(QWidget(), 0, 0)
        grid.addWidget(self.ruler_h, 0, 1)
        grid.addWidget(self.ruler_v, 1, 0)
        grid.addWidget(splitter, 1, 1)
        grid.setColumnStretch(1, 1)
        grid.setRowStretch(1, 1)

        layout = QVBoxLayout()

        layout.addLayout(top)
        layout.addLayout(row2)
        layout.addLayout(grid, 1)
        layout.addWidget(self.lbl_offset)

        container = QWidget()
        container.setLayout(layout)

        self.setCentralWidget(container)
        self.build_docks()
        # clavier : flèches = déplacer le PDF actif ; Ctrl +/-/0 = zoom
        self.viewer.installEventFilter(self)
        for seq, fn in (("Ctrl++", lambda: self.zoom_step(1.25)), ("Ctrl+=", lambda: self.zoom_step(1.25)),
                        ("Ctrl+-", lambda: self.zoom_step(0.8)), ("Ctrl+0", self.fit_view)):
            QShortcut(QKeySequence(seq), self).activated.connect(fn)

    def build_docks(self):
        """Panneaux : MiniMap, Layers, Séparations (Output Preview), Colors, Report."""
        def dock(title, widget, area):
            d = QDockWidget(title, self)
            d.setWidget(widget)
            self.addDockWidget(area, d)
            return d

        R, B = Qt.RightDockWidgetArea, Qt.BottomDockWidgetArea
        dock("MiniMap", self.minimap, R)

        # Layers
        self.list_layers = QListWidget()
        self.list_layers.itemChanged.connect(self.on_layer_toggled)
        self.dock_layers = dock("Layers", self.list_layers, R)

        # Séparation Preview
        w = QWidget(); v = QVBoxLayout(w)
        self.list_seps = QListWidget()
        self.list_seps.itemChanged.connect(self.on_sep_toggled)
        self.list_seps.itemClicked.connect(self.on_sep_clicked)   # clic sur l'œil = afficher/masquer
        hb = QHBoxLayout()
        b_all, b_none = QPushButton("Tout afficher"), QPushButton("Tout masquer")
        b_all.clicked.connect(lambda: self.set_all_seps(True))
        b_none.clicked.connect(lambda: self.set_all_seps(False))
        hb.addWidget(b_all); hb.addWidget(b_none)
        v.addWidget(self.list_seps); v.addLayout(hb)
        self.dock_seps = dock("Separations Preview", w, R)

        # Colors
        self.tbl_colors = QTableWidget(0, 4)
        self.tbl_colors.setHorizontalHeaderLabels(["Couleur", "Type", "Couv. orig. %", "Couv. modif. %"])
        self.tbl_colors.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.dock_colors = dock("Colors", self.tbl_colors, B)

        # Report
        w = QWidget(); v = QVBoxLayout(w)
        self.lbl_summary = QLabel("Aucune analyse")
        self.tbl_report = QTableWidget(0, 10)
        self.tbl_report.setHorizontalHeaderLabels(["Statut", "Type", "Changement", "Objet", "ΔX mm", "ΔY mm",
                                                   "X orig", "Y orig", "X modif", "Y modif"])
        self.tbl_report.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tbl_report.cellDoubleClicked.connect(self.goto_finding)
        hb = QHBoxLayout()
        b_an, b_rep = QPushButton("Analyser packaging"), QPushButton("Generate Report")
        b_an.clicked.connect(self.run_analysis)
        b_rep.clicked.connect(self.generate_report)
        hb.addWidget(b_an); hb.addWidget(b_rep); hb.addStretch(1)
        v.addWidget(self.lbl_summary); v.addWidget(self.tbl_report); v.addLayout(hb)
        self.dock_report = dock("Report", w, B)
        self.tabifyDockWidget(self.dock_colors, self.dock_report)
        # Inspecteurs : Color Recipe / Overprint / Knockout / Trap
        self.tabs_insp = QTabWidget()
        self.txt_recipe = QTextEdit(); self.txt_recipe.setReadOnly(True)
        self.txt_recipe.setPlainText("Activez PICK RECIPE puis cliquez sur un objet.")
        self.tbl_op = InspectorTable("Overprint")
        self.tbl_ko = InspectorTable("Knockout")
        self.txt_trap = QTextEdit(); self.txt_trap.setReadOnly(True)
        for w_, t_ in ((self.txt_recipe, "Color Recipe"), (self.tbl_op, "Overprint Inspector"),
                       (self.tbl_ko, "Knockout Inspector"), (self.txt_trap, "Trap Inspector")):
            self.tabs_insp.addTab(w_, t_)
        self.dock_insp = dock("Inspectors", self.tabs_insp, B)
        self.tabifyDockWidget(self.dock_report, self.dock_insp)
        self.dock_colors.raise_()

    # ------------------------------------------------------------------
    # CONVERSIONS D'UNITÉS
    # ------------------------------------------------------------------
    def px_to_unit(self, px):
        pt = px / self.scale
        return pt if self.unit == "pt" else (pt / 72.0 if self.unit == "inch" else pt * 25.4 / 72.0)

    def px_per_unit(self):
        return self.scale * {"pt": 1.0, "inch": 72.0, "mm": 72.0 / 25.4}[self.unit]

    def change_unit(self, u):
        self.unit = u
        self.on_modified_moved(refresh_only=True)
        self.update_rulers()

    def update_rulers(self):
        self.ruler_h.update(); self.ruler_v.update(); self.minimap.viewport().update()

    # ------------------------------------------------------------------
    # CHARGEMENT
    # ------------------------------------------------------------------
    def pdf_to_png(self, pdf):
        """Utilitaire conservé : export PNG d'une page (non utilisé pour l'analyse)."""
        doc = fitz.open(pdf)
        page = doc[0]
        pix = page.get_pixmap(
            matrix=fitz.Matrix(3, 3)
        )
        png = pdf + ".png"
        pix.save(png)
        return png

    def _open_pdf(self, title):
        file, _ = QFileDialog.getOpenFileName(
            self,
            title,
            "",
            "PDF (*.pdf)"
        )
        if not file:
            return None
        try:
            return PdfDoc(file)
        except Exception as ex:
            QMessageBox.critical(self, "Erreur", f"Impossible d'ouvrir le PDF :\n{ex}")
            return None

    def _confirm_replace(self, current):
        """Un PDF est déjà chargé : propose de le remplacer, sans fermer l'application."""
        if current is None:
            return True
        r = QMessageBox.question(self, "Remplacer le PDF", "Remplacer le PDF actuel ?",
                                 QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        return r == QMessageBox.Yes

    def load_original(self):
        if not self._confirm_replace(self.orig):
            return
        d = self._open_pdf("Open Original PDF")
        if d:
            if self.orig:
                self.orig.close()
            self.orig = d
            self.original_item.setTransform(QTransform())
            self.original_item.setPos(0, 0)
            self.findings, self._diff_rgb, self.diff_dirty = [], None, True
            self.fill_report_table()
            self.after_load(first=True)

    def load_modified(self):
        if not self._confirm_replace(self.mod):
            return
        d = self._open_pdf("Open Modified PDF")
        if d:
            if self.mod:
                self.mod.close()
            self.mod = d
            self.modified_item.setTransform(QTransform())
            self.modified_item.setPos(0, 0)
            self.findings, self._diff_rgb, self.diff_dirty = [], None, True
            self.fill_report_table()
            self.after_load(first=True)

    def after_load(self, first=False):
        n = max(d.page_count for d in (self.orig, self.mod) if d)
        self.page_spin.blockSignals(True)
        self.page_spin.setRange(1, n)
        self.page_spin.blockSignals(False)
        self.rebuild_layers()
        self.apply_movable()
        self.reload_images()
        if first:
            self.fit_view()

    def change_page(self, v):
        self.page = v - 1
        for it in (self.original_item, self.modified_item):
            it.setTransform(QTransform())
            it.setPos(0, 0)
        self.findings = []
        self.fill_report_table()
        self.reload_images()

    # ------------------------------------------------------------------
    # RENDU / SÉPARATIONS
    # ------------------------------------------------------------------
    def _pno(self, d):
        return min(self.page, d.page_count - 1)

    def displayed_rgb(self, d):
        """Image affichée : rendu réel si tout est visible, sinon recomposition
        multiplicative des plaques visibles (comme Output Preview)."""
        pno = self._pno(d)
        present = [c for c in self.channels if c.present]
        if all(c.visible for c in present):
            return d.render_rgb(pno, self.scale)
        planes = d.planes(pno, self.scale)
        h, w = next(iter(planes.values())).shape
        out = np.ones((h, w, 3), np.float32)
        for c in present:
            if not c.visible:
                continue
            arrs = [planes[n] for n in c.names if n in planes]
            if not arrs:
                continue
            dens = np.maximum.reduce(arrs).astype(np.float32) / 255.0
            for k in range(3):
                out[:, :, k] *= 1.0 - dens * (1.0 - c.rgb[k] / 255.0)
        return (out * 255).astype(np.uint8)

    def rebuild_channels(self):
        """Canaux = uniquement les encres qui portent réellement de l'encre sur la page
        (lues nativement : DeviceCMYK, Separation, DeviceN). Jamais de couleur inexistante."""
        old = {c.name: c.visible for c in self.channels}
        present, spots = set(), {}
        for d in (self.orig, self.mod):
            if d:
                present |= d.present(self._pno(d), self.scale)
                for n in d.page_spots(self._pno(d)):
                    spots.setdefault(n, d)
        chans = [Channel(n, "process", rgb, [n], group="CMYK") for n, rgb in PROCESS.items() if n in present]
        for n in sorted(spots):
            if n not in present:
                continue
            d = spots[n]
            typ = d.spots[n]
            rgb = SPECIAL_RGB.get(typ) or d.spot_rgb.get(n) or KNOWN_PMS.get(norm(n)) or guess_rgb(n)
            chans.append(Channel(n, "spot", tuple(rgb), [n], group="PMS" if typ == "PMS" else typ))
        chans.sort(key=lambda c: (0 if c.group == "CMYK" else 1 if c.group == "PMS" else 2, c.group))
        for c in chans:
            c.visible = old.get(c.name, True)
        self.channels = chans
        self.fill_sep_list()

    def fill_sep_list(self):
        self.list_seps.blockSignals(True)
        self.list_seps.clear()
        last = None
        for i, c in enumerate(self.channels):
            if c.group != last:                      # titre de groupe : CMYK / PMS / ...
                h = QListWidgetItem(c.group)
                f = h.font(); f.setBold(True); h.setFont(f)
                h.setFlags(Qt.NoItemFlags)
                self.list_seps.addItem(h)
                last = c.group
            it = QListWidgetItem()
            it.setData(Qt.UserRole, i)
            pm = QPixmap(14, 14); pm.fill(QColor(*c.rgb))
            it.setIcon(QIcon(pm))
            it.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable | Qt.ItemIsSelectable)
            it.setCheckState(Qt.Checked if c.visible else Qt.Unchecked)
            self.list_seps.addItem(it)
            self._sep_text(it, c)
        self.list_seps.blockSignals(False)

    def _sep_text(self, it, c):
        it.setText(f"{'👁' if c.visible else '◌'}  {c.name}")

    def on_sep_clicked(self, it):
        """Clic sur l'œil / le nom = afficher ou masquer (la case à cocher gère son propre clic)."""
        if it.data(Qt.UserRole) is None:
            return
        pos = self.list_seps.viewport().mapFromGlobal(QCursor.pos())
        if pos.x() - self.list_seps.visualItemRect(it).left() < 22:
            return
        it.setCheckState(Qt.Unchecked if it.checkState() == Qt.Checked else Qt.Checked)

    def on_sep_toggled(self, it):
        i = it.data(Qt.UserRole)
        if i is None:
            return
        c = self.channels[i]
        c.visible = it.checkState() == Qt.Checked
        self.list_seps.blockSignals(True)
        self._sep_text(it, c)
        self.list_seps.blockSignals(False)
        self.reload_images(rebuild=False)            # mise à jour temps réel

    def set_all_seps(self, on):
        for c in self.channels:
            c.visible = on
        self.fill_sep_list()
        self.reload_images(rebuild=False)

    # ------------------------------------------------------------------
    # LAYERS (calques OCG)
    # ------------------------------------------------------------------
    def rebuild_layers(self):
        self.list_layers.blockSignals(True)
        self.list_layers.clear()
        names = {}
        for d in (self.orig, self.mod):
            if d:
                for n in d.layers:
                    names.setdefault(n, True)
        for n in names:
            it = QListWidgetItem(n)
            it.setFlags(it.flags() | Qt.ItemIsUserCheckable)
            it.setCheckState(Qt.Checked)
            self.list_layers.addItem(it)
        self.list_layers.blockSignals(False)

    def on_layer_toggled(self, it):
        for d in (self.orig, self.mod):          # appliqué par nom aux deux PDF
            if d:
                d.set_layer(it.text(), it.checkState() == Qt.Checked)
        self.reload_images()

    # ------------------------------------------------------------------
    # AFFICHAGE
    # ------------------------------------------------------------------
    def reload_images(self, rebuild=True):
        ref = self.orig or self.mod
        if ref is None:
            return
        self.scale = ref.scale_for(self._pno(ref))
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            if rebuild:
                self.rebuild_channels()
            if self.orig:
                self.orig_rgb = self.displayed_rgb(self.orig)
                self.original_item.setPixmap(to_pixmap(self.orig_rgb))
            if self.mod:
                self.mod_rgb = self.displayed_rgb(self.mod)
                self.modified_item.setPixmap(to_pixmap(self.mod_rgb))
            self.diff_dirty = True
            r = self.original_item.boundingRect() if self.orig else self.modified_item.boundingRect()
            m = 0.25 * max(r.width(), r.height())
            self.scene.setSceneRect(r.adjusted(-m, -m, m, m))
            self.refresh_colors_table()
            if rebuild:
                self.refresh_inspectors()
            self.st_page.setText(f"Page {self.page + 1}/{self.page_spin.maximum()}")
            self.refresh_view()
        finally:
            QApplication.restoreOverrideCursor()

    def set_view(self, mode):
        self.view_mode = mode
        self.refresh_view()

    def refresh_view(self):
        m = self.view_mode
        if m in ("diff", "sbs") and not (self.orig and self.mod):
            QMessageBox.information(self, "Info", "Chargez l'original ET le modifié.")
            m = self.view_mode = "overlay"
        self.original_item.setVisible(bool(self.orig) and m != "diff")
        self.modified_item.setVisible(bool(self.mod) and m in ("overlay", "blink"))
        self.modified_item.setOpacity(1.0 if m == "blink" else self.opacity_slider.value() / 100.0)
        self.diff_item.setVisible(m == "diff")
        self.viewer2.setVisible(m == "sbs")
        if m == "diff":
            self.build_difference()
        elif m == "sbs":
            self.sbs_item.setPixmap(to_pixmap(self.warped_modified()))
            self.sbs_item.setPos(0, 0)
            self.sbs_item.setTransform(self.original_item.sceneTransform())
            self.scene_b.setSceneRect(self.scene.sceneRect())
        if m == "blink":
            self.blink_timer.start()
        else:
            self.blink_timer.stop()
        self.btn_overlay.setChecked(m == "overlay")
        self.btn_diff.setChecked(m == "diff")
        self.btn_sbs.setChecked(m == "sbs")
        self.btn_blink.setChecked(m == "blink")
        self.minimap.refit()
        self.update_rulers()

    # ---- fonctionnalités existantes ----
    def toggle_lock(self):
        self.locked = not self.locked
        self.apply_movable()
        self.btn_lock.setText("UNLOCK" if self.locked else "LOCK")

    def reset_view(self):
        for it in (self.original_item, self.modified_item):
            it.setTransform(QTransform())
            it.setPos(0, 0)
        self.zoom_slider.setValue(100)
        self.opacity_slider.setValue(50)
        self.findings = []
        self.fill_report_table()
        self.diff_dirty = True
        self.set_view("overlay")

    def toggle_blink(self):
        self.set_view("overlay" if self.view_mode == "blink" else "blink")

    def blink(self):
        if self.mod:
            self.modified_item.setVisible(not self.modified_item.isVisible())

    def set_zoom(self, v):
        z = v / 100.0
        self.lbl_zoom.setText(f"Zoom {v}%")
        self.st_zoom.setText(f"Zoom {v}%")
        for vw in (self.viewer, self.viewer2):
            vw.setTransform(QTransform.fromScale(z, z))
        self.update_rulers()

    def set_opacity(self, v):
        self.lbl_opacity.setText(f"Opacity {v}%")
        if self.view_mode == "overlay":
            self.modified_item.setOpacity(v / 100.0)

    def fit_view(self):
        if not self.orig and not self.mod:
            return
        item = self.original_item if self.orig else self.modified_item
        self.viewer.fitInView(item.sceneBoundingRect(), Qt.KeepAspectRatio)
        z = int(min(800, max(10, self.viewer.transform().m11() * 100)))
        self.zoom_slider.setValue(z)
        self.set_zoom(z)

    def show_colors(self):
        vis = self.dock_colors.isVisible() and not self.dock_colors.visibleRegion().isEmpty()
        if vis:
            self.dock_colors.hide()
        else:
            self.dock_colors.show(); self.dock_colors.raise_()

    # ------------------------------------------------------------------
    # DÉPLACEMENT / ALIGNEMENT
    # ------------------------------------------------------------------
    def modified_matrix(self):
        """Matrice 2x3 pixels(modifié) -> pixels(original) : déplacement manuel
        ET auto-align inclus (transformation de scène réelle de l'item)."""
        # relatif à l'original (qui peut aussi avoir été déplacé) : T_mod * T_orig^-1
        t = self.modified_item.sceneTransform() * self.original_item.sceneTransform().inverted()[0]
        return np.array([[t.m11(), t.m21(), t.dx()],
                         [t.m12(), t.m22(), t.dy()]], np.float64)

    def warped_modified(self):
        """Image modifiée recalée dans le repère de l'original."""
        H, W = self.orig_rgb.shape[:2]
        return cv2.warpAffine(self.mod_rgb, self.modified_matrix(), (W, H),
                              flags=cv2.INTER_LINEAR, borderMode=cv2.BORDER_CONSTANT,
                              borderValue=(255, 255, 255))

    def on_modified_moved(self, refresh_only=False):
        p = self.active_item().pos()
        self.lbl_offset.setText(
            f"X={self.px_to_unit(p.x()):.2f}  Y={self.px_to_unit(p.y()):.2f} {self.unit}")
        if not refresh_only:
            self.diff_dirty = True
            self.findings = []

    def auto_align(self):
        if not (self.orig and self.mod):
            QMessageBox.information(self, "Auto Align", "Chargez l'original ET le modifié.")
            return
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:   # on aligne sur le rendu réel (toutes séparations)
            M, info = estimate_alignment(self.orig.render_rgb(self._pno(self.orig), self.scale),
                                         self.mod.render_rgb(self._pno(self.mod), self.scale))
        finally:
            QApplication.restoreOverrideCursor()
        if M is None:
            QMessageBox.warning(self, "Auto Align", "Alignement impossible (pas assez de détails).")
            return
        self.modified_item.setPos(0, 0)
        self.modified_item.setTransform(      # M est exprimé dans le repère de l'original
            QTransform(M[0, 0], M[1, 0], M[0, 1], M[1, 1], M[0, 2], M[1, 2]) * self.original_item.sceneTransform())
        a, s, dx, dy = info
        self.lbl_offset.setText(
            f"Auto-align : rotation {a:.3f}°  échelle {s:.4f}  "
            f"ΔX {self.px_to_unit(dx):.2f}  ΔY {self.px_to_unit(dy):.2f} {self.unit}")
        self.diff_dirty = True
        self.findings = []
        self.refresh_view()

    # ------------------------------------------------------------------
    # PDF ACTIF, CLAVIER, OUTILS (PICK / MEASURE), INSPECTEURS
    # ------------------------------------------------------------------
    def active_pair(self):
        """(PdfDoc, item graphique) du PDF actif."""
        if self.active == "original" and self.orig:
            return self.orig, self.original_item
        if self.mod:
            return self.mod, self.modified_item
        return self.orig, self.original_item

    def active_item(self):
        return self.active_pair()[1]

    def apply_movable(self):
        act = self.active_item()
        for it in (self.original_item, self.modified_item):
            it.setFlag(QGraphicsItem.ItemIsMovable, (it is act) and not self.locked)

    def toggle_active(self):
        self.active = "original" if self.active == "modified" else "modified"
        self.btn_active.setText("ACTIVE: " + self.active.upper())
        self.apply_movable()
        self.on_modified_moved(refresh_only=True)
        self.refresh_inspectors()

    def eventFilter(self, obj, ev):
        """Flèches = déplace le PDF actif (Shift = x10)."""
        if obj is self.viewer and ev.type() == QEvent.KeyPress:
            d = {Qt.Key_Left: (-1, 0), Qt.Key_Right: (1, 0), Qt.Key_Up: (0, -1), Qt.Key_Down: (0, 1)}.get(ev.key())
            if d and not self.locked and self.view_mode in ("overlay", "blink") and (self.orig or self.mod):
                it = self.active_item()
                k = 10 if ev.modifiers() & Qt.ShiftModifier else 1
                it.setPos(it.pos().x() + d[0] * k, it.pos().y() + d[1] * k)
                return True
        return super().eventFilter(obj, ev)

    def zoom_step(self, f):
        self.zoom_slider.setValue(max(10, min(800, int(self.zoom_slider.value() * f))))

    def on_measure_toggled(self, on):
        if on:
            self.btn_pick.setChecked(False)
        self.viewer.set_measure_mode(on)

    def on_pick_toggled(self, on):
        if on:
            self.btn_measure.setChecked(False)
        self.viewer.set_pick_mode(on)

    def on_picked(self, p):
        """Clic sur un objet : recette couleur réelle lue dans le PDF."""
        d, it = self.active_pair()
        if d is None:
            return
        q = it.mapFromScene(p)
        an = d.analysis(self._pno(d))
        txt = recipe_report(an, q.x() / self.scale, q.y() / self.scale) if an else "Analyse indisponible (pikepdf requis)."
        txt += "\n\nRendu natif au pixel : " + self.color_at(int(q.x()), int(q.y()), d)
        self.txt_recipe.setPlainText(txt)
        self.tabs_insp.setCurrentIndex(0)
        self.dock_insp.show(); self.dock_insp.raise_()

    def refresh_inspectors(self):
        d, _ = self.active_pair()
        if d is None:
            return
        an = d.analysis(self._pno(d))
        if an is None:
            self.tbl_op.set_rows([]); self.tbl_ko.set_rows([])
            self.txt_trap.setPlainText("Analyse indisponible : installez pikepdf (pip install pikepdf).")
            return
        op, ko = inspector_rows(an)
        self.tbl_op.set_rows(op)
        self.tbl_ko.set_rows(ko)
        self.txt_trap.setPlainText(trap_report(an))

    # ------------------------------------------------------------------
    # DIFFÉRENCE INTELLIGENTE
    # ------------------------------------------------------------------
    def show_difference(self):
        self.set_view("diff")

    def build_difference(self):
        if self.diff_dirty or self._diff_rgb is None:
            warped = self.warped_modified()
            if not self.findings:
                self.run_analysis(silent=True)
            boxes = []
            for f in self.findings:
                if f.change == "déplacé" and f.status != "OK":
                    for b in (f.bbox_o, f.bbox_m):
                        if b:
                            boxes.append(tuple(int(v * self.scale) for v in b))
            self._diff_rgb = compute_difference(self.orig_rgb, warped, boxes)
            self.diff_item.setPixmap(to_pixmap(self._diff_rgb))
            self.diff_item.setPos(0, 0)      # la différence suit le repère de l'original
            self.diff_item.setTransform(self.original_item.sceneTransform())
            self.diff_dirty = False

    # ------------------------------------------------------------------
    # COMPARAISON PACKAGING
    # ------------------------------------------------------------------
    def run_analysis(self, silent=False):
        if not (self.orig and self.mod):
            if not silent:
                QMessageBox.information(self, "Analyse", "Chargez l'original ET le modifié.")
            return
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            M, s = self.modified_matrix(), self.scale

            def to_orig(b):
                pts = np.array([[b[0], b[1]], [b[2], b[1]], [b[2], b[3]], [b[0], b[3]]]) * s
                q = pts @ M[:, :2].T + M[:, 2]
                return (q[:, 0].min() / s, q[:, 1].min() / s, q[:, 0].max() / s, q[:, 1].max() / s)

            po, pm = self._pno(self.orig), self._pno(self.mod)
            self.findings = compare_objects(self.orig.objects(po), self.mod.objects(pm), to_orig,
                                            lambda b: self.orig.decode_code(po, b),
                                            lambda b: self.mod.decode_code(pm, b))
            self.diff_dirty = True
            self.fill_report_table()
        finally:
            QApplication.restoreOverrideCursor()

    def fill_report_table(self):
        order = {"CRITICAL": 0, "WARNING": 1, "OK": 2}
        rows = sorted([f for f in self.findings if f.change != "identique"],
                      key=lambda f: order[f.status])
        self._rows = rows
        t = self.tbl_report
        t.setRowCount(len(rows))
        for i, f in enumerate(rows):
            xo, xm = f.xy("o"), f.xy("m")
            vals = [f.status, f.kind, f.change, f.label + (" — " + f.details if f.details else ""),
                    f"{f.dx:+.2f}", f"{f.dy:+.2f}", n2(xo, 0), n2(xo, 1), n2(xm, 0), n2(xm, 1)]
            for j, v in enumerate(vals):
                it = QTableWidgetItem(v)
                if j == 0:
                    it.setBackground(QBrush(QColor(STATUS_COLORS[f.status])))
                    it.setForeground(QBrush(QColor("white")))
                t.setItem(i, j, it)
        c = {k: sum(1 for f in self.findings if f.change == k)
             for k in ("déplacé", "ajouté", "supprimé", "modifié", "identique")}
        self.lbl_summary.setText(
            f"Déplacés {c['déplacé']} · Ajoutés {c['ajouté']} · Supprimés {c['supprimé']} · "
            f"Texte modifié {c['modifié']} · Identiques {c['identique']}"
            if self.findings else "Aucune analyse")

    def goto_finding(self, row, col):
        f = self._rows[row]
        b = f.bbox_o or f.bbox_m
        if b:
            c = center(b)
            self.viewer.centerOn(c[0] * self.scale, c[1] * self.scale)

    # ------------------------------------------------------------------
    # COULEURS NATIVES
    # ------------------------------------------------------------------
    def detected_colors(self):
        """[(nom, type, couv. original, couv. modifié)] depuis les plaques natives."""
        cov = [d.coverage(self._pno(d), self.scale) if d else {} for d in (self.orig, self.mod)]
        rows = [(n, "CMYK (DeviceCMYK)", cov[0].get(n), cov[1].get(n)) for n in PROCESS
                if any(d and n in d.present(self._pno(d), self.scale) for d in (self.orig, self.mod))]
        spots = {}
        for d in (self.orig, self.mod):
            if d:
                for n, t in d.spots.items():
                    if n not in d.present(self._pno(d), self.scale):
                        continue          # jamais de couleur absente
                    spots[n] = f"{t} ({d.spot_cs.get(n, 'Separation')})"
        rows += [(n, t, cov[0].get(n), cov[1].get(n)) for n, t in sorted(spots.items())]
        return rows

    def refresh_colors_table(self):
        rows = self.detected_colors()
        t = self.tbl_colors
        t.setRowCount(len(rows))
        for i, r in enumerate(rows):
            for j, v in enumerate(r):
                t.setItem(i, j, QTableWidgetItem("—" if v is None else
                                                 (f"{v:.1f}%" if j >= 2 else str(v))))

    def color_at(self, x, y, d=None):
        d = d or self.orig or self.mod
        pl = d.planes(self._pno(d), self.scale)
        h, w = next(iter(pl.values())).shape
        if not (0 <= x < w and 0 <= y < h):
            return "—"
        parts = [f"{k[0]}{pl[k][y, x] * 100 / 255:.0f}" for k in PROCESS]
        parts += [f"{n} {pl[n][y, x] * 100 / 255:.0f}%" for n in pl if n not in PROCESS and pl[n][y, x] > 5]
        return " ".join(parts)

    # ------------------------------------------------------------------
    # MESURES / BARRE D'ÉTAT
    # ------------------------------------------------------------------
    def on_mouse_moved(self, p):
        self.cursor_scene = p
        u = self.unit
        d, it = self.active_pair()
        q = it.mapFromScene(p) if d else p          # coordonnées dans le PDF actif
        self.st_coord.setText(f"X {self.px_to_unit(q.x()):.2f}  Y {self.px_to_unit(q.y()):.2f} {u}")
        if d:
            self.st_color.setText("Couleur : " + self.color_at(int(q.x()), int(q.y()), d))
        self.ruler_h.update(); self.ruler_v.update()

    def on_measured(self, p0, p1):
        dx, dy = p1.x() - p0.x(), p1.y() - p0.y()
        dist = math.hypot(dx, dy)
        mm = dist / self.scale * 25.4 / 72
        pt = dist / self.scale
        self.lbl_measure.setText(
            f"X={self.px_to_unit(p0.x()):.2f} Y={self.px_to_unit(p0.y()):.2f} | "
            f"ΔX={self.px_to_unit(dx):.2f} ΔY={self.px_to_unit(dy):.2f} {self.unit} | "
            f"Distance = {mm:.2f} mm = {pt:.2f} pt = {pt / 72:.3f} in")

    # ------------------------------------------------------------------
    # RAPPORT PDF / EXCEL
    # ------------------------------------------------------------------
    def generate_report(self):
        if not self.findings:
            self.run_analysis()
        if not self.findings:
            return
        path, flt = QFileDialog.getSaveFileName(self, "Generate Report", "rapport_packaging",
                                                "PDF (*.pdf);;Excel (*.xlsx)")
        if not path:
            return
        xlsx = path.lower().endswith(".xlsx") or (not path.lower().endswith(".pdf") and "xlsx" in flt)
        if not path.lower().endswith((".pdf", ".xlsx")):
            path += ".xlsx" if xlsx else ".pdf"
        try:
            msg = self._report_xlsx(path) if xlsx else self._report_pdf(path)
        except Exception as ex:
            QMessageBox.critical(self, "Rapport", f"Échec : {ex}")
            return
        QMessageBox.information(self, "Rapport", msg or f"Rapport créé :\n{path}")

    def _lists(self):
        f = [x for x in self.findings if x.change != "identique"]
        return ([x for x in f if x.change == "déplacé"],
                [x for x in f if x.change == "supprimé"],
                [x for x in f if x.change in ("ajouté", "modifié")])

    def _special(self):
        """PMS / White Ink / Varnish... réellement présents sur la page."""
        res = {}
        for d in (self.orig, self.mod):
            if d:
                pres = d.present(self._pno(d), self.scale)
                for n, t in d.spots.items():
                    if n in pres:
                        res.setdefault(t, set()).add(n)
        return res

    def _sections(self):
        """Contenu commun aux exports PDF / Excel : [(titre, colonnes, lignes, note)]."""
        fcols = ["Statut", "Type", "Changement", "Objet", "X orig mm", "Y orig mm",
                 "X modif mm", "Y modif mm", "ΔX mm", "ΔY mm"]
        F = lambda f: (f.status, f.kind, f.change, f.label + (" — " + f.details if f.details else ""),
                       n2(f.xy("o"), 0), n2(f.xy("o"), 1), n2(f.xy("m"), 0), n2(f.xy("m"), 1),
                       f"{f.dx:+.2f}", f"{f.dy:+.2f}")
        fl = [f for f in self.findings if f.change != "identique"]
        S = [("Couleurs détectées", ["Couleur", "Type", "Couv. orig. %", "Couv. modif. %"],
              [(n, t, "" if a is None else f"{a:.1f}", "" if b is None else f"{b:.1f}")
               for n, t, a, b in self.detected_colors()], "")]
        sp = self._special()
        S.append(("PMS / White Ink / Varnish détectés", ["Type", "Noms détectés"],
                  [(k, ", ".join(sorted(sp[k])) if k in sp else "non détecté")
                   for k in ("PMS", "White Ink", "Varnish", "Primer", "Emboss", "Braille", "Spot")], ""))
        for lbl, d in (("Original", self.orig), ("Modifié", self.mod)):
            an = d.analysis(self._pno(d)) if d else None
            if an is None:
                continue
            op, ko = inspector_rows(an)
            on_op = [r for r in op if r[3] == "ON"]
            S.append((f"Overprint – {lbl}", ["Objet", "Catégorie", "Couleur", "Overprint", "Remarque"], on_op[:300],
                      f"{len(on_op)} objets en overprint sur {len(op)} (300 premiers listés)"))
            on_ko = sorted((r for r in ko if r[3] == "ON"), key=lambda r: r[1] == "objet")
            S.append((f"Knockout – {lbl}", ["Objet", "Catégorie", "Couleur", "Knockout", "Remarque"], on_ko[:300],
                      f"{len(on_ko)} objets knockout sur {len(ko)} (textes et fonds listés en premier)"))
            S.append((f"Trapping – {lbl}", ["Détail"], [(l,) for l in trap_report(an).splitlines() if l.strip()], ""))
        S.append(("Objets déplacés", fcols, [F(f) for f in fl if f.change == "déplacé" and f.kind not in CODE_KINDS], ""))
        S.append(("Objets supprimés / ajoutés", fcols,
                  [F(f) for f in fl if f.change in ("ajouté", "supprimé") and f.kind not in CODE_KINDS + ("texte",)], ""))
        S.append(("Textes modifiés / ajoutés / supprimés", fcols,
                  [F(f) for f in fl if f.kind == "texte" and f.change != "déplacé"], ""))
        S.append(("Codes-barres / QR / DataMatrix modifiés", fcols, [F(f) for f in fl if f.kind in CODE_KINDS], ""))
        return S

    def _report_pdf(self, path):
        e = lambda s: str(s).replace("&", "&amp;").replace("<", "&lt;")
        html = (f"<h1>Compare Report – Packaging Compare Pro</h1><p>Original : {e(self.orig.path)}<br>"
                f"Modifié : {e(self.mod.path)}<br>Page {self.page + 1} – {datetime.datetime.now():%Y-%m-%d %H:%M}<br>"
                f"Tolérances : OK ≤ {TOL_OK_MM} mm, WARNING ≤ {TOL_WARN_MM} mm, sinon CRITICAL</p>")
        for title, cols, rows, note in self._sections():
            html += f"<h2>{e(title)} ({len(rows)})</h2>" + (f"<p>{e(note)}</p>" if note else "")
            if not rows:
                html += "<p><i>Aucun</i></p>"
                continue
            html += "<table border='1' cellspacing='0' cellpadding='3' width='100%'><tr>" + \
                    "".join(f"<th bgcolor='#dddddd'>{e(c)}</th>" for c in cols) + "</tr>"
            for r in rows:
                html += "<tr>" + "".join(
                    f"<td bgcolor='{STATUS_COLORS[v]}'>{e(v)}</td>" if v in STATUS_COLORS else f"<td>{e(v)}</td>"
                    for v in r) + "</tr>"
            html += "</table>"
        td = QTextDocument()
        if self._diff_rgb is not None:
            k = min(1, 1200 / self._diff_rgb.shape[1])
            img = cv2.resize(self._diff_rgb, None, fx=k, fy=k, interpolation=cv2.INTER_AREA)
            h, w = img.shape[:2]
            qi = QImage(np.ascontiguousarray(img).data, w, h, 3 * w, QImage.Format_RGB888).copy()
            td.addResource(QTextDocument.ResourceType.ImageResource, QUrl("diff.png"), qi)
            html += "<h2>Vue Difference</h2><p><img src='diff.png' width='600'></p>"
        td.setHtml(html)
        wr = QPdfWriter(path)
        wr.setPageSize(QPageSize(QPageSize.A4))
        wr.setPageMargins(QMarginsF(12, 12, 12, 12), QPageLayout.Millimeter)
        wr.setResolution(96)
        td.print_(wr)
        return None

    def _report_xlsx(self, path):
        sections = self._sections()
        try:
            from openpyxl import Workbook
            from openpyxl.styles import PatternFill, Font
        except ImportError:   # repli CSV si openpyxl absent
            p = path.rsplit(".", 1)[0] + ".csv"
            with open(p, "w", newline="", encoding="utf-8-sig") as fh:
                w = csv.writer(fh, delimiter=";")
                for title, cols, rows, note in sections:
                    w.writerow([title]); w.writerow(cols)
                    w.writerows(rows); w.writerow([])
            return f"openpyxl absent : CSV créé\n{p}"
        wb = Workbook()
        wb.remove(wb.active)
        for title, cols, rows, note in sections:
            name = re.sub(r"[\[\]:*?/\\]", "-", title)[:28]
            ws = wb.create_sheet(f"{name}")
            ws.append(cols)
            for c in ws[1]:
                c.font = Font(bold=True)
            for r in rows:
                ws.append(list(r))
                for c in ws[ws.max_row]:
                    if c.value in STATUS_COLORS:
                        c.fill = PatternFill("solid", fgColor=STATUS_COLORS[c.value][1:])
            if note:
                ws.append([]); ws.append([note])
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = 30
        wb.save(path)
        return None


def main():
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
