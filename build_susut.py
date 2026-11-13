"""Pengendalian Susut/Laporan Susut/*.xlsx  ->  docs/data/susut.json

Setiap file Excel menjadi satu "laporan"; setiap sheet ditampilkan sebagai tabel
(nilai terformat, sel gabungan, tebal/warna latar, kolom/baris tersembunyi dibuang).

Pakai:  python build_susut.py
"""
import colorsys
import datetime as dt
import json
import re
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles.colors import COLOR_INDEX
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).parent
SRC = ROOT / "Pengendalian Susut" / "Laporan Susut"
OUT = ROOT / "docs" / "data" / "susut.json"
BULAN = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"]


# ---------- warna ----------
def theme_colors(wb):
    """Palet tema workbook -> list hex (urutan indeks tema Excel: lt1, dk1, lt2, dk2, accent1..6, hlink, folHlink)."""
    xml = wb.loaded_theme.decode("utf8") if wb.loaded_theme else ""
    out = {}
    for name in ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"]:
        m = re.search(rf"<a:{name}>.*?(?:val|lastClr)=\"([0-9A-Fa-f]{{6}})\"", xml, re.S)
        out[name] = m.group(1) if m else None
    order = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"]
    dflt = ["FFFFFF", "000000", "EEECE1", "1F497D", "4F81BD", "C0504D", "9BBB59", "8064A2", "4BACC6", "F79646", "0000FF", "800080"]
    return [out[n] or dflt[i] for i, n in enumerate(order)]


def tint(hexrgb, t):
    if not t:
        return hexrgb
    r, g, b = (int(hexrgb[i:i + 2], 16) / 255 for i in (0, 2, 4))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    l = l * (1 + t) if t < 0 else l * (1 - t) + t
    r, g, b = colorsys.hls_to_rgb(h, min(1, max(0, l)), s)
    return "%02x%02x%02x" % (round(r * 255), round(g * 255), round(b * 255))


def color_hex(c, theme):
    if c is None:
        return None
    try:
        if c.type == "rgb" and isinstance(c.rgb, str) and len(c.rgb) == 8:
            return c.rgb[2:].lower()
        if c.type == "theme":
            return tint(theme[c.theme], c.tint or 0).lower()
        if c.type == "indexed" and c.indexed is not None and c.indexed < len(COLOR_INDEX):
            v = COLOR_INDEX[c.indexed]
            return v[2:].lower() if isinstance(v, str) and len(v) == 8 else None
    except Exception:
        return None
    return None


def lum(h):
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    return (0.299 * r + 0.587 * g + 0.114 * b) / 255


# ---------- format angka (gaya Indonesia: 1.234,56) ----------
def id_num(x, dec):
    s = f"{abs(x):,.{dec}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_value(v, nf):
    """-> (teks tampil, is_angka)"""
    if v is None:
        return None, False
    if isinstance(v, bool):
        return ("BENAR" if v else "SALAH"), False
    if isinstance(v, (dt.datetime, dt.date)):
        if re.search(r"mmm", nf, re.I):
            return f"{BULAN[v.month - 1]}-{str(v.year)[2:]}", False
        return v.strftime("%d/%m/%Y"), False
    if isinstance(v, (int, float)):
        sec = nf.split(";")
        pos = sec[0]
        pct = "%" in pos
        m = re.search(r"\.(0+)", pos.replace('"', ""))
        if nf == "General" or not (m or "#" in pos or "0" in pos):
            if isinstance(v, float) and not v.is_integer():
                t = f"{v:.6f}".rstrip("0").rstrip(".") if abs(v) < 1e6 else f"{v:.2f}"
                return t.replace(".", ","), True
            return str(int(v)), True
        dec = len(m.group(1)) if m else 0
        if pct:
            v = v * 100
        if round(v, dec) == 0 and len(sec) >= 3 and "-" in sec[2] and v == 0:
            return "-", True
        if v == 0 and len(sec) >= 3 and "-" in sec[2]:
            return "-", True
        t = id_num(v, dec)
        if pct:
            t += "%"
        if v < 0 and round(v, dec) != 0:
            t = f"({t})" if (len(sec) > 1 and "(" in sec[1]) else "-" + t
        return t, True
    s = str(v)
    return s, False


# ---------- satu sheet ----------
def build_sheet(ws, theme):
    maxr, maxc = ws.max_row, ws.max_column
    hid_rows = {r for r, d in ws.row_dimensions.items() if d.hidden}
    hid_cols = set()
    for k, d in ws.column_dimensions.items():
        if d.hidden:
            lo, hi = d.min or 0, d.max or 0
            hid_cols.update(range(lo, min(hi, maxc) + 1))
    cols = [c for c in range(1, maxc + 1) if c not in hid_cols]
    rows = [r for r in range(1, maxr + 1) if r not in hid_rows]
    cpos = {c: i for i, c in enumerate(cols)}

    # sel gabungan
    anchor, covered = {}, set()
    for mr in ws.merged_cells.ranges:
        vc = [c for c in range(mr.min_col, mr.max_col + 1) if c in cpos]
        vr = [r for r in range(mr.min_row, mr.max_row + 1) if r not in hid_rows]
        if not vc or not vr:
            continue
        anchor[(vr[0], vc[0])] = (len(vr), len(vc))
        for r in vr:
            for c in vc:
                if (r, c) != (vr[0], vc[0]):
                    covered.add((r, c))
        # sel sumber nilai bisa tersembunyi bila baris/kolom pertama tersembunyi
    styles, sidx = [], {}

    def style_id(c):
        f, fl, al = c.font, c.fill, c.alignment
        st = {}
        if f and f.b:
            st["b"] = 1
        if f and f.i:
            st["i"] = 1
        bg = color_hex(fl.fgColor, theme) if fl and fl.fill_type == "solid" else None
        if bg == "ffffff":
            bg = None
        fg = color_hex(f.color, theme) if f and f.color else None
        if bg:
            st["bg"] = bg
            # teks selalu terbaca di atas warna latar
            if fg is None or abs(lum(fg) - lum(bg)) < 0.35:
                fg = "111111" if lum(bg) > 0.5 else "ffffff"
            st["fg"] = fg
        elif fg:
            # warna teks khas (merah/biru) dipertahankan; hitam/putih/abu ikut tema
            r, g, b = (int(fg[i:i + 2], 16) for i in (0, 2, 4))
            if max(r, g, b) - min(r, g, b) > 40:
                st["fg"] = fg
        if al and al.horizontal in ("center", "centerContinuous"):
            st["al"] = "c"
        elif al and al.horizontal == "right":
            st["al"] = "r"
        elif al and al.horizontal == "left":
            st["al"] = "l"
        if al and al.vertical in ("center", "bottom") and al.wrap_text:
            st["w"] = 1
        elif al and al.wrap_text:
            st["w"] = 1
        if al and al.indent:
            st["in"] = int(al.indent)
        if c.border and (c.border.bottom.style or c.border.top.style):
            st["bd"] = 1
        key = json.dumps(st, sort_keys=True)
        if key not in sidx:
            sidx[key] = len(styles)
            styles.append(st)
        return sidx[key]

    out_rows, last_nonempty = [], 0
    for r in rows:
        line, has = [], False
        for c in cols:
            if (r, c) in covered:
                continue
            cell = ws.cell(r, c)
            txt, isnum = fmt_value(cell.value, cell.number_format)
            sid = style_id(cell)
            rs, cs = anchor.get((r, c), (1, 1))
            err = isinstance(cell.value, str) and cell.value.startswith("#") and cell.value.endswith(("!", "?", "A"))
            if txt is None and sid == 0 and rs == 1 and cs == 1:
                line.append(None)
                continue
            line.append([txt, sid, cs, rs, 1 if isnum else 0, 1 if err else 0])
            if txt is not None or sid:
                has = True
        # buang sel kosong di ujung kanan
        while line and line[-1] is None:
            line.pop()
        out_rows.append(line)
        if any(x is not None and x[0] is not None for x in line):
            last_nonempty = len(out_rows)
    out_rows = out_rows[:last_nonempty]  # buang baris kosong di bawah

    # lebar kolom (karakter)
    widths = []
    for c in cols:
        d = ws.column_dimensions.get(get_column_letter(c))
        w = d.width if d is not None and d.width else None
        if w is None:  # kolom dalam grup (min..max)
            for k, dd in ws.column_dimensions.items():
                if dd.min and dd.max and dd.min <= c <= dd.max and dd.width:
                    w = dd.width
                    break
        widths.append(round(w if w else 8.43, 1))

    # baris beku (kolom beku)
    fr = fc = 0
    if ws.freeze_panes:
        m = re.match(r"([A-Z]+)(\d+)", ws.freeze_panes)
        from openpyxl.utils import column_index_from_string as ci
        fc = sum(1 for c in cols if c < ci(m.group(1)))
        fr = sum(1 for r in rows if r < int(m.group(2)))

    return {
        "name": ws.title,
        "widths": widths,
        "styles": styles,
        "rows": out_rows,
        "freeze": [fr, fc],
    }


def build_file(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    theme = theme_colors(wb)
    sheets = []
    for ws in wb.worksheets:
        if ws.sheet_state != "visible":
            continue
        sh = build_sheet(ws, theme)
        sheets.append(sh)
        print(f"  {ws.title:14} {len(sh['rows']):4} baris x {len(sh['widths']):3} kolom")
    return {
        "file": path.name,
        "title": path.stem,
        "modified": dt.datetime.fromtimestamp(path.stat().st_mtime).isoformat(timespec="seconds"),
        "sheets": sheets,
    }


def main():
    files = sorted(p for p in SRC.glob("*.xlsx") if not p.name.startswith("~$"))
    if not files:
        sys.exit(f"Tidak ada file Excel di {SRC}")
    reports = []
    for p in files:
        print(p.name)
        reports.append(build_file(p))
    data = {
        "generated_at": dt.datetime.now().isoformat(timespec="seconds"),
        "path": ["Pengendalian Susut", "Laporan Susut"],
        "reports": reports,
    }
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(data, ensure_ascii=False, separators=(",", ":")), encoding="utf8")
    print(f"Tulis {OUT} ({OUT.stat().st_size / 1e6:.2f} MB)")


if __name__ == "__main__":
    main()
