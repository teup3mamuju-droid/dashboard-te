"""
Konversi file Excel Laporan III-09 (TUL 309) menjadi docs/data/data.json
untuk dashboard web.

Cara pakai:
    python build_data.py

Struktur folder yang dibaca:
    Transaksi Energi/309/<LEVEL>/<TAHUN>/Laporan_III_09_<JENIS>_LAPORAN1_<YYYYMM>[_...].xlsx
    LEVEL  : UP, KABUPATEN, UPI
    JENIS  : NORMAL (pascabayar), LPB (prabayar). TOTAL dihitung = NORMAL + LPB.
"""
from __future__ import annotations

import json
import re
import sys
import warnings
from datetime import datetime
from pathlib import Path

import openpyxl

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

ROOT = Path(__file__).resolve().parent
SRC = ROOT / "Transaksi Energi" / "309"
OUT = ROOT / "docs" / "data" / "data.json"
LEVELS = ["UP", "KABUPATEN", "UPI"]
JENIS = ["NORMAL", "LPB"]

# Kolom C..M (index 2..12 pada tuple baris, kolom A = index 0)
FIELDS = ["c", "va", "kwh", "lwbp", "wbp", "kvarh", "beban", "rpkwh", "rpkvarh", "ttlb", "rp"]
NUM_FIELDS = [f for f in FIELDS if f != "ttlb"]

FNAME_RE = re.compile(
    r"^Laporan_III_09_(?P<jenis>[A-Z]+)_LAPORAN1_(?P<ym>\d{6})"
    r"(?:_[^_]+_(?P<d>\d{2})_(?P<mon>[A-Za-z]{3})_(?P<y>\d{4})_(?P<t>\d{2}-\d{2}-\d{2}))?\.xlsx$"
)
MONTHS_EN = {m: i for i, m in enumerate(
    ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"], 1)}


def num(v) -> float:
    if v is None or isinstance(v, str):
        try:
            return float(str(v).replace(",", "")) if v not in (None, "", "-") else 0.0
        except ValueError:
            return 0.0
    return float(v)


def norm_label(s: str) -> str:
    s = re.sub(r"\s+", " ", str(s)).strip()
    s = re.sub(r"\s+I$", "", s)  # "S.1 / 450 VA I" -> "S.1 / 450 VA"
    return s


def file_stamp(m: re.Match, path: Path) -> datetime:
    if m.group("t"):
        return datetime(int(m.group("y")), MONTHS_EN.get(m.group("mon").title(), 1), int(m.group("d")),
                        *map(int, m.group("t").split("-")))
    return datetime.fromtimestamp(path.stat().st_mtime)


def parse_sheet(path: Path):
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    unit = str(rows[2][1] or "").strip()
    status = str(rows[7][1] or "").split(":", 1)[-1].strip()

    tarif: dict[str, dict] = {}
    order: list[str] = []
    jumlah_gol: dict[str, dict] = {}
    jumlah_total = None
    parent = None
    for r in rows[11:]:
        label = r[1] if len(r) > 1 else None
        if label is None or str(label).strip() == "":
            continue
        raw = str(label).strip()
        vals = {f: num(r[i + 2]) if i + 2 < len(r) else 0.0 for i, f in enumerate(FIELDS)}
        if raw.upper().startswith("JUMLAH"):
            key = raw.upper().replace("JUMLAH", "").strip()
            if key:
                jumlah_gol[key] = vals
            else:
                jumlah_total = vals
            parent = None
            continue
        if raw in ("II", "III", "IV") and parent:
            # Sub-blok tarif: tambahkan LWBP & biaya pemakaian ke tarif induk
            for f in ("lwbp", "wbp", "rpkwh", "rpkvarh"):
                tarif[parent][f] += vals[f]
            continue
        name = norm_label(raw)
        if name in tarif:  # duplikasi label (mis. "B.2 ... II" berdiri sendiri)
            name = f"{name} ({raw.split()[-1]})"
        tarif[name] = vals
        order.append(name)
        parent = name

    # Validasi terhadap baris JUMLAH
    warns = []
    if jumlah_total:
        for f in ("c", "kwh", "rp"):
            s = sum(t[f] for t in tarif.values())
            if abs(s - jumlah_total[f]) > max(1.0, abs(jumlah_total[f]) * 1e-6):
                warns.append(f"{path.name}: total {f} {s:,.0f} != JUMLAH {jumlah_total[f]:,.0f}")
    for g, jv in jumlah_gol.items():
        s = sum(t["rp"] for n, t in tarif.items() if n[0] == g)
        if abs(s - jv["rp"]) > max(1.0, abs(jv["rp"]) * 1e-6):
            warns.append(f"{path.name}: golongan {g} rp {s:,.0f} != JUMLAH {jv['rp']:,.0f}")

    data = {n: [round(tarif[n][f], 3) for f in NUM_FIELDS] for n in order}
    return unit, status, data, order, warns


def main() -> int:
    out = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "fields": NUM_FIELDS,
        "levels": {},
    }
    all_warns: list[str] = []
    for level in LEVELS:
        folder = SRC / level
        out["levels"][level] = {}
        if not folder.exists():
            continue
        # pilih file terbaru per (jenis, bulan, unit-folder)
        chosen: dict[tuple, tuple] = {}
        for p in folder.rglob("Laporan_III_09_*_LAPORAN1_*.xlsx"):
            m = FNAME_RE.match(p.name)
            if not m or m.group("jenis") not in JENIS:
                continue
            key = (str(p.parent), m.group("jenis"), m.group("ym"))
            st = file_stamp(m, p)
            if key not in chosen or st > chosen[key][0]:
                chosen[key] = (st, p, m)
        for (_, jenis, ym), (st, p, m) in sorted(chosen.items(), key=lambda kv: kv[0]):
            unit, status, data, order, warns = parse_sheet(p)
            all_warns += warns
            u = out["levels"][level].setdefault(unit or p.parent.name, {"tarif": [], "months": {}})
            for n in order:
                if n not in u["tarif"]:
                    u["tarif"].append(n)
            mk = f"{ym[:4]}-{ym[4:]}"
            mon = u["months"].setdefault(mk, {"status": {}, "source": {}})
            mon["status"][jenis] = status
            mon["source"][jenis] = p.name
            mon[jenis] = data
            print(f"[{level}] {unit} {mk} {jenis:6} <- {p.name}")

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(out, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    print(f"\nTersimpan: {OUT.relative_to(ROOT)} ({OUT.stat().st_size/1024:.1f} KB)")
    if all_warns:
        print(f"\n{len(all_warns)} PERINGATAN validasi:")
        for w in all_warns:
            print("  -", w)
    else:
        print("Validasi: semua total cocok dengan baris JUMLAH.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
