"""
Konversi file Excel Laporan III-09 (TUL 309) menjadi docs/data/data.json
untuk dashboard web.

Cara pakai:
    python build_data.py

Folder yang dibaca: Transaksi Energi/309/<LEVEL>/<TAHUN>/*.xlsx  (LEVEL = UP, KABUPATEN, UPI)

Format yang dikenali:
  A. Laporan_III_09_<JENIS>_LAPORAN1_<YYYYMM>[_<kode>][_dd_Mon_yyyy_hh-mm-ss].xlsx
     (format aplikasi pusat; JENIS = NORMAL / LPB / TOTAL)
  B. TUL_III_09_KAB_<NAMA>_<YYYYMM>_<JENIS>_<YYYYMMDD>.xlsx   (format kabupaten)

Aturan:
  - Jika NORMAL & LPB ada untuk suatu bulan, TOTAL dihitung = NORMAL + LPB (file TOTAL diabaikan).
  - Jika hanya TOTAL yang ada, TOTAL dipakai apa adanya (pascabayar/prabayar tidak tersedia).
  - Beberapa file untuk unit/bulan/jenis yang sama -> dipakai yang terbaru.
  - Nama unit untuk level UPI diambil dari kode di nama file; nama bisa diatur di unit_names.json.
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
NAMES_FILE = ROOT / "unit_names.json"
LEVELS = ["UP", "UPI", "KABUPATEN"]
CODE_UNIT_LEVELS = {"UPI"}  # level yang unitnya dibedakan oleh kode di nama file

FIELDS = ["c", "va", "kwh", "lwbp", "wbp", "kvarh", "beban", "rpkwh", "rpkvarh", "rp"]
SUB_FIELDS_A = ("lwbp", "wbp", "rpkwh", "rpkvarh")             # format A: Jumlah Rp hanya di baris induk
SUB_FIELDS_B = ("lwbp", "wbp", "beban", "rpkwh", "rpkvarh", "rp")  # format B: tiap blok punya Jumlah Rp

RE_A = re.compile(
    r"^Laporan_III_09_(?P<jenis>[A-Z]+)_LAPORAN1_(?P<ym>\d{6})"
    r"(?:_(?P<code>(?!\d{2}_[A-Za-z]{3}_)[0-9A-Za-z]+))?"
    r"(?:_(?P<d>\d{2})_(?P<mon>[A-Za-z]{3})_(?P<y>\d{4})_(?P<t>\d{2}-\d{2}-\d{2}))?\.xlsx$"
)
RE_B = re.compile(r"^TUL_III_09_KAB_(?P<name>.+?)_(?P<ym>\d{6})_(?P<jenis>[A-Z]+)_(?P<stamp>\d{8})\.xlsx$")
MONTHS_EN = {m: i for i, m in enumerate(
    ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"], 1)}


def num(v) -> float:
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "")
    try:
        return float(s) if s not in ("", "-") else 0.0
    except ValueError:
        return 0.0


def norm_label(s: str) -> str:
    s = re.sub(r"\s+", " ", str(s)).strip()
    return re.sub(r"\s+I$", "", s)  # "S.1 / 450 VA I" -> "S.1 / 450 VA"


def cell(rows, r, c):
    try:
        v = rows[r][c]
    except IndexError:
        return ""
    return "" if v is None else str(v).strip()


# ---------------------------------------------------------------- format A
def parse_a(path: Path):
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    info = {
        "up": cell(rows, 2, 1),
        "sub": cell(rows, 3, 1),  # nama ULP (kadang terisi)
        "status": cell(rows, 7, 1).split(":", 1)[-1].strip(),
    }
    tarif, order, jumlah_total, parent = {}, [], None, None
    for r in rows[11:]:
        raw = cell([r], 0, 1)
        if not raw:
            continue
        vals = {f: num(r[i + 2]) if i + 2 < len(r) else 0.0 for i, f in
                enumerate(["c", "va", "kwh", "lwbp", "wbp", "kvarh", "beban", "rpkwh", "rpkvarh", "_ttlb", "rp"])}
        vals.pop("_ttlb")
        if raw.upper().startswith("JUMLAH"):
            if raw.upper().strip() == "JUMLAH":
                jumlah_total = vals
            parent = None
            continue
        if raw in ("II", "III", "IV") and parent:
            for f in SUB_FIELDS_A:
                tarif[parent][f] += vals[f]
            continue
        name = norm_label(raw)
        if name in tarif:
            name = f"{name} ({raw.split()[-1]})"
        tarif[name], parent = vals, name
        order.append(name)

    warns = []
    if jumlah_total:
        for f in ("c", "kwh", "rp"):
            s = sum(t[f] for t in tarif.values())
            if abs(s - jumlah_total[f]) > max(1.0, abs(jumlah_total[f]) * 1e-6):
                warns.append(f"{path.name}: total {f} {s:,.0f} != JUMLAH {jumlah_total[f]:,.0f}")
    else:
        warns.append(f"{path.name}: baris JUMLAH tidak ditemukan")
    return info, tarif, order, warns


# ---------------------------------------------------------------- format B
def parse_b(path: Path):
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    info = {"name": cell(rows, 1, 0).split("·")[0].strip(), "status": "", "pulled": ""}
    m = re.search(r"ditarik\s+(\S+)", cell(rows, 1, 0))
    if m:
        info["pulled"] = m.group(1)
    hdr = next((i for i, r in enumerate(rows) if r and str(r[0] or "").strip().lower().startswith("golongan")), None)
    if hdr is None:
        raise ValueError(f"{path.name}: header 'Golongan tarif' tidak ditemukan")
    cols = [str(c or "").strip().lower() for c in rows[hdr]]

    def ci(*names):
        for n in names:
            if n in cols:
                return cols.index(n)
        raise ValueError(f"{path.name}: kolom {names[0]} tidak ditemukan")

    idx = {
        "c": ci("pelanggan"), "va": ci("daya (va)"), "kwh": ci("kwh"), "lwbp": ci("lwbp"), "wbp": ci("wbp"),
        "kvarh": ci("kel. kvarh"), "beban": ci("biaya beban (rp)"), "rpkwh": ci("kwh (rp)"),
        "rpkvarh": ci("kel. kvarh (rp)"), "rp": ci("jumlah (rp)"),
    }
    i_blok = cols.index("blok") if "blok" in cols else None
    tarif, order, parent = {}, [], None
    for r in rows[hdr + 1:]:
        if not r or not str(r[0] or "").strip():
            continue
        label = norm_label(r[0])
        blok = str(r[i_blok] or "").strip().upper() if i_blok is not None and i_blok < len(r) else ""
        vals = {f: num(r[i]) if i < len(r) else 0.0 for f, i in idx.items()}
        if blok in ("II", "III", "IV") and parent == label:
            for f in SUB_FIELDS_B:
                tarif[label][f] += vals[f]
            continue
        tarif[label], parent = vals, label
        order.append(label)

    warns = []
    for n, t in tarif.items():
        s = t["beban"] + t["rpkwh"] + t["rpkvarh"]
        if abs(s - t["rp"]) > max(2.0, abs(t["rp"]) * 1e-6):
            warns.append(f"{path.name}: {n} jumlah Rp {t['rp']:,.0f} != komponen {s:,.0f}")
    return info, tarif, order, warns


# ---------------------------------------------------------------- main
def stamp_a(m, path: Path) -> datetime:
    if m.group("t"):
        return datetime(int(m.group("y")), MONTHS_EN.get(m.group("mon").title(), 1), int(m.group("d")),
                        *map(int, m.group("t").split("-")))
    return datetime.fromtimestamp(path.stat().st_mtime)


def main() -> int:
    names = {}
    if NAMES_FILE.exists():
        names = {k: v for k, v in json.loads(NAMES_FILE.read_text(encoding="utf-8")).items() if not k.startswith("_")}

    out = {"generated_at": datetime.now().isoformat(timespec="seconds"), "fields": FIELDS, "levels": {}}
    all_warns, unknown = [], []

    for level in LEVELS:
        folder = SRC / level
        lv = out["levels"][level] = {}
        if not folder.exists():
            continue
        # 1) kumpulkan kandidat file: key (unitkey, jenis, ym) -> file terbaru
        chosen = {}
        for p in sorted(folder.rglob("*.xlsx")):
            if p.name.startswith("~$"):
                continue
            ma, mb = RE_A.match(p.name), RE_B.match(p.name)
            if ma:
                code = ma.group("code") or ""
                ukey = code if level in CODE_UNIT_LEVELS and code else "_sheet"
                key, st, fmt = (ukey, ma.group("jenis"), ma.group("ym")), stamp_a(ma, p), "A"
            elif mb:
                key = ("KAB " + mb.group("name"), mb.group("jenis"), mb.group("ym"))
                st, fmt = datetime.strptime(mb.group("stamp"), "%Y%m%d"), "B"
            else:
                if "LAPORAN2" not in p.name:
                    unknown.append(str(p.relative_to(ROOT)))
                continue
            if key[1] not in ("NORMAL", "LPB", "TOTAL"):
                continue
            if key not in chosen or st > chosen[key][0]:
                chosen[key] = (st, p, fmt)

        # 2) parse
        learned = {}  # kode -> nama ULP dari isi file
        parsed = []
        for (ukey, jenis, ym), (st, p, fmt) in sorted(chosen.items(), key=lambda kv: (kv[0][0], kv[0][2], kv[0][1])):
            if fmt == "A":
                info, tarif, order, warns = parse_a(p)
                if ukey == "_sheet":
                    uname = info["sub"] or info["up"] or p.parent.name
                else:
                    uname = None
                    if info["sub"]:
                        learned[ukey] = info["sub"]
                parent_name = info["up"]
            else:
                info, tarif, order, warns = parse_b(p)
                uname, parent_name = info["name"] or ukey, ""
            all_warns += warns
            parsed.append((ukey, uname, parent_name, jenis, ym, p, info, tarif, order))

        for ukey, uname, parent_name, jenis, ym, p, info, tarif, order in parsed:
            if uname is None:  # unit berdasarkan kode
                uname = names.get(ukey) or learned.get(ukey) or f"Unit {ukey}"
            u = lv.setdefault(uname, {"code": ukey if ukey != "_sheet" else "", "parent": parent_name,
                                      "tarif": [], "months": {}})
            for n in order:
                if n not in u["tarif"]:
                    u["tarif"].append(n)
            mk = f"{ym[:4]}-{ym[4:]}"
            mon = u["months"].setdefault(mk, {"status": {}, "source": {}})
            mon["status"][jenis] = info.get("status") or (("ditarik " + info["pulled"][:10]) if info.get("pulled") else "")
            mon["source"][jenis] = p.name
            mon[jenis] = {n: [round(tarif[n][f], 3) for f in FIELDS] for n in order}

        # 3) TOTAL tidak perlu disimpan jika NORMAL+LPB lengkap
        for uname, u in lv.items():
            for mk, mon in u["months"].items():
                if "NORMAL" in mon and "LPB" in mon and "TOTAL" in mon:
                    del mon["TOTAL"], mon["status"]["TOTAL"], mon["source"]["TOTAL"]
            for mk, mon in u["months"].items():
                if "TOTAL" not in mon and ("NORMAL" in mon) != ("LPB" in mon):
                    ada, kurang = ("NORMAL", "LPB") if "NORMAL" in mon else ("LPB", "NORMAL")
                    all_warns.append(f"{uname} {mk}: hanya ada file {ada} — file {kurang} belum ada, total bulan ini tidak lengkap")
            months = sorted(u["months"])
            print(f"[{level:9}] {uname:22} {len(months):2} bulan ({months[0]} s.d. {months[-1]})"
                  f"  jenis: {sorted({j for m in u['months'].values() for j in m if j in ('NORMAL','LPB','TOTAL')})}")

    # 4) cek silang: jumlah unit UPI == UP per bulan
    up = out["levels"].get("UP", {})
    upi = out["levels"].get("UPI", {})
    for up_name, upd in up.items():
        kids = [u for u in upi.values() if u["parent"] == up_name]
        if not kids:
            continue
        for mk, mon in sorted(upd["months"].items()):
            if not all(mk in k["months"] for k in kids):
                continue
            up_c = sum(r[0] for j in ("NORMAL", "LPB") for r in mon.get(j, {}).values()) or \
                sum(r[0] for r in mon.get("TOTAL", {}).values())
            kid_c = sum(r[0] for k in kids for r in k["months"][mk].get("TOTAL", {}).values())
            if abs(up_c - kid_c) > 0.5:
                all_warns.append(f"Cek silang {mk}: pelanggan {up_name} {up_c:,.0f} != jumlah {len(kids)} unit UPI {kid_c:,.0f}")

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(out, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    print(f"\nTersimpan: {OUT.relative_to(ROOT)} ({OUT.stat().st_size / 1024:.1f} KB)")
    if unknown:
        print(f"\n{len(unknown)} file tidak dikenali (dilewati):")
        for u in unknown:
            print("  -", u)
    unnamed = sorted({u for lv in out["levels"].values() for u in lv if u.startswith("Unit ")})
    if unnamed:
        print(f"\nUnit tanpa nama (isi di unit_names.json): {', '.join(unnamed)}")
    if all_warns:
        print(f"\n{len(all_warns)} PERINGATAN validasi:")
        for w in all_warns:
            print("  -", w)
    else:
        print("Validasi: semua total cocok.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
