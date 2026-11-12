/* Pembaca Excel TUL III-09 -> data.json, di browser.
 * Port 1:1 dari build_data.py — jika logika diubah di salah satu, ubah juga yang lain.
 * Bisa dipakai di browser (window.TulBuilder) maupun Node (module.exports) untuk pengujian. */
(function (root) {
  'use strict';

  const LEVELS = ['UP', 'UPI', 'KABUPATEN'];
  const CODE_UNIT_LEVELS = new Set(['UPI']);
  const FIELDS = ['c', 'va', 'kwh', 'lwbp', 'wbp', 'kvarh', 'beban', 'rpkwh', 'rpkvarh', 'rp'];
  const COLS_A = ['c', 'va', 'kwh', 'lwbp', 'wbp', 'kvarh', 'beban', 'rpkwh', 'rpkvarh', '_ttlb', 'rp'];
  const SUB_FIELDS_A = ['lwbp', 'wbp', 'rpkwh', 'rpkvarh'];
  const SUB_FIELDS_B = ['lwbp', 'wbp', 'beban', 'rpkwh', 'rpkvarh', 'rp'];
  const RE_A = /^Laporan_III_09_([A-Z]+)_LAPORAN1_(\d{6})(?:_((?!\d{2}_[A-Za-z]{3}_)[0-9A-Za-z]+))?(?:_(\d{2})_([A-Za-z]{3})_(\d{4})_(\d{2}-\d{2}-\d{2}))?\.xlsx$/;
  const RE_B = /^TUL_III_09_KAB_(.+?)_(\d{6})_([A-Z]+)_(\d{8})\.xlsx$/;
  const MONTHS_EN = { Jan: 1, Feb: 2, Mar: 3, Apr: 4, May: 5, Jun: 6, Jul: 7, Aug: 8, Sep: 9, Oct: 10, Nov: 11, Dec: 12 };

  function num(v) {
    if (v === null || v === undefined) return 0;
    if (typeof v === 'number') return v;
    const s = String(v).trim().replace(/,/g, '');
    if (s === '' || s === '-') return 0;
    const n = Number(s);
    return Number.isFinite(n) ? n : 0;
  }
  const normLabel = (s) => String(s).replace(/\s+/g, ' ').trim().replace(/\s+I$/, '');
  function cell(rows, r, c) {
    const row = rows[r];
    const v = row ? row[c] : undefined;
    return v === null || v === undefined ? '' : String(v).trim();
  }
  const r3 = (x) => Math.round(x * 1000) / 1000;
  const fmt0 = (x) => Math.round(x).toLocaleString('en-US');
  const tol = (ref, min) => Math.max(min, Math.abs(ref) * 1e-6);

  /** Baca sheet pertama menjadi array baris (kolom A = index 0, baris 1 = index 0). */
  function sheetRows(XLSX, buf) {
    const wb = XLSX.read(buf, { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    if (!ws || !ws['!ref']) return [];
    const rg = XLSX.utils.decode_range(ws['!ref']);
    rg.s.r = 0; rg.s.c = 0;
    return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true, blankrows: true, range: rg });
  }

  // ------------------------------------------------------------ format A
  function parseA(rows, fname) {
    const st = cell(rows, 7, 1);
    const info = { up: cell(rows, 2, 1), sub: cell(rows, 3, 1), status: (st.includes(':') ? st.slice(st.indexOf(':') + 1) : st).trim() };
    const tarif = {}, order = [];
    let jumlahTotal = null, parent = null;
    for (const r of rows.slice(11)) {
      if (!r) continue;
      const raw = cell([r], 0, 1);
      if (!raw) continue;
      const vals = {};
      COLS_A.forEach((f, i) => { vals[f] = i + 2 < r.length ? num(r[i + 2]) : 0; });
      delete vals._ttlb;
      if (raw.toUpperCase().startsWith('JUMLAH')) {
        if (raw.toUpperCase().trim() === 'JUMLAH') jumlahTotal = vals;
        parent = null;
        continue;
      }
      if (['II', 'III', 'IV'].includes(raw) && parent) {
        for (const f of SUB_FIELDS_A) tarif[parent][f] += vals[f];
        continue;
      }
      let name = normLabel(raw);
      if (tarif[name]) name = `${name} (${raw.split(/\s+/).pop()})`;
      tarif[name] = vals; parent = name;
      order.push(name);
    }
    const warns = [];
    if (jumlahTotal) {
      for (const f of ['c', 'kwh', 'rp']) {
        const s = Object.values(tarif).reduce((a, t) => a + t[f], 0);
        if (Math.abs(s - jumlahTotal[f]) > tol(jumlahTotal[f], 1)) warns.push(`${fname}: total ${f} ${fmt0(s)} != JUMLAH ${fmt0(jumlahTotal[f])}`);
      }
    } else warns.push(`${fname}: baris JUMLAH tidak ditemukan`);
    return { info, tarif, order, warns };
  }

  // ------------------------------------------------------------ format B
  function parseB(rows, fname) {
    const head = cell(rows, 1, 0);
    const info = { name: head.split('·')[0].trim(), status: '', pulled: '' };
    const m = head.match(/ditarik\s+(\S+)/);
    if (m) info.pulled = m[1];
    const hdr = rows.findIndex((r) => r && String(r[0] == null ? '' : r[0]).trim().toLowerCase().startsWith('golongan'));
    if (hdr < 0) throw new Error(`${fname}: header 'Golongan tarif' tidak ditemukan`);
    const cols = rows[hdr].map((c) => String(c == null ? '' : c).trim().toLowerCase());
    const ci = (n) => { const i = cols.indexOf(n); if (i < 0) throw new Error(`${fname}: kolom ${n} tidak ditemukan`); return i; };
    const idx = {
      c: ci('pelanggan'), va: ci('daya (va)'), kwh: ci('kwh'), lwbp: ci('lwbp'), wbp: ci('wbp'),
      kvarh: ci('kel. kvarh'), beban: ci('biaya beban (rp)'), rpkwh: ci('kwh (rp)'),
      rpkvarh: ci('kel. kvarh (rp)'), rp: ci('jumlah (rp)'),
    };
    const iBlok = cols.indexOf('blok');
    const tarif = {}, order = [];
    let parent = null;
    for (const r of rows.slice(hdr + 1)) {
      if (!r || !String(r[0] == null ? '' : r[0]).trim()) continue;
      const label = normLabel(r[0]);
      const blok = iBlok >= 0 && iBlok < r.length ? String(r[iBlok] == null ? '' : r[iBlok]).trim().toUpperCase() : '';
      const vals = {};
      for (const [f, i] of Object.entries(idx)) vals[f] = i < r.length ? num(r[i]) : 0;
      if (['II', 'III', 'IV'].includes(blok) && parent === label) {
        for (const f of SUB_FIELDS_B) tarif[label][f] += vals[f];
        continue;
      }
      tarif[label] = vals; parent = label;
      order.push(label);
    }
    const warns = [];
    for (const [n, t] of Object.entries(tarif)) {
      const s = t.beban + t.rpkwh + t.rpkvarh;
      if (Math.abs(s - t.rp) > tol(t.rp, 2)) warns.push(`${fname}: ${n} jumlah Rp ${fmt0(t.rp)} != komponen ${fmt0(s)}`);
    }
    return { info, tarif, order, warns };
  }

  function levelOf(path) {
    for (const seg of path.split(/[\\/]/).slice(0, -1)) {
      const u = seg.toUpperCase();
      if (LEVELS.includes(u)) return u;
    }
    return null;
  }

  /**
   * files: [{ path, name, lastModified, rows }]   (rows = hasil sheetRows)
   * names: { kode: namaUnit } — nama tetap untuk unit berkode (mis. dari data lama / unit_names.json)
   */
  function build(files, names) {
    names = names || {};
    const now = new Date();
    const pad = (n) => String(n).padStart(2, '0');
    const out = {
      generated_at: `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}T${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`,
      fields: FIELDS, levels: {},
    };
    const warns = [], unknown = [], log = [];

    for (const level of LEVELS) {
      const lv = out.levels[level] = {};
      const chosen = new Map();
      for (const f of files) {
        if (levelOf(f.path) !== level || f.name.startsWith('~$') || !/\.xlsx$/i.test(f.name)) continue;
        const ma = f.name.match(RE_A), mb = f.name.match(RE_B);
        let key, st, fmt;
        if (ma) {
          const code = ma[3] || '';
          const ukey = CODE_UNIT_LEVELS.has(level) && code ? code : '_sheet';
          key = [ukey, ma[1], ma[2]];
          st = ma[7] ? new Date(+ma[6], (MONTHS_EN[ma[5][0].toUpperCase() + ma[5].slice(1).toLowerCase()] || 1) - 1, +ma[4],
            ...ma[7].split('-').map(Number)).getTime() : f.lastModified;
          fmt = 'A';
        } else if (mb) {
          key = ['KAB ' + mb[1], mb[3], mb[2]];
          st = new Date(+mb[4].slice(0, 4), +mb[4].slice(4, 6) - 1, +mb[4].slice(6, 8)).getTime();
          fmt = 'B';
        } else {
          if (!f.name.includes('LAPORAN2')) unknown.push(f.path);
          continue;
        }
        if (!['NORMAL', 'LPB', 'TOTAL'].includes(key[1])) continue;
        const k = key.join('|');
        if (!chosen.has(k) || st > chosen.get(k).st) chosen.set(k, { key, st, f, fmt });
      }

      const learned = {}, parsed = [];
      const sorted = [...chosen.values()].sort((a, b) => {
        const ka = [a.key[0], a.key[2], a.key[1]], kb = [b.key[0], b.key[2], b.key[1]];
        for (let i = 0; i < 3; i++) if (ka[i] !== kb[i]) return ka[i] < kb[i] ? -1 : 1;
        return 0;
      });
      for (const { key: [ukey, jenis, ym], f, fmt } of sorted) {
        let res, uname, parentName;
        if (fmt === 'A') {
          res = parseA(f.rows, f.name);
          if (ukey === '_sheet') uname = res.info.sub || res.info.up || level;
          else { uname = null; if (res.info.sub) learned[ukey] = res.info.sub; }
          parentName = res.info.up;
        } else {
          res = parseB(f.rows, f.name);
          uname = res.info.name || ukey; parentName = '';
        }
        warns.push(...res.warns);
        parsed.push({ ukey, uname, parentName, jenis, ym, f, ...res });
      }

      for (const p of parsed) {
        const uname = p.uname === null ? (names[p.ukey] || learned[p.ukey] || `Unit ${p.ukey}`) : p.uname;
        const u = lv[uname] = lv[uname] || { code: p.ukey !== '_sheet' ? p.ukey : '', parent: p.parentName, tarif: [], months: {} };
        for (const n of p.order) if (!u.tarif.includes(n)) u.tarif.push(n);
        const mk = `${p.ym.slice(0, 4)}-${p.ym.slice(4)}`;
        const mon = u.months[mk] = u.months[mk] || { status: {}, source: {} };
        mon.status[p.jenis] = p.info.status || (p.info.pulled ? 'ditarik ' + p.info.pulled.slice(0, 10) : '');
        mon.source[p.jenis] = p.f.name;
        const rows = {};
        for (const n of p.order) rows[n] = FIELDS.map((k) => r3(p.tarif[n][k]));
        mon[p.jenis] = rows;
      }

      for (const [uname, u] of Object.entries(lv)) {
        for (const mon of Object.values(u.months)) {
          if (mon.NORMAL && mon.LPB && mon.TOTAL) { delete mon.TOTAL; delete mon.status.TOTAL; delete mon.source.TOTAL; }
        }
        for (const [mk, mon] of Object.entries(u.months)) {
          if (!mon.TOTAL && !!mon.NORMAL !== !!mon.LPB) warns.push(`${uname} ${mk}: hanya ada file ${mon.NORMAL ? 'NORMAL' : 'LPB'} — file ${mon.NORMAL ? 'LPB' : 'NORMAL'} belum ada, total bulan ini tidak lengkap`);
        }
        const months = Object.keys(u.months).sort();
        const jen = [...new Set(Object.values(u.months).flatMap((m) => ['NORMAL', 'LPB', 'TOTAL'].filter((j) => m[j])))].sort();
        log.push({ level, unit: uname, months: months.length, from: months[0], to: months[months.length - 1], jenis: jen });
      }
    }

    // cek silang: jumlah unit UPI == UP per bulan
    const up = out.levels.UP || {}, upi = out.levels.UPI || {};
    for (const [upName, upd] of Object.entries(up)) {
      const kids = Object.values(upi).filter((u) => u.parent === upName);
      if (!kids.length) continue;
      for (const mk of Object.keys(upd.months).sort()) {
        if (!kids.every((k) => k.months[mk])) continue;
        const mon = upd.months[mk];
        const sumC = (rows) => Object.values(rows || {}).reduce((a, r) => a + r[0], 0);
        const upC = (sumC(mon.NORMAL) + sumC(mon.LPB)) || sumC(mon.TOTAL);
        const kidC = kids.reduce((a, k) => a + sumC(k.months[mk].TOTAL), 0);
        if (Math.abs(upC - kidC) > 0.5) warns.push(`Cek silang ${mk}: pelanggan ${upName} ${fmt0(upC)} != jumlah ${kids.length} unit UPI ${fmt0(kidC)}`);
      }
    }
    const unnamed = [...new Set(Object.values(out.levels).flatMap((lv) => Object.keys(lv).filter((u) => u.startsWith('Unit '))))].sort();
    return { data: out, log, warns, unknown, unnamed };
  }

  const api = { build, sheetRows, levelOf, LEVELS };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.TulBuilder = api;
})(typeof window !== 'undefined' ? window : globalThis);
