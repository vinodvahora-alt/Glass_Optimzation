import React, { useState, useMemo, useRef, useCallback } from "react";
import * as XLSX from "xlsx";

/* ================= OM GLASS TRADERS — CUTTING OPTIMIZER =================
   MaxRects nesting engine (BSSF / BLSF / BAF / Bottom-Left) with kerf,
   edge margin, rotation & grain control, remnant reuse, SVG layouts,
   stats dashboard, PNG export and printable cutting plan.
======================================================================== */

const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@500;600;700&family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;700&display=swap');
:root{
  --paper:#EDF0F2; --panel:#FFFFFF; --ink:#182430; --sub:#5B6B77;
  --line:#CBD4D9; --accent:#D64F1E; --accent-dk:#B23E12;
  --waste-bg:#F8E7E3; --waste-ln:#C0392B; --left-bg:#E1F2E6; --left-ln:#2E7D4F;
}
*{box-sizing:border-box}
.gco{background:var(--paper);color:var(--ink);font-family:'Inter',sans-serif;min-height:100vh;font-size:13px}
.gco h1,.gco h2,.gco .disp{font-family:'Barlow Condensed',sans-serif;letter-spacing:.02em}
.gco .mono{font-family:'JetBrains Mono',monospace}
.topbar{background:var(--ink);color:#F4F7F8;padding:10px 20px;display:flex;align-items:center;gap:16px;flex-wrap:wrap}
.topbar h1{margin:0;font-size:24px;font-weight:700;text-transform:uppercase}
.topbar .tag{font-size:11px;color:#9FB2BE;letter-spacing:.14em;text-transform:uppercase}
.tabs{display:flex;gap:2px;margin-left:auto}
.tabbtn{font-family:'Barlow Condensed',sans-serif;font-size:15px;font-weight:600;letter-spacing:.06em;text-transform:uppercase;
  background:transparent;color:#AEBFC9;border:none;padding:8px 14px;cursor:pointer;border-bottom:3px solid transparent}
.tabbtn.on{color:#fff;border-bottom-color:var(--accent)}
.tabbtn:hover{color:#fff}
.wrap{padding:16px 20px;max-width:1400px;margin:0 auto}
.panel{background:var(--panel);border:1px solid var(--line);border-radius:6px;padding:14px;margin-bottom:14px}
.panel h2{margin:0 0 10px;font-size:18px;font-weight:700;text-transform:uppercase;color:var(--ink);
  border-bottom:2px solid var(--ink);padding-bottom:6px;display:flex;align-items:center;gap:8px}
.panel h2 .n{background:var(--accent);color:#fff;font-size:12px;padding:1px 7px;border-radius:3px;font-family:'JetBrains Mono',monospace}
table.grid{width:100%;border-collapse:collapse;font-size:12.5px}
table.grid th{font-family:'Barlow Condensed',sans-serif;font-size:13px;font-weight:600;letter-spacing:.06em;text-transform:uppercase;
  text-align:left;color:var(--sub);border-bottom:2px solid var(--ink);padding:5px 8px;white-space:nowrap}
table.grid td{border-bottom:1px solid var(--line);padding:4px 8px;vertical-align:middle}
table.grid tr:hover td{background:#F4F7F8}
.num{font-family:'JetBrains Mono',monospace;text-align:right}
input.f,select.f{border:1px solid var(--line);border-radius:4px;padding:5px 7px;font:inherit;background:#fff;color:var(--ink);width:100%}
input.f:focus,select.f:focus{outline:2px solid var(--accent);outline-offset:-1px;border-color:var(--accent)}
input.f.sm{width:80px}
.btn{font-family:'Barlow Condensed',sans-serif;font-size:14px;font-weight:600;letter-spacing:.06em;text-transform:uppercase;
  border:1px solid var(--ink);background:#fff;color:var(--ink);border-radius:4px;padding:7px 14px;cursor:pointer}
.btn:hover{background:var(--ink);color:#fff}
.btn.pri{background:var(--accent);border-color:var(--accent);color:#fff}
.btn.pri:hover{background:var(--accent-dk);border-color:var(--accent-dk)}
.btn.big{font-size:18px;padding:10px 26px}
.btn.ghost{border-color:var(--line);color:var(--sub)}
.btn.ghost:hover{background:#EEF2F4;color:var(--ink)}
.btn:disabled{opacity:.45;cursor:not-allowed}
.btn.danger{border-color:#C0392B;color:#C0392B}
.btn.danger:hover{background:#C0392B;color:#fff}
.rowbtns{display:flex;gap:6px}
.statgrid{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:10px}
.stat{border:1px solid var(--line);border-left:4px solid var(--ink);border-radius:4px;padding:8px 10px;background:#fff}
.stat .v{font-family:'JetBrains Mono',monospace;font-size:20px;font-weight:700}
.stat .l{font-family:'Barlow Condensed',sans-serif;font-size:12px;font-weight:600;letter-spacing:.08em;text-transform:uppercase;color:var(--sub)}
.stat.good{border-left-color:var(--left-ln)} .stat.bad{border-left-color:var(--waste-ln)} .stat.acc{border-left-color:var(--accent)}
.sheetcard{border:1px solid var(--line);border-radius:6px;background:#fff;margin-bottom:16px;overflow:hidden}
.titleblock{display:flex;flex-wrap:wrap;border-bottom:2px solid var(--ink);background:#F7F9FA}
.titleblock .cell{padding:6px 12px;border-right:1px solid var(--line)}
.titleblock .cell .k{font-family:'Barlow Condensed',sans-serif;font-size:10.5px;font-weight:600;letter-spacing:.1em;text-transform:uppercase;color:var(--sub)}
.titleblock .cell .x{font-family:'JetBrains Mono',monospace;font-size:14px;font-weight:700}
.svgbox{position:relative;background:#FBFCFC;overflow:hidden;cursor:grab}
.svgbox:active{cursor:grabbing}
.zoombtns{position:absolute;top:8px;right:8px;display:flex;gap:4px;z-index:5}
.zoombtns button{width:28px;height:28px;border:1px solid var(--ink);background:#fff;border-radius:4px;cursor:pointer;font-weight:700}
.pc{transition:opacity .1s}
.pc:hover{opacity:.82;cursor:pointer}
.pc.sel rect{stroke:#000;stroke-width:2.5}
.legend{display:flex;flex-wrap:wrap;gap:10px;padding:8px 12px;border-top:1px solid var(--line);font-size:11.5px}
.legend .it{display:flex;align-items:center;gap:5px}
.legend .sw{width:14px;height:14px;border:1px solid rgba(0,0,0,.35);border-radius:2px}
.detail{position:sticky;top:10px}
.kv{display:grid;grid-template-columns:auto 1fr;gap:3px 12px;font-size:12.5px}
.kv .k{color:var(--sub)} .kv .x{font-family:'JetBrains Mono',monospace;font-weight:600}
.settings{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:10px}
.setting label{display:block;font-family:'Barlow Condensed',sans-serif;font-size:12px;font-weight:600;
  letter-spacing:.08em;text-transform:uppercase;color:var(--sub);margin-bottom:3px}
.chip{display:inline-block;font-family:'JetBrains Mono',monospace;font-size:11px;background:#EEF2F4;
  border:1px solid var(--line);border-radius:3px;padding:1px 6px}
.warn{background:#FDF3E7;border:1px solid #E8B25C;border-radius:4px;padding:8px 12px;margin:8px 0;font-size:12.5px}
.imp{border:2px dashed var(--line);border-radius:6px;padding:14px;text-align:center;color:var(--sub)}
textarea.f{width:100%;min-height:90px;font-family:'JetBrains Mono',monospace;font-size:12px;border:1px solid var(--line);border-radius:4px;padding:8px}
@media print{
  .noprint{display:none!important}
  .gco{background:#fff}
  .sheetcard{page-break-inside:avoid;border:1px solid #000}
  .wrap{max-width:none;padding:0}
}
`;

/* ---------------- geometry / MaxRects ---------------- */
const HEURISTICS = {
  "G-AUTO": "Guillotine Auto", "G-V": "Guillotine V-Strips", "G-H": "Guillotine H-Strips",
  BSSF: "MaxRects Short Side", BLSF: "MaxRects Long Side", BAF: "MaxRects Area Fit", BL: "MaxRects Bottom-Left",
};

function scoreFit(fr, pw, ph, h) {
  const dw = fr.w - pw, dh = fr.h - ph;
  const ss = Math.min(dw, dh), ls = Math.max(dw, dh);
  if (h === "BSSF") return ss * 100000 + ls;
  if (h === "BLSF") return ls * 100000 + ss;
  if (h === "BAF") return (fr.w * fr.h - pw * ph) * 1000 + ss;
  return fr.y * 100000 + fr.x; // BL
}

function splitFree(free, u) {
  const out = [];
  for (const f of free) {
    if (u.x >= f.x + f.w || u.x + u.w <= f.x || u.y >= f.y + f.h || u.y + u.h <= f.y) { out.push(f); continue; }
    if (u.x > f.x) out.push({ x: f.x, y: f.y, w: u.x - f.x, h: f.h });
    if (u.x + u.w < f.x + f.w) out.push({ x: u.x + u.w, y: f.y, w: f.x + f.w - (u.x + u.w), h: f.h });
    if (u.y > f.y) out.push({ x: f.x, y: f.y, w: f.w, h: u.y - f.y });
    if (u.y + u.h < f.y + f.h) out.push({ x: f.x, y: u.y + u.h, w: f.w, h: f.y + f.h - (u.y + u.h) });
  }
  // prune contained rects
  const pr = [];
  for (let i = 0; i < out.length; i++) {
    const a = out[i]; let cont = false;
    if (a.w < 1 || a.h < 1) continue;
    for (let j = 0; j < out.length; j++) {
      if (i === j) continue; const b = out[j];
      if (a.x >= b.x && a.y >= b.y && a.x + a.w <= b.x + b.w && a.y + a.h <= b.y + b.h &&
          (b.w > a.w || b.h > a.h || j < i)) { cont = true; break; }
    }
    if (!cont) pr.push(a);
  }
  return pr;
}

/* Guillotine strip packing — every cut runs edge-to-edge (score & snap).
   dir 'V': vertical full-length rip cuts create strips along X, pieces
   stacked inside each strip with cross cuts. dir 'H': same, rotated. */
function packGuillotine(W, H, parts, opt, dir) {
  const orders = [
    a => a, // priority + area desc (presorted)
    a => [...a].sort((p, q) => Math.max(q.l, q.w) - Math.max(p.l, p.w) || q.l * q.w - p.l * p.w),
    a => [...a].sort((p, q) => q.w - p.w || q.l - p.l),
    a => [...a].sort((p, q) => q.l - p.l || q.w - p.w),
  ];
  let best = null;
  const consider = r => {
    const area = r.placed.reduce((s, p) => s + p.w * p.h, 0);
    if (!best || r.placed.length > best.n || (r.placed.length === best.n && area > best.area))
      best = { n: r.placed.length, area, r };
  };
  for (const ord of orders) consider(packGuillotineOnce(W, H, ord(parts), opt, dir));
  consider(packShelfOnce(W, H, parts, opt, dir, true));
  consider(packShelfOnce(W, H, parts, opt, dir, false));
  return best.r;
}

/* Shelf / level guillotine: full-length rip cuts create levels, pieces sit
   side by side inside each level, each trimmed to size. 3-stage snappable. */
function packShelfOnce(W, H, parts, opt, dir, tall) {
  const uW = W - 2 * opt.margin, uH = H - 2 * opt.margin;
  const AW = dir === "H" ? uW : uH;  // along the level
  const AH = dir === "H" ? uH : uW;  // level stacking direction
  const k = opt.kerf;
  const items = parts.map(p => {
    const canR = opt.rotation && p.rotate !== false && !opt.grain;
    const d1 = dir === "H" ? { pw: p.l, ph: p.w, rot: false } : { pw: p.w, ph: p.l, rot: false };
    const d2 = dir === "H" ? { pw: p.w, ph: p.l, rot: true } : { pw: p.l, ph: p.w, rot: true };
    const o = !canR ? d1 : ((tall ? d1.ph >= d2.ph : d1.ph <= d2.ph) ? d1 : d2);
    return { p, ...o, canR };
  }).sort((a, b) => b.ph - a.ph || b.pw - a.pw);
  const shelves = []; const placed = [], remaining = [], free = [];
  let yNext = 0;
  const put = (sh, pw, ph, rot, p) => {
    const px = sh.x, py = sh.y;
    placed.push(dir === "H" ? { x: px, y: py, w: pw, h: ph, rot, part: p } : { x: py, y: px, w: ph, h: pw, rot, part: p });
    if (sh.h - ph > k + 1) { // trim sliver above the piece within the level
      const s = { x: px, y: py + ph + k, w: pw, h: sh.h - ph - k };
      free.push(dir === "H" ? s : { x: s.y, y: s.x, w: s.h, h: s.w });
    }
    sh.x += pw + k;
  };
  for (const it of items) {
    let best = null;
    const orientsIn = it.canR ? [[it.pw, it.ph, it.rot], [it.ph, it.pw, !it.rot]] : [[it.pw, it.ph, it.rot]];
    for (const sh of shelves)
      for (const [pw, ph, rot] of orientsIn)
        if (ph <= sh.h + 1e-6 && sh.x + pw <= AW + 1e-6) {
          const wh = sh.h - ph;
          if (!best || wh < best.wh) best = { sh, pw, ph, rot, wh };
        }
    if (best) { put(best.sh, best.pw, best.ph, best.rot, it.p); continue; }
    // open a new level
    let opened = false;
    for (const [pw, ph, rot] of orientsIn) {
      const y = yNext + (shelves.length ? k : 0);
      if (y + ph <= AH + 1e-6 && pw <= AW + 1e-6) {
        const sh = { x: 0, y, h: ph };
        shelves.push(sh); yNext = y + ph;
        put(sh, pw, ph, rot, it.p);
        opened = true; break;
      }
    }
    if (!opened) remaining.push(it.p);
  }
  for (const sh of shelves) if (AW - sh.x > 1) {
    const s = { x: sh.x, y: sh.y, w: AW - sh.x, h: sh.h };
    free.push(dir === "H" ? s : { x: s.y, y: s.x, w: s.h, h: s.w });
  }
  if (AH - yNext > 1) {
    const s = { x: 0, y: yNext + k, w: AW, h: Math.max(0, AH - yNext - k) };
    if (s.h > 1) free.push(dir === "H" ? s : { x: s.y, y: s.x, w: s.h, h: s.w });
  }
  const strips = shelves.slice(0, -1).map(sh => sh.y + sh.h + k / 2)
    .concat(AH - yNext > k + 1 && shelves.length ? [yNext + k / 2] : []);
  return { placed, remaining, free, usableW: uW, usableH: uH, strips, dir: dir === "H" ? "H" : "V" };
}

function packGuillotineOnce(W, H, parts, opt, dir) {
  const uW = W - 2 * opt.margin, uH = H - 2 * opt.margin;
  const SW = dir === "V" ? uW : uH;   // strip-advance axis
  const SH = dir === "V" ? uH : uW;   // stack axis
  const k = opt.kerf;
  let rem = parts.slice();
  const placed = [], free = [], strips = [];
  let x = 0;

  const orients = p => {
    const canR = opt.rotation && p.rotate !== false && !opt.grain;
    const o1 = dir === "V" ? { a: p.l, b: p.w, rot: false } : { a: p.w, b: p.l, rot: false };
    return canR ? [o1, dir === "V" ? { a: p.w, b: p.l, rot: true } : { a: p.l, b: p.w, rot: true }] : [o1];
  };

  let sGuard = 0;
  while (rem.length && x < SW - 1 && sGuard++ < 2000) {
    // seed = largest remaining part (list is presorted); orientation that wastes least strip area
    let seed = null;
    for (let i = 0; i < rem.length && !seed; i++) {
      const p = rem[i];
      const twins = rem.filter(q => q.name === p.name && q.l === p.l && q.w === p.w).length;
      let best = null;
      for (const o of orients(p)) {
        if (x + o.a > SW + 1e-6 || o.b > SH + 1e-6) continue;
        const per = Math.floor((SH + k) / (o.b + k));
        if (per < 1) continue;
        const eff = (Math.min(twins, per) * o.b) / SH; // stacked coverage of the strip
        if (!best || eff > best.eff) best = { ...o, eff };
      }
      if (best) seed = best;
    }
    if (!seed) break;

    const stripW = seed.a;
    let y = 0, fGuard = 0;
    while (fGuard++ < 10000) {
      // best filler for this strip: exact-width match first (clean cross cuts), then largest area
      let pick = null, pi = -1;
      for (let i = 0; i < rem.length; i++) {
        for (const o of orients(rem[i])) {
          if (o.a > stripW + 1e-6 || y + o.b > SH + 1e-6) continue;
          const score = (Math.abs(o.a - stripW) < 1e-6 ? 1e12 : 0) + rem[i].l * rem[i].w * 1e3 + o.a;
          if (!pick || score > pick.score) { pick = { ...o, score }; pi = i; }
        }
      }
      if (!pick) break;
      const p = rem[pi];
      placed.push(dir === "V"
        ? { x, y, w: pick.a, h: pick.b, rot: pick.rot, part: p }
        : { x: y, y: x, w: pick.b, h: pick.a, rot: pick.rot, part: p });
      if (stripW - pick.a > k + 1) { // side sliver inside the strip
        const sx = x + pick.a + k, sw = stripW - pick.a - k;
        free.push(dir === "V" ? { x: sx, y, w: sw, h: pick.b } : { x: y, y: sx, w: pick.b, h: sw });
      }
      rem.splice(pi, 1);
      y += pick.b + k;
    }
    if (SH - y > 1) // top-of-strip remainder
      free.push(dir === "V" ? { x, y, w: stripW, h: SH - y } : { x: y, y: x, w: SH - y, h: stripW });
    if (x + stripW < SW - 1) strips.push(x + stripW + k / 2); // full-length rip cut position
    x += stripW + k;
  }
  if (SW - x > 1)
    free.push(dir === "V" ? { x, y: 0, w: SW - x, h: SH } : { x: 0, y: x, w: SH, h: SW - x });
  return { placed, remaining: rem, free, usableW: uW, usableH: uH, strips, dir };
}

// Pack a list of part instances onto one sheet. Returns placed, remaining, free rects.
function packOneSheet(W, H, parts, opt, heur) {
  if (heur === "G-V" || heur === "G-H") return packGuillotine(W, H, parts, opt, heur === "G-V" ? "V" : "H");
  const usableW = W - 2 * opt.margin, usableH = H - 2 * opt.margin;
  let free = [{ x: 0, y: 0, w: usableW, h: usableH }];
  const placed = [], remaining = [];
  for (const p of parts) {
    const canRot = opt.rotation && p.rotate !== false && !opt.grain;
    let best = null;
    for (const fr of free) {
      for (const rot of canRot ? [false, true] : [false]) {
        const pw = rot ? p.w : p.l, ph = rot ? p.l : p.w; // l along X, w along Y
        if (pw <= fr.w + 1e-6 && ph <= fr.h + 1e-6) {
          const s = scoreFit(fr, pw, ph, heur);
          if (!best || s < best.s) best = { s, x: fr.x, y: fr.y, pw, ph, rot };
        }
      }
    }
    if (!best) { remaining.push(p); continue; }
    placed.push({ x: best.x, y: best.y, w: best.pw, h: best.ph, rot: best.rot, part: p });
    const used = {
      x: best.x, y: best.y,
      w: Math.min(best.pw + opt.kerf, usableW - best.x),
      h: Math.min(best.ph + opt.kerf, usableH - best.y),
    };
    free = splitFree(free, used);
  }
  return { placed, remaining, free, usableW, usableH };
}

// Greedy non-overlapping leftover selection from free rects
function pickLeftovers(free, opt) {
  const cand = free
    .filter(f => Math.min(f.w, f.h) >= opt.minLeftSide && f.w * f.h >= opt.minLeftArea * 1e6)
    .sort((a, b) => b.w * b.h - a.w * a.h);
  const sel = [];
  for (const c of cand) {
    let ov = false;
    for (const s of sel)
      if (!(c.x >= s.x + s.w || c.x + c.w <= s.x || c.y >= s.y + s.h || c.y + c.h <= s.y)) { ov = true; break; }
    if (!ov) sel.push(c);
  }
  return sel;
}

/* ---------------- multi-sheet optimizer ---------------- */
function optimize(stocks, parts, remnants, opt) {
  // expand parts by qty, sort: priority asc, area desc
  const inst = [];
  parts.forEach((p, pi) => {
    for (let i = 0; i < p.qty; i++)
      inst.push({ ...p, uid: pi + "_" + i });
  });
  inst.sort((a, b) => (a.priority - b.priority) || (b.l * b.w - a.l * a.w) || (Math.max(b.l, b.w) - Math.max(a.l, a.w)));

  // group by glass type + thickness
  const groups = {};
  for (const p of inst) {
    const k = (p.glass || "ANY") + "|" + (p.thk || "ANY");
    (groups[k] = groups[k] || []).push(p);
  }

  const stockLeft = stocks.map(s => ({ ...s, left: s.qty }));
  const remLeft = remnants.map(r => ({ ...r, used: false }));
  const heurs =
    opt.heuristic === "AUTO" ? ["G-V", "G-H", "BSSF", "BAF"] :
    opt.heuristic === "G-AUTO" ? ["G-V", "G-H"] : [opt.heuristic];
  const sheets = []; const unplaced = [];

  const matches = (src, g, t) =>
    (!g || g === "ANY" || !src.glass || src.glass.toLowerCase() === g.toLowerCase()) &&
    (!t || t === "ANY" || !src.thk || +src.thk === +t);

  for (const key of Object.keys(groups)) {
    let rem = groups[key];
    const [g, t] = key.split("|");

    // 1) consume matching remnants first (smallest usable first)
    if (opt.useRemnants) {
      const rc = remLeft.filter(r => !r.used && matches(r, g, t)).sort((a, b) => a.l * a.w - b.l * b.w);
      for (const r of rc) {
        if (!rem.length) break;
        let best = null;
        for (const h of heurs) {
          const res = packOneSheet(r.l, r.w, rem, opt, h);
          if (res.placed.length && (!best || res.placed.length > best.res.placed.length)) best = { res, h };
        }
        if (best && best.res.placed.length) {
          r.used = true;
          sheets.push({ src: "REMNANT", id: r.id, W: r.l, H: r.w, cost: 0, glass: r.glass, thk: r.thk, ...best.res, heur: best.h });
          rem = best.res.remaining;
        }
      }
    }

    // 2) full stock sheets — per sheet pick the stock size giving best utilization
    let guard = 0;
    while (rem.length && guard++ < 5000) {
      const cands = stockLeft.filter(s => s.left > 0 && matches(s, g, t));
      if (!cands.length) { unplaced.push(...rem); break; }
      let best = null;
      for (const s of cands) for (const h of heurs) {
        const res = packOneSheet(s.l, s.w, rem, opt, h);
        if (!res.placed.length) continue;
        const areaPlaced = res.placed.reduce((a, p) => a + p.w * p.h, 0);
        const util = areaPlaced / (s.l * s.w);
        // prefer: most parts placed, then utilization, then cheaper sheet
        const score = res.placed.length * 1e9 + util * 1e6 - s.cost * 0.001;
        if (!best || score > best.score) best = { score, s, h, res };
      }
      if (!best) { unplaced.push(...rem); break; }
      best.s.left -= 1;
      sheets.push({ src: "STOCK", id: best.s.id, W: best.s.l, H: best.s.w, cost: best.s.cost, glass: best.s.glass, thk: best.s.thk, ...best.res, heur: best.h });
      rem = best.res.remaining;
    }
  }

  // consolidation: repeatedly try to merge the two least-filled stock sheets of a glass group
  let merged = true, cGuard = 0;
  while (merged && cGuard++ < 60) {
    merged = false;
    const byG = {};
    sheets.forEach((s, i) => { if (s.src === "STOCK") (byG[(s.glass || "") + "|" + (s.thk || "")] = byG[(s.glass || "") + "|" + (s.thk || "")] || []).push(i); });
    for (const key of Object.keys(byG)) {
      const idxs = byG[key];
      if (idxs.length < 2) continue;
      const u = i => sheets[i].placed.reduce((a, p) => a + p.w * p.h, 0) / (sheets[i].W * sheets[i].H);
      const [i1, i2] = [...idxs].sort((a, b) => u(a) - u(b)).slice(0, 2);
      const pool = [...sheets[i1].placed, ...sheets[i2].placed].map(p => p.part)
        .sort((a, b) => (a.priority - b.priority) || (b.l * b.w - a.l * a.w));
      const [g, t] = key.split("|");
      const cands = stockLeft.filter(s =>
        matches(s, g, t) && (s.left > 0 || s.id === sheets[i1].id || s.id === sheets[i2].id));
      let hit = null;
      for (const s of cands) {
        for (const h of heurs) {
          const res = packOneSheet(s.l, s.w, pool, opt, h);
          if (res.remaining.length === 0) {
            const areaP = res.placed.reduce((a, p) => a + p.w * p.h, 0);
            if (!hit || areaP / (s.l * s.w) > hit.util) hit = { s, h, res, util: areaP / (s.l * s.w) };
          }
        }
      }
      if (hit) {
        const back1 = stockLeft.find(s => s.id === sheets[i1].id); if (back1) back1.left++;
        const back2 = stockLeft.find(s => s.id === sheets[i2].id); if (back2) back2.left++;
        hit.s.left--;
        const nsheet = { src: "STOCK", id: hit.s.id, W: hit.s.l, H: hit.s.w, cost: hit.s.cost, glass: hit.s.glass, thk: hit.s.thk, ...hit.res, heur: hit.h };
        for (const i of [i1, i2].sort((a, b) => b - a)) sheets.splice(i, 1);
        sheets.push(nsheet);
        merged = true;
        break;
      }
    }
  }

  // leftovers + waste per sheet
  sheets.forEach((sh, i) => {
    sh.no = i + 1;
    sh.leftovers = pickLeftovers(sh.free, opt);
    sh.areaParts = sh.placed.reduce((a, p) => a + p.w * p.h, 0);
    sh.areaLeft = sh.leftovers.reduce((a, r) => a + r.w * r.h, 0);
    sh.areaSheet = sh.W * sh.H;
    sh.util = sh.areaParts / sh.areaSheet;
  });
  return { sheets, unplaced, stockLeft, remLeft };
}

/* ---------------- helpers ---------------- */
const PART_COLORS = ["#3E7CB1", "#59A96A", "#E7B10A", "#8E6FB6", "#D97C9E", "#4FB0AE", "#B0713E", "#6C8AE4", "#A3B65C", "#C96B6B", "#5F9EA0", "#B88AC9"];
const fmt = (n, d = 0) => Number(n).toLocaleString("en-IN", { maximumFractionDigits: d, minimumFractionDigits: 0 });
const m2 = a => (a / 1e6).toFixed(2);
const inr = n => "₹" + fmt(Math.round(n));
const today = () => new Date().toISOString().slice(0, 10);

const DEMO_STOCKS = [
  { id: "S01", glass: "Clear Float", thk: 5, colour: "Clear", l: 2440, w: 1830, qty: 20, cost: 3800, supplier: "Saint-Gobain", loc: "Rack A1", rem: "" },
  { id: "S02", glass: "Clear Float", thk: 5, colour: "Clear", l: 3300, w: 2140, qty: 12, cost: 6900, supplier: "AIS", loc: "Rack A2", rem: "" },
  { id: "S03", glass: "Clear Float", thk: 5, colour: "Clear", l: 3660, w: 2440, qty: 5, cost: 9200, supplier: "AIS", loc: "Rack B1", rem: "Jumbo" },
];
const DEMO_PARTS = [
  { name: "A", l: 850, w: 450, qty: 12, glass: "Clear Float", thk: 5, rotate: true, priority: 1, label: "Window Fix" },
  { name: "B", l: 1200, w: 650, qty: 8, glass: "Clear Float", thk: 5, rotate: true, priority: 1, label: "Door Panel" },
  { name: "C", l: 600, w: 500, qty: 20, glass: "Clear Float", thk: 5, rotate: true, priority: 2, label: "Vent" },
];

/* ---------------- Sheet layout (SVG with zoom/pan) ---------------- */
function SheetSVG({ sh, colors, sel, onSel }) {
  const [view, setView] = useState({ s: 1, tx: 0, ty: 0 });
  const drag = useRef(null);
  const M = sh.W > 3000 ? 60 : 40; // outer margin in mm-space for dims
  const vb = `${-M} ${-M} ${sh.W + 2 * M} ${sh.H + 2 * M}`;
  const fs = Math.max(sh.W, sh.H) / 55; // base font size in mm-space

  const onWheel = e => {
    e.preventDefault();
    setView(v => ({ ...v, s: Math.min(8, Math.max(0.5, v.s * (e.deltaY < 0 ? 1.15 : 0.87))) }));
  };
  const onDown = e => { drag.current = { x: e.clientX, y: e.clientY, tx: view.tx, ty: view.ty }; };
  const onMove = e => {
    if (!drag.current) return;
    setView(v => ({ ...v, tx: drag.current.tx + (e.clientX - drag.current.x), ty: drag.current.ty + (e.clientY - drag.current.y) }));
  };
  const onUp = () => { drag.current = null; };

  return (
    <div className="svgbox" onWheel={onWheel} onMouseDown={onDown} onMouseMove={onMove} onMouseUp={onUp} onMouseLeave={onUp}>
      <div className="zoombtns noprint">
        <button onClick={() => setView(v => ({ ...v, s: Math.min(8, v.s * 1.25) }))}>+</button>
        <button onClick={() => setView(v => ({ ...v, s: Math.max(0.5, v.s / 1.25) }))}>−</button>
        <button onClick={() => setView({ s: 1, tx: 0, ty: 0 })} title="Reset">⟲</button>
      </div>
      <div style={{ transform: `translate(${view.tx}px,${view.ty}px) scale(${view.s})`, transformOrigin: "50% 50%" }}>
        <svg viewBox={vb} style={{ width: "100%", display: "block", maxHeight: 520 }} xmlns="http://www.w3.org/2000/svg">
          <defs>
            <pattern id={"hatch" + sh.no} width="18" height="18" patternTransform="rotate(45)" patternUnits="userSpaceOnUse">
              <rect width="18" height="18" fill="#F8E7E3" />
              <line x1="0" y1="0" x2="0" y2="18" stroke="#C0392B" strokeWidth="3" opacity="0.45" />
            </pattern>
          </defs>
          {/* sheet: waste hatch as base */}
          <rect x="0" y="0" width={sh.W} height={sh.H} fill={"url(#hatch" + sh.no + ")"} stroke="#182430" strokeWidth={fs / 5} />
          {/* dims */}
          <text x={sh.W / 2} y={-M / 3} textAnchor="middle" fontFamily="'JetBrains Mono',monospace" fontWeight="700" fontSize={fs}>{sh.W} mm</text>
          <text x={-M / 3} y={sh.H / 2} textAnchor="middle" fontFamily="'JetBrains Mono',monospace" fontWeight="700" fontSize={fs}
            transform={`rotate(-90 ${-M / 3} ${sh.H / 2})`}>{sh.H} mm</text>
          {/* leftovers */}
          {sh.leftovers.map((r, i) => (
            <g key={"L" + i}>
              <rect x={r.x + (sh.mgn || 0)} y={r.y + (sh.mgn || 0)} width={r.w} height={r.h}
                fill="#E1F2E6" stroke="#2E7D4F" strokeWidth={fs / 8} strokeDasharray={`${fs / 2},${fs / 3}`} />
              <text x={r.x + (sh.mgn || 0) + r.w / 2} y={r.y + (sh.mgn || 0) + r.h / 2} textAnchor="middle" dominantBaseline="middle"
                fontFamily="'JetBrains Mono',monospace" fontSize={fs * 0.75} fill="#2E7D4F" fontWeight="600">
                {Math.round(r.w)}×{Math.round(r.h)}
              </text>
            </g>
          ))}
          {/* full-length rip cuts (guillotine) */}
          {sh.strips && sh.strips.map((pos, i) => {
            const c = pos + (sh.mgn || 0);
            return sh.dir === "V" ? (
              <g key={"C" + i}>
                <line x1={c} y1={0} x2={c} y2={sh.H} stroke="#14202B" strokeWidth={fs / 7} strokeDasharray={`${fs},${fs * 0.6}`} />
                <text x={c} y={sh.H + M * 0.55} textAnchor="middle" fontFamily="'JetBrains Mono',monospace" fontSize={fs * 0.75} fill="#5B6B77">C{i + 1}</text>
              </g>
            ) : (
              <g key={"C" + i}>
                <line x1={0} y1={c} x2={sh.W} y2={c} stroke="#14202B" strokeWidth={fs / 7} strokeDasharray={`${fs},${fs * 0.6}`} />
                <text x={sh.W + M * 0.35} y={c} dominantBaseline="middle" fontFamily="'JetBrains Mono',monospace" fontSize={fs * 0.75} fill="#5B6B77">C{i + 1}</text>
              </g>
            );
          })}
          {/* placed pieces */}
          {sh.placed.map((p, i) => {
            const x = p.x + (sh.mgn || 0), y = p.y + (sh.mgn || 0);
            const c = colors[p.part.name] || "#888";
            const isSel = sel && sel.sheet === sh.no && sel.idx === i;
            const small = Math.min(p.w, p.h) < fs * 3.2;
            return (
              <g key={i} className={"pc" + (isSel ? " sel" : "")} onClick={e => { e.stopPropagation(); onSel({ sheet: sh.no, idx: i, p, sh }); }}>
                <rect x={x} y={y} width={p.w} height={p.h} fill={c} fillOpacity="0.85" stroke="#14202B" strokeWidth={fs / 10} />
                <circle cx={x + fs * 1.1} cy={y + fs * 1.1} r={fs * 0.85} fill="#14202B" />
                <text x={x + fs * 1.1} y={y + fs * 1.1} textAnchor="middle" dominantBaseline="central"
                  fontFamily="'JetBrains Mono',monospace" fontWeight="700" fontSize={fs * 0.8} fill="#fff">{i + 1}</text>
                {!small && <>
                  <text x={x + p.w / 2} y={y + p.h / 2 - fs * 0.25} textAnchor="middle" fontFamily="'Barlow Condensed',sans-serif"
                    fontWeight="700" fontSize={fs * 1.15} fill="#0E1822">{p.part.name}{p.rot ? " ⟳" : ""}</text>
                  <text x={x + p.w / 2} y={y + p.h / 2 + fs * 0.95} textAnchor="middle" fontFamily="'JetBrains Mono',monospace"
                    fontSize={fs * 0.8} fill="#0E1822">{Math.round(p.w)}×{Math.round(p.h)}</text>
                </>}
              </g>
            );
          })}
        </svg>
      </div>
    </div>
  );
}

/* Standalone SVG markup for a sheet (independent of the mounted DOM) */
function sheetToSVGString(sh, colors, opt) {
  const M = sh.W > 3000 ? 60 : 40, fs = Math.max(sh.W, sh.H) / 55, mgn = sh.mgn || 0;
  const e = [];
  e.push(`<defs><pattern id="h" width="18" height="18" patternTransform="rotate(45)" patternUnits="userSpaceOnUse"><rect width="18" height="18" fill="#F8E7E3"/><line x1="0" y1="0" x2="0" y2="18" stroke="#C0392B" stroke-width="3" opacity="0.45"/></pattern></defs>`);
  e.push(`<rect x="0" y="0" width="${sh.W}" height="${sh.H}" fill="url(#h)" stroke="#182430" stroke-width="${fs / 5}"/>`);
  e.push(`<text x="${sh.W / 2}" y="${-M / 3}" text-anchor="middle" font-family="monospace" font-weight="700" font-size="${fs}">${sh.W} mm</text>`);
  e.push(`<text x="${-M / 3}" y="${sh.H / 2}" text-anchor="middle" font-family="monospace" font-weight="700" font-size="${fs}" transform="rotate(-90 ${-M / 3} ${sh.H / 2})">${sh.H} mm</text>`);
  for (const r of sh.leftovers || []) {
    e.push(`<rect x="${r.x + mgn}" y="${r.y + mgn}" width="${r.w}" height="${r.h}" fill="#E1F2E6" stroke="#2E7D4F" stroke-width="${fs / 8}" stroke-dasharray="${fs / 2},${fs / 3}"/>`);
    e.push(`<text x="${r.x + mgn + r.w / 2}" y="${r.y + mgn + r.h / 2}" text-anchor="middle" dominant-baseline="middle" font-family="monospace" font-size="${fs * 0.75}" fill="#2E7D4F" font-weight="600">${Math.round(r.w)}x${Math.round(r.h)}</text>`);
  }
  if (sh.strips) sh.strips.forEach((pos, i) => {
    const c = pos + mgn;
    if (sh.dir === "V") {
      e.push(`<line x1="${c}" y1="0" x2="${c}" y2="${sh.H}" stroke="#14202B" stroke-width="${fs / 7}" stroke-dasharray="${fs},${fs * 0.6}"/>`);
      e.push(`<text x="${c}" y="${sh.H + M * 0.55}" text-anchor="middle" font-family="monospace" font-size="${fs * 0.75}" fill="#5B6B77">C${i + 1}</text>`);
    } else {
      e.push(`<line x1="0" y1="${c}" x2="${sh.W}" y2="${c}" stroke="#14202B" stroke-width="${fs / 7}" stroke-dasharray="${fs},${fs * 0.6}"/>`);
      e.push(`<text x="${sh.W + M * 0.35}" y="${c}" dominant-baseline="middle" font-family="monospace" font-size="${fs * 0.75}" fill="#5B6B77">C${i + 1}</text>`);
    }
  });
  sh.placed.forEach((p, i) => {
    const x = p.x + mgn, y = p.y + mgn, c = colors[p.part.name] || "#888";
    const small = Math.min(p.w, p.h) < fs * 3.2;
    e.push(`<rect x="${x}" y="${y}" width="${p.w}" height="${p.h}" fill="${c}" fill-opacity="0.85" stroke="#14202B" stroke-width="${fs / 10}"/>`);
    e.push(`<circle cx="${x + fs * 1.1}" cy="${y + fs * 1.1}" r="${fs * 0.85}" fill="#14202B"/>`);
    e.push(`<text x="${x + fs * 1.1}" y="${y + fs * 1.1}" text-anchor="middle" dominant-baseline="central" font-family="monospace" font-weight="700" font-size="${fs * 0.8}" fill="#fff">${i + 1}</text>`);
    if (!small) {
      e.push(`<text x="${x + p.w / 2}" y="${y + p.h / 2 - fs * 0.25}" text-anchor="middle" font-family="sans-serif" font-weight="700" font-size="${fs * 1.15}" fill="#0E1822">${String(p.part.name).replace(/&/g, "&amp;").replace(/</g, "&lt;")}${p.rot ? " (R)" : ""}</text>`);
      e.push(`<text x="${x + p.w / 2}" y="${y + p.h / 2 + fs * 0.95}" text-anchor="middle" font-family="monospace" font-size="${fs * 0.8}" fill="#0E1822">${Math.round(p.w)}x${Math.round(p.h)}</text>`);
    }
  });
  return { xml: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${-M} ${-M} ${sh.W + 2 * M} ${sh.H + 2 * M}" width="${sh.W + 2 * M}" height="${sh.H + 2 * M}">${e.join("")}</svg>`, vw: sh.W + 2 * M, vh: sh.H + 2 * M };
}

function svgStringToCanvas(xml, vw, vh, pxWidth) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      const s = pxWidth / vw;
      const cv = document.createElement("canvas");
      cv.width = Math.round(vw * s); cv.height = Math.round(vh * s);
      const ctx = cv.getContext("2d");
      ctx.fillStyle = "#FFFFFF"; ctx.fillRect(0, 0, cv.width, cv.height);
      ctx.drawImage(img, 0, 0, cv.width, cv.height);
      resolve(cv);
    };
    img.onerror = () => reject(new Error("SVG render failed"));
    img.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(xml);
  });
}

function saveBlob(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = name;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 3000);
}

async function exportPNG(sh, colors, opt) {
  try {
    const { xml, vw, vh } = sheetToSVGString(sh, colors, opt);
    const cv = await svgStringToCanvas(xml, vw, vh, 2200);
    cv.toBlob(b => { if (b) saveBlob(b, `sheet-${sh.no}-${sh.id}.png`); }, "image/png");
  } catch (err) { alert("PNG export failed: " + err.message); }
}

/* Minimal PDF writer — A4 landscape, JPEG page images + Helvetica text. */
const pdfTxt = s => String(s).replace(/₹/g, "Rs ").replace(/[×✕]/g, "x").replace(/[^\x20-\x7E]/g, "").replace(/([()\\])/g, "\\$1");
function buildPDF(pages) {
  const enc = new TextEncoder();
  const chunks = []; const offsets = []; let off = 0;
  const push = d => { const b = typeof d === "string" ? enc.encode(d) : d; chunks.push(b); off += b.length; };
  const startObj = n => { offsets[n] = off; push(n + " 0 obj\n"); };
  const PW = 842, PH = 595, n = pages.length;
  push("%PDF-1.4\n");
  startObj(1); push("<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
  startObj(2); push(`<< /Type /Pages /Kids [${pages.map((_, i) => (4 + i * 3) + " 0 R").join(" ")}] /Count ${n} >>\nendobj\n`);
  startObj(3); push("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\n");
  pages.forEach((pg, i) => {
    const pnum = 4 + i * 3, cnum = pnum + 1, inum = pnum + 2;
    const res = pg.jpeg ? `/XObject << /Im0 ${inum} 0 R >> /Font << /F1 3 0 R >>` : `/Font << /F1 3 0 R >>`;
    startObj(pnum);
    push(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${PW} ${PH}] /Resources << ${res} >> /Contents ${cnum} 0 R >>\nendobj\n`);
    let c = `BT /F1 14 Tf 30 ${PH - 32} Td (${pdfTxt(pg.title)}) Tj ET\n`;
    if (pg.jpeg) {
      const availW = PW - 60, availH = PH - 90;
      const s = Math.min(availW / pg.pw, availH / pg.ph);
      const dw = pg.pw * s, dh = pg.ph * s;
      c += `q ${dw.toFixed(2)} 0 0 ${dh.toFixed(2)} ${((PW - dw) / 2).toFixed(2)} ${(((PH - 50 - dh) / 2)).toFixed(2)} cm /Im0 Do Q\n`;
    }
    if (pg.lines) {
      let y = PH - 70;
      for (const ln of pg.lines) { c += `BT /F1 11 Tf 40 ${y} Td (${pdfTxt(ln)}) Tj ET\n`; y -= 17; if (y < 30) break; }
    }
    startObj(cnum);
    push(`<< /Length ${c.length} >>\nstream\n${c}endstream\nendobj\n`);
    startObj(inum);
    if (pg.jpeg) {
      push(`<< /Type /XObject /Subtype /Image /Width ${pg.pw} /Height ${pg.ph} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${pg.jpeg.length} >>\nstream\n`);
      push(pg.jpeg);
      push("\nendstream\nendobj\n");
    } else push("<< >>\nendobj\n");
  });
  const xrefStart = off;
  const size = 4 + n * 3;
  let xr = `xref\n0 ${size}\n0000000000 65535 f \n`;
  for (let i = 1; i < size; i++) xr += String(offsets[i] || 0).padStart(10, "0") + " 00000 n \n";
  push(xr);
  push(`trailer\n<< /Size ${size} /Root 1 0 R >>\nstartxref\n${xrefStart}\n%%EOF`);
  return new Blob(chunks, { type: "application/pdf" });
}

async function exportPDFPlan(result, stats, colors, opt) {
  try {
    const pages = [];
    for (const sh of result.sheets) {
      const { xml, vw, vh } = sheetToSVGString(sh, colors, opt);
      const cv = await svgStringToCanvas(xml, vw, vh, 1600);
      const b64 = cv.toDataURL("image/jpeg", 0.88).split(",")[1];
      const bin = atob(b64);
      const bytes = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      pages.push({
        jpeg: bytes, pw: cv.width, ph: cv.height,
        title: `Sheet #${sh.no}  |  ${sh.src === "REMNANT" ? "Remnant" : "Stock"} ${sh.id}  |  ${sh.W} x ${sh.H} mm  |  ${sh.glass || ""} ${sh.thk || ""}mm  |  ${sh.placed.length} pcs  |  Yield ${(sh.util * 100).toFixed(1)}%`,
      });
    }
    pages.push({
      title: "Production Summary — " + today(),
      lines: [
        `Sheets used: ${stats.sheets}    Pieces placed: ${stats.nParts}`,
        `Stock area: ${m2(stats.areaSheet)} m2    Yield: ${stats.yieldPct.toFixed(1)}%    True waste: ${stats.wastePct.toFixed(1)}% (${m2(stats.waste)} m2)`,
        `Reusable leftovers: ${stats.nLeft} pcs, ${m2(stats.areaLeft)} m2`,
        `Material cost: Rs ${fmt(Math.round(stats.cost))}    Cost per piece: Rs ${fmt(Math.round(stats.costPer))}`,
        `Kerf: ${opt.kerf} mm    Edge margin: ${opt.margin} mm    Method: ${HEURISTICS[opt.heuristic] || opt.heuristic}`,
        ``,
        ...result.sheets.map(sh => `Sheet #${sh.no} (${sh.id}, ${sh.W}x${sh.H}): ` +
          Object.entries(sh.placed.reduce((m, p) => { m[p.part.name] = (m[p.part.name] || 0) + 1; return m; }, {}))
            .map(([k, v]) => `${k} x${v}`).join(", ")),
      ],
    });
    saveBlob(buildPDF(pages), "cutting-plan.pdf");
  } catch (err) { alert("PDF export failed: " + err.message); }
}

function downloadCSV(rows, name) {
  const csv = rows.map(r => r.map(c => `"${String(c ?? "").replace(/"/g, '""')}"`).join(",")).join("\n");
  saveBlob(new Blob([csv], { type: "text/csv" }), name + ".csv");
}

/* ---------------- main app ---------------- */
export default function GlassCuttingOptimizer() {
  const [tab, setTab] = useState("parts");
  const [stocks, setStocks] = useState(DEMO_STOCKS);
  const [parts, setParts] = useState(DEMO_PARTS);
  const [remnants, setRemnants] = useState([]);
  const [opt, setOpt] = useState({
    kerf: 4, margin: 10, rotation: true, grain: false, useRemnants: true,
    minLeftSide: 150, minLeftArea: 0.09, heuristic: "G-AUTO",
  });
  const [result, setResult] = useState(null);
  const [sel, setSel] = useState(null);
  const [remSaved, setRemSaved] = useState(false);
  const [runMs, setRunMs] = useState(0);
  const fileRef = useRef(null);
  const [paste, setPaste] = useState("");

  const colors = useMemo(() => {
    const m = {}; parts.forEach((p, i) => { m[p.name] = PART_COLORS[i % PART_COLORS.length]; }); return m;
  }, [parts]);

  const run = () => {
    const t0 = performance.now();
    const res = optimize(stocks, parts.filter(p => p.qty > 0 && p.l > 0 && p.w > 0), remnants, opt);
    res.sheets.forEach(sh => { sh.mgn = opt.margin; });
    setRunMs(Math.round(performance.now() - t0));
    setResult(res); setSel(null); setRemSaved(false); setTab("layout");
  };

  const stats = useMemo(() => {
    if (!result) return null;
    const S = result.sheets;
    const areaSheet = S.reduce((a, s) => a + s.areaSheet, 0);
    const areaParts = S.reduce((a, s) => a + s.areaParts, 0);
    const areaLeft = S.reduce((a, s) => a + s.areaLeft, 0);
    const cost = S.reduce((a, s) => a + (s.cost || 0), 0);
    const nParts = S.reduce((a, s) => a + s.placed.length, 0);
    const savings = areaLeft > 0 && areaSheet > 0 ? cost * (areaLeft / areaSheet) : 0;
    return {
      sheets: S.length, areaSheet, areaParts, areaLeft,
      waste: areaSheet - areaParts - areaLeft,
      wastePct: areaSheet ? ((areaSheet - areaParts - areaLeft) / areaSheet) * 100 : 0,
      yieldPct: areaSheet ? (areaParts / areaSheet) * 100 : 0,
      nLeft: S.reduce((a, s) => a + s.leftovers.length, 0),
      cost, nParts, costPer: nParts ? cost / nParts : 0, savings,
    };
  }, [result]);

  const saveLeftovers = () => {
    if (!result || remSaved) return;
    const add = [];
    result.sheets.forEach(sh => sh.leftovers.forEach((r, i) => add.push({
      id: `R${Date.now().toString(36).slice(-4).toUpperCase()}${sh.no}${i}`,
      l: Math.floor(r.w), w: Math.floor(r.h), glass: sh.glass, thk: sh.thk,
      date: today(), ref: `Sheet #${sh.no} (${sh.id})`,
    })));
    setRemnants(rs => [...rs.filter(r => !result.remLeft.find(u => u.used && u.id === r.id)), ...add]);
    setRemSaved(true);
  };

  const applyStockConsumption = () => {
    if (!result) return;
    setStocks(result.stockLeft.map(s => ({ ...s, qty: s.left })));
  };

  /* ---- import ---- */
  const mapRow = r => {
    const g = k => { for (const key of Object.keys(r)) if (key.toLowerCase().replace(/[^a-z]/g, "").includes(k)) return r[key]; };
    const name = g("part") ?? g("name") ?? "";
    const l = parseFloat(g("length") ?? g("len") ?? 0), w = parseFloat(g("width") ?? g("wid") ?? 0);
    if (!name || !l || !w) return null;
    return {
      name: String(name), l, w, qty: parseInt(g("qty") ?? g("quantity") ?? 1) || 1,
      glass: String(g("glass") ?? g("type") ?? ""), thk: parseFloat(g("thick") ?? 0) || "",
      rotate: !/^(n|no|false|0)$/i.test(String(g("rotat") ?? "yes")),
      priority: parseInt(g("prior") ?? 1) || 1, label: String(g("label") ?? ""),
    };
  };
  const importFile = e => {
    const f = e.target.files[0]; if (!f) return;
    const rd = new FileReader();
    rd.onload = ev => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "array" });
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
        const mapped = rows.map(mapRow).filter(Boolean);
        if (mapped.length) setParts(ps => [...ps, ...mapped]);
      } catch (err) { alert("Could not read file: " + err.message); }
    };
    rd.readAsArrayBuffer(f);
    e.target.value = "";
  };
  const importPaste = () => {
    const lines = paste.trim().split(/\n+/).map(l => l.split(/[,\t]/).map(c => c.trim())).filter(l => l.length >= 3);
    const mapped = lines.map(c => {
      const l = parseFloat(c[1]), w = parseFloat(c[2]);
      if (!c[0] || !l || !w || isNaN(l)) return null;
      return { name: c[0], l, w, qty: parseInt(c[3]) || 1, glass: c[4] || "", thk: parseFloat(c[5]) || "", rotate: !/^(n|no)$/i.test(c[6] || "y"), priority: parseInt(c[7]) || 1, label: c[8] || "" };
    }).filter(Boolean);
    if (mapped.length) { setParts(ps => [...ps, ...mapped]); setPaste(""); }
  };

  const upd = (setter) => (i, k, v) => setter(a => a.map((x, j) => j === i ? { ...x, [k]: v } : x));
  const updStock = upd(setStocks), updPart = upd(setParts);
  const numf = (i, k, fn) => e => fn(i, k, e.target.value === "" ? "" : +e.target.value);

  return (
    <div className="gco">
      <style>{CSS}</style>
      <div className="topbar noprint">
        <div>
          <h1>Om Glass — Cutting Optimizer</h1>
          <div className="tag">MaxRects nesting · kerf & grain aware · remnant reuse</div>
        </div>
        <div className="tabs">
          {[["stock", "Stock"], ["parts", "Cutting List"], ["layout", "Layouts"], ["remnants", "Leftovers"], ["reports", "Reports"]].map(([k, l]) => (
            <button key={k} className={"tabbtn" + (tab === k ? " on" : "")} onClick={() => setTab(k)}>{l}
              {k === "remnants" && remnants.length > 0 && <span className="chip" style={{ marginLeft: 5 }}>{remnants.length}</span>}
            </button>
          ))}
        </div>
        <button className="btn pri big" onClick={run}>Optimize ▸</button>
      </div>
      <div className="wrap">

        {/* ============ STOCK TAB ============ */}
        {tab === "stock" && (
          <div className="panel noprint">
            <h2><span className="n">1</span> Stock Sheet Inventory</h2>
            <div style={{ overflowX: "auto" }}>
              <table className="grid">
                <thead><tr>
                  <th>ID</th><th>Glass Type</th><th>Thk (mm)</th><th>Colour</th><th>Length</th><th>Width</th>
                  <th>Qty</th><th>Cost/Sheet ₹</th><th>Supplier</th><th>Location</th><th>Remarks</th><th></th>
                </tr></thead>
                <tbody>
                  {stocks.map((s, i) => (
                    <tr key={i}>
                      <td><input className="f sm mono" value={s.id} onChange={e => updStock(i, "id", e.target.value)} /></td>
                      <td><input className="f" style={{ minWidth: 110 }} value={s.glass} onChange={e => updStock(i, "glass", e.target.value)} /></td>
                      <td><input className="f sm num" type="number" value={s.thk} onChange={numf(i, "thk", updStock)} /></td>
                      <td><input className="f sm" value={s.colour} onChange={e => updStock(i, "colour", e.target.value)} /></td>
                      <td><input className="f sm num" type="number" value={s.l} onChange={numf(i, "l", updStock)} /></td>
                      <td><input className="f sm num" type="number" value={s.w} onChange={numf(i, "w", updStock)} /></td>
                      <td><input className="f sm num" type="number" value={s.qty} onChange={numf(i, "qty", updStock)} /></td>
                      <td><input className="f sm num" type="number" value={s.cost} onChange={numf(i, "cost", updStock)} /></td>
                      <td><input className="f" style={{ minWidth: 90 }} value={s.supplier} onChange={e => updStock(i, "supplier", e.target.value)} /></td>
                      <td><input className="f sm" value={s.loc} onChange={e => updStock(i, "loc", e.target.value)} /></td>
                      <td><input className="f" style={{ minWidth: 80 }} value={s.rem} onChange={e => updStock(i, "rem", e.target.value)} /></td>
                      <td><button className="btn danger" onClick={() => setStocks(a => a.filter((_, j) => j !== i))}>✕</button></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <div style={{ marginTop: 10 }}>
              <button className="btn" onClick={() => setStocks(a => [...a, { id: "S" + String(a.length + 1).padStart(2, "0"), glass: "Clear Float", thk: 5, colour: "Clear", l: 2440, w: 1830, qty: 1, cost: 0, supplier: "", loc: "", rem: "" }])}>+ Add Stock Sheet</button>
            </div>
          </div>
        )}

        {/* ============ PARTS TAB ============ */}
        {tab === "parts" && (<>
          <div className="panel noprint">
            <h2><span className="n">2</span> Cutting Order</h2>
            <div style={{ overflowX: "auto" }}>
              <table className="grid">
                <thead><tr>
                  <th>Part</th><th>Length</th><th>Width</th><th>Qty</th><th>Glass Type</th><th>Thk</th>
                  <th>Rotate</th><th>Priority</th><th>Label</th><th className="num">Area m²</th><th></th>
                </tr></thead>
                <tbody>
                  {parts.map((p, i) => (
                    <tr key={i}>
                      <td style={{ whiteSpace: "nowrap" }}>
                        <span style={{ display: "inline-block", width: 12, height: 12, background: colors[p.name], border: "1px solid #333", borderRadius: 2, marginRight: 6, verticalAlign: "middle" }} />
                        <input className="f sm" value={p.name} onChange={e => updPart(i, "name", e.target.value)} style={{ width: 70 }} />
                      </td>
                      <td><input className="f sm num" type="number" value={p.l} onChange={numf(i, "l", updPart)} /></td>
                      <td><input className="f sm num" type="number" value={p.w} onChange={numf(i, "w", updPart)} /></td>
                      <td><input className="f sm num" type="number" value={p.qty} onChange={numf(i, "qty", updPart)} /></td>
                      <td><input className="f" style={{ minWidth: 100 }} value={p.glass} onChange={e => updPart(i, "glass", e.target.value)} /></td>
                      <td><input className="f sm num" type="number" value={p.thk} onChange={numf(i, "thk", updPart)} /></td>
                      <td><select className="f" value={p.rotate ? "Y" : "N"} onChange={e => updPart(i, "rotate", e.target.value === "Y")}><option value="Y">Yes</option><option value="N">No</option></select></td>
                      <td><input className="f sm num" type="number" value={p.priority} onChange={numf(i, "priority", updPart)} /></td>
                      <td><input className="f" style={{ minWidth: 90 }} value={p.label} onChange={e => updPart(i, "label", e.target.value)} /></td>
                      <td className="num">{m2(p.l * p.w * p.qty)}</td>
                      <td><button className="btn danger" onClick={() => setParts(a => a.filter((_, j) => j !== i))}>✕</button></td>
                    </tr>
                  ))}
                </tbody>
                <tfoot><tr>
                  <td colSpan={3} style={{ fontWeight: 700 }}>Total</td>
                  <td className="num" style={{ fontWeight: 700 }}>{parts.reduce((a, p) => a + (+p.qty || 0), 0)}</td>
                  <td colSpan={5}></td>
                  <td className="num" style={{ fontWeight: 700 }}>{m2(parts.reduce((a, p) => a + p.l * p.w * p.qty, 0))}</td><td></td>
                </tr></tfoot>
              </table>
            </div>
            <div style={{ marginTop: 10, display: "flex", gap: 8, flexWrap: "wrap" }}>
              <button className="btn" onClick={() => setParts(a => [...a, { name: "P" + (a.length + 1), l: 600, w: 400, qty: 1, glass: "", thk: "", rotate: true, priority: 1, label: "" }])}>+ Add Part</button>
              <button className="btn" onClick={() => fileRef.current.click()}>⬆ Import Excel / CSV</button>
              <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={importFile} />
              <button className="btn ghost" onClick={() => setParts([])}>Clear List</button>
            </div>
          </div>
          <div className="panel noprint">
            <h2>Paste From Excel</h2>
            <p style={{ color: "var(--sub)", margin: "0 0 8px" }}>Columns: Part, Length, Width, Qty, Glass, Thickness, Rotate(Y/N), Priority, Label — comma or tab separated.</p>
            <textarea className="f" value={paste} onChange={e => setPaste(e.target.value)} placeholder={"A, 850, 450, 12\nB, 1200, 650, 8"} />
            <div style={{ marginTop: 8 }}><button className="btn" onClick={importPaste}>Add Pasted Rows</button></div>
          </div>
        </>)}

        {/* ============ SETTINGS (always above layout) ============ */}
        {(tab === "parts" || tab === "layout") && (
          <div className="panel noprint">
            <h2><span className="n">3</span> Optimization Controls</h2>
            <div className="settings">
              <div className="setting"><label>Kerf / Blade (mm)</label><input className="f" type="number" value={opt.kerf} onChange={e => setOpt(o => ({ ...o, kerf: +e.target.value || 0 }))} /></div>
              <div className="setting"><label>Edge Margin (mm)</label><input className="f" type="number" value={opt.margin} onChange={e => setOpt(o => ({ ...o, margin: +e.target.value || 0 }))} /></div>
              <div className="setting"><label>Rotation</label><select className="f" value={opt.rotation ? "Y" : "N"} onChange={e => setOpt(o => ({ ...o, rotation: e.target.value === "Y" }))}><option value="Y">Allowed (per part)</option><option value="N">Off</option></select></div>
              <div className="setting"><label>Grain Direction</label><select className="f" value={opt.grain ? "Y" : "N"} onChange={e => setOpt(o => ({ ...o, grain: e.target.value === "Y" }))}><option value="N">No grain</option><option value="Y">Locked (no rotate)</option></select></div>
              <div className="setting"><label>Algorithm</label><select className="f" value={opt.heuristic} onChange={e => setOpt(o => ({ ...o, heuristic: e.target.value }))}>
                <option value="G-AUTO">Guillotine — Score &amp; Snap (recommended)</option>
                <option value="G-V">Guillotine — Vertical strips</option>
                <option value="G-H">Guillotine — Horizontal strips</option>
                <option value="BSSF">MaxRects — Best Short Side (CNC)</option>
                <option value="BLSF">MaxRects — Best Long Side (CNC)</option>
                <option value="BAF">MaxRects — Best Area Fit (CNC)</option>
                <option value="BL">MaxRects — Bottom-Left (CNC)</option>
                <option value="AUTO">Auto — best of everything</option>
              </select></div>
              <div className="setting"><label>Min Leftover Side (mm)</label><input className="f" type="number" value={opt.minLeftSide} onChange={e => setOpt(o => ({ ...o, minLeftSide: +e.target.value || 0 }))} /></div>
              <div className="setting"><label>Min Leftover Area (m²)</label><input className="f" type="number" step="0.01" value={opt.minLeftArea} onChange={e => setOpt(o => ({ ...o, minLeftArea: +e.target.value || 0 }))} /></div>
              <div className="setting"><label>Use Remnants First</label><select className="f" value={opt.useRemnants ? "Y" : "N"} onChange={e => setOpt(o => ({ ...o, useRemnants: e.target.value === "Y" }))}><option value="Y">Yes</option><option value="N">No</option></select></div>
            </div>
          </div>
        )}

        {/* ============ LAYOUT TAB ============ */}
        {tab === "layout" && !result && (
          <div className="panel"><div className="imp">No optimization run yet. Enter your cutting list, then press <b>Optimize ▸</b>.</div></div>
        )}
        {tab === "layout" && result && stats && (<>
          <div className="panel">
            <h2><span className="n">4</span> Statistics Dashboard <span className="chip noprint">{runMs} ms · {result.sheets[0] ? "MaxRects" : ""}</span></h2>
            <div className="statgrid">
              <div className="stat acc"><div className="v">{stats.sheets}</div><div className="l">Sheets Used</div></div>
              <div className="stat"><div className="v">{stats.nParts}</div><div className="l">Pieces Placed</div></div>
              <div className="stat"><div className="v">{m2(stats.areaSheet)}</div><div className="l">Stock Area m²</div></div>
              <div className="stat good"><div className="v">{stats.yieldPct.toFixed(1)}%</div><div className="l">Yield</div></div>
              <div className="stat bad"><div className="v">{stats.wastePct.toFixed(1)}%</div><div className="l">True Waste</div></div>
              <div className="stat bad"><div className="v">{m2(stats.waste)}</div><div className="l">Waste m²</div></div>
              <div className="stat good"><div className="v">{stats.nLeft}</div><div className="l">Leftovers Saved</div></div>
              <div className="stat good"><div className="v">{m2(stats.areaLeft)}</div><div className="l">Leftover m²</div></div>
              <div className="stat acc"><div className="v">{inr(stats.cost)}</div><div className="l">Material Cost</div></div>
              <div className="stat"><div className="v">{inr(stats.costPer)}</div><div className="l">Cost / Piece</div></div>
              <div className="stat good"><div className="v">{inr(stats.savings)}</div><div className="l">Value in Leftovers</div></div>
            </div>
            {result.unplaced.length > 0 && (
              <div className="warn">⚠ <b>{result.unplaced.length} piece(s) could not be placed</b> — stock exhausted or piece larger than any sheet:
                {" "}{[...new Set(result.unplaced.map(p => `${p.name} (${p.l}×${p.w})`))].join(", ")}</div>
            )}
            <div className="noprint" style={{ marginTop: 12, display: "flex", gap: 8, flexWrap: "wrap" }}>
              <button className="btn pri" onClick={saveLeftovers} disabled={remSaved || stats.nLeft === 0}>{remSaved ? "✓ Leftovers Saved" : "Save Leftovers to Bank"}</button>
              <button className="btn" onClick={applyStockConsumption}>Deduct Used Sheets from Stock</button>
              <button className="btn" onClick={() => exportPDFPlan(result, stats, colors, opt)}>⬇ Download PDF Cutting Plan</button>
            </div>
          </div>

          <div style={{ display: "grid", gridTemplateColumns: sel ? "1fr 280px" : "1fr", gap: 14 }}>
            <div>
              {result.sheets.map(sh => (
                <div className="sheetcard" key={sh.no}>
                  <div className="titleblock">
                    <div className="cell"><div className="k">Sheet</div><div className="x">#{sh.no}</div></div>
                    <div className="cell"><div className="k">{sh.src === "REMNANT" ? "Remnant" : "Stock"}</div><div className="x">{sh.id}</div></div>
                    <div className="cell"><div className="k">Size</div><div className="x">{sh.W} × {sh.H}</div></div>
                    <div className="cell"><div className="k">Glass</div><div className="x" style={{ fontFamily: "'Inter',sans-serif", fontWeight: 600 }}>{sh.glass} {sh.thk}mm</div></div>
                    <div className="cell"><div className="k">Pieces</div><div className="x">{sh.placed.length}</div></div>
                    <div className="cell"><div className="k">Yield</div><div className="x">{(sh.util * 100).toFixed(1)}%</div></div>
                    <div className="cell"><div className="k">Method</div><div className="x" style={{ fontFamily: "'Inter',sans-serif", fontWeight: 600, fontSize: 12.5 }}>{HEURISTICS[sh.heur] || sh.heur}</div></div>
                    <div className="cell noprint" style={{ marginLeft: "auto", borderRight: "none", display: "flex", alignItems: "center" }}>
                      <button className="btn ghost" onClick={() => exportPNG(sh, colors, opt)}>⬇ PNG</button>
                    </div>
                  </div>
                  <SheetSVG sh={sh} colors={colors} sel={sel} onSel={setSel} />
                  <div className="legend">
                    {[...new Set(sh.placed.map(p => p.part.name))].map(n => (
                      <span className="it" key={n}><span className="sw" style={{ background: colors[n] }} /> {n}</span>
                    ))}
                    <span className="it"><span className="sw" style={{ background: "#E1F2E6", borderColor: "#2E7D4F", borderStyle: "dashed" }} /> Reusable leftover</span>
                    <span className="it"><span className="sw" style={{ background: "repeating-linear-gradient(45deg,#F8E7E3,#F8E7E3 3px,#C0392B55 3px,#C0392B55 5px)" }} /> Waste</span>
                    <span className="it mono" style={{ marginLeft: "auto", color: "var(--sub)" }}>kerf {opt.kerf}mm · margin {opt.margin}mm · ① cut order{sh.strips ? " · ┅ C1,C2 full-length snap cuts" : ""}</span>
                  </div>
                </div>
              ))}
            </div>
            {sel && (
              <div className="panel detail noprint">
                <h2>Piece Detail</h2>
                <div className="kv">
                  <span className="k">Part</span><span className="x">{sel.p.part.name}</span>
                  <span className="k">Label</span><span className="x">{sel.p.part.label || "—"}</span>
                  <span className="k">Cut size</span><span className="x">{Math.round(sel.p.w)} × {Math.round(sel.p.h)} mm</span>
                  <span className="k">Ordered as</span><span className="x">{sel.p.part.l} × {sel.p.part.w}</span>
                  <span className="k">Rotated</span><span className="x">{sel.p.rot ? "Yes ⟳" : "No"}</span>
                  <span className="k">Area</span><span className="x">{m2(sel.p.w * sel.p.h)} m²</span>
                  <span className="k">Glass</span><span className="x">{sel.p.part.glass || "Any"} {sel.p.part.thk || ""}</span>
                  <span className="k">Sheet</span><span className="x">#{sel.sheet} ({sel.sh.id})</span>
                  <span className="k">Cut seq</span><span className="x">{sel.idx + 1}</span>
                  <span className="k">Position</span><span className="x">x{Math.round(sel.p.x + opt.margin)}, y{Math.round(sel.p.y + opt.margin)}</span>
                  <span className="k">Total qty</span><span className="x">{sel.p.part.qty}</span>
                </div>
                <button className="btn ghost" style={{ marginTop: 10, width: "100%" }} onClick={() => setSel(null)}>Close</button>
              </div>
            )}
          </div>
        </>)}

        {/* ============ REMNANTS TAB ============ */}
        {tab === "remnants" && (
          <div className="panel noprint">
            <h2><span className="n">5</span> Leftover Bank</h2>
            {remnants.length === 0 && <div className="imp">No remnants saved yet. Run an optimization, then press <b>Save Leftovers to Bank</b>. Future jobs will consume these before new stock.</div>}
            {remnants.length > 0 && (
              <table className="grid">
                <thead><tr><th>ID</th><th>Length</th><th>Width</th><th className="num">Area m²</th><th>Glass</th><th>Thk</th><th>Date</th><th>From</th><th></th></tr></thead>
                <tbody>{remnants.map((r, i) => (
                  <tr key={r.id}>
                    <td className="mono">{r.id}</td><td className="num">{r.l}</td><td className="num">{r.w}</td>
                    <td className="num">{m2(r.l * r.w)}</td><td>{r.glass}</td><td className="num">{r.thk}</td>
                    <td className="mono">{r.date}</td><td>{r.ref}</td>
                    <td><button className="btn danger" onClick={() => setRemnants(a => a.filter((_, j) => j !== i))}>✕</button></td>
                  </tr>
                ))}</tbody>
              </table>
            )}
            <div style={{ marginTop: 10 }}>
              <button className="btn" onClick={() => setRemnants(a => [...a, { id: "RM" + (a.length + 1), l: 1000, w: 600, glass: "Clear Float", thk: 5, date: today(), ref: "Manual entry" }])}>+ Add Remnant Manually</button>
            </div>
          </div>
        )}

        {/* ============ REPORTS TAB ============ */}
        {tab === "reports" && (
          <div className="panel noprint">
            <h2><span className="n">6</span> Reports & Exports</h2>
            {!result && <div className="imp">Run an optimization first to generate reports.</div>}
            {result && stats && (
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                <button className="btn pri" onClick={() => exportPDFPlan(result, stats, colors, opt)}>⬇ Cutting Plan PDF</button>
                <button className="btn" onClick={() => downloadCSV([
                  ["Stock ID", "Glass", "Thk", "Size", "Opening Qty", "Used", "Remaining"],
                  ...result.stockLeft.map(s => [s.id, s.glass, s.thk, `${s.l}x${s.w}`, s.qty, s.qty - s.left, s.left]),
                ], "stock-consumption")}>⬇ Stock Consumption CSV</button>
                <button className="btn" onClick={() => downloadCSV([
                  ["Sheet", "Source", "ID", "Size", "Seq", "Part", "Label", "Cut L", "Cut W", "Rotated", "X", "Y"],
                  ...result.sheets.flatMap(sh => sh.placed.map((p, i) => [sh.no, sh.src, sh.id, `${sh.W}x${sh.H}`, i + 1, p.part.name, p.part.label, Math.round(p.w), Math.round(p.h), p.rot ? "Y" : "N", Math.round(p.x + opt.margin), Math.round(p.y + opt.margin)])),
                ], "production-summary")}>⬇ Production Summary CSV</button>
                <button className="btn" onClick={() => downloadCSV([
                  ["Sheet", "Leftover L", "Leftover W", "Area m2", "Glass", "Thk"],
                  ...result.sheets.flatMap(sh => sh.leftovers.map(r => [sh.no, Math.round(r.w), Math.round(r.h), m2(r.w * r.h), sh.glass, sh.thk])),
                ], "leftover-report")}>⬇ Leftover Report CSV</button>
                <button className="btn" onClick={() => downloadCSV([
                  ["Metric", "Value"],
                  ["Sheets used", stats.sheets], ["Pieces placed", stats.nParts],
                  ["Stock area m2", m2(stats.areaSheet)], ["Yield %", stats.yieldPct.toFixed(2)],
                  ["Waste %", stats.wastePct.toFixed(2)], ["Leftover m2", m2(stats.areaLeft)],
                  ["Total cost", Math.round(stats.cost)], ["Cost per piece", Math.round(stats.costPer)],
                ], "cost-analysis")}>⬇ Cost Analysis CSV</button>
              </div>
            )}
            {result && (
              <p style={{ color: "var(--sub)", marginTop: 12 }}>
                The PDF cutting plan contains one layout page per sheet plus a production summary page, ready for the shop floor.
                PNG export is available on each sheet card in the Layouts tab.
              </p>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
