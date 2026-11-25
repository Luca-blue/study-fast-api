// Shared drawing helpers. All coordinates are in "px" of a 1920x1080 canvas
// (same as the HTML deck) and converted to inches (144 px = 1 in).
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const md = require("react-icons/md");

const IN = (px) => px / 144;
const FONT = "Meiryo";
const MONO = "Consolas";
const COL = {
  black: "0A0A0A", white: "FFFFFF", pink: "F2D4CF", green: "E5EDD6", gray: "F5F5F5",
  dark: "333333", muted: "9A9A9A", codeStr: "C7DDB0", codeType: "A8D8EA",
};
const LINE = 2.25; // 3px

// ---------- layout ----------
const FRAME = { x: 40, y: 34, w: 1834, h: 1006 };
const HEAD_H = 132;
const BODY = { x: 40, y: 34 + HEAD_H, w: 1834, h: 1006 - HEAD_H };

function split(total, specs) {
  const fixed = specs.filter((s) => typeof s === "number").reduce((a, b) => a + b, 0);
  const fr = specs.filter((s) => typeof s === "string").reduce((a, s) => a + parseFloat(s), 0);
  const rest = total - fixed;
  let pos = 0;
  return specs.map((s) => {
    const size = typeof s === "number" ? s : (rest * parseFloat(s)) / fr;
    const r = [pos, size];
    pos += size;
    return r;
  });
}
// grid(cols, rows, area) -> cell(r, c, r2, c2) returns rect spanning rows r..r2, cols c..c2
function grid(cols, rows, area = BODY) {
  const cx = split(area.w, cols), ry = split(area.h, rows);
  return (r, c, r2 = r, c2 = c) => ({
    x: area.x + cx[c][0], y: area.y + ry[r][0],
    w: cx[c2][0] + cx[c2][1] - cx[c][0], h: ry[r2][0] + ry[r2][1] - ry[r][0],
  });
}
const inset = (r, p = 40, py = p) => ({ x: r.x + p, y: r.y + py, w: r.w - 2 * p, h: r.h - 2 * py });

// ---------- text width estimate (for pills / auto boxes) ----------
function textW(s, px, mono = false) {
  let w = 0;
  for (const ch of s) w += /[　-鿿＀-￯]/.test(ch) ? px : px * (mono ? 0.56 : 0.58);
  return w;
}

// ---------- primitives ----------
function rect(slide, r, fill = COL.white, opts = {}) {
  slide.addShape("rect", {
    x: IN(r.x), y: IN(r.y), w: IN(r.w), h: IN(r.h),
    fill: { color: fill }, line: { color: COL.black, width: opts.lineW ?? LINE },
    shadow: opts.shadow ? { type: "outer", color: COL.black, opacity: 1, blur: 0, offset: opts.shadow === "lg" ? 4.5 : 3, angle: 45 } : undefined,
    objectName: opts.name,
  });
}
// runs: string or [{text, options}] ; px font size
function text(slide, runs, r, o = {}) {
  const px = o.px ?? 28;
  slide.addText(runs, {
    x: IN(r.x), y: IN(r.y), w: IN(r.w), h: IN(r.h),
    fontFace: o.mono ? MONO : FONT, fontSize: px / 2, color: o.color ?? COL.black, bold: !!o.bold,
    margin: 0, valign: o.valign ?? "top", align: o.align ?? "left", isTextBox: true,
    lineSpacingMultiple: o.lh ?? 1.25, paraSpaceAfter: o.psa ?? 0, fit: "none",
    fill: o.fill ? { color: o.fill } : undefined, objectName: o.name,
  });
}
// bold + normal mixed: "**bold** normal" -> runs
function md2runs(s, base = {}) {
  const out = [];
  s.split("\n").forEach((line, li, arr) => {
    const parts = line.split(/(\*\*[^*]+\*\*|`[^`]+`)/).filter(Boolean);
    parts.forEach((p, pi) => {
      const o = { ...base };
      let t = p;
      if (p.startsWith("**")) { t = p.slice(2, -2); o.bold = true; }
      else if (p.startsWith("`")) { t = p.slice(1, -1); o.fontFace = MONO; o.bold = true; o.highlight = undefined; }
      if (pi === parts.length - 1 && li < arr.length - 1) o.breakLine = true;
      out.push({ text: t, options: o });
    });
    if (!parts.length && li < arr.length - 1) out.push({ text: "", options: { breakLine: true } });
  });
  return out;
}
function label(slide, s, x, y, o = {}) {
  const px = o.px ?? 21;
  const w = textW(s, px) + 34;
  slide.addText(s, {
    x: IN(x), y: IN(y), w: IN(w), h: IN(px + 16),
    fontFace: FONT, fontSize: px / 2, bold: true, color: o.color ?? COL.white,
    fill: { color: o.fill ?? COL.black }, margin: 0, align: "center", valign: "middle", isTextBox: true,
  });
  return { w, h: px + 16 };
}
function badge(slide, s, x, y, size = 40, fill = COL.pink, border = COL.black) {
  slide.addText(s, {
    x: IN(x), y: IN(y), w: IN(size), h: IN(size), fontFace: MONO, bold: true, fontSize: size * 0.27,
    color: COL.black, fill: { color: fill }, line: { color: border, width: 1.5 },
    margin: 0, align: "center", valign: "middle", isTextBox: true,
  });
}

// ---------- code ----------
const KW = /\b(from|import|def|return|class|with|as|for|in|not|del|True|False|None|if|const|await|async|function|let)\b/;
const TOKEN = /(#.*$|\/\/.*$)|(f?"(?:[^"\\]|\\.)*"|`[^`]*`)|(@[\w.]+)|\b(from|import|def|return|class|with|as|for|in|not|del|True|False|None|if|const|await|async|function|let)\b|\b(str|int|bool|list|max)\b/g;
function highlight(code) {
  const all = [];
  const lines = code.split("\n");
  lines.forEach((line, i) => {
    const runs = [];
    const push = (t, o = {}) => t && runs.push({ text: t, options: { color: COL.white, ...o } });
    if (line.startsWith("!!")) { push(line.slice(2), { color: COL.codeStr, bold: true }); }
    else if (line.startsWith("$ ") || line.startsWith("> ")) { push(line.slice(0, 2), { color: COL.muted }); push(line.slice(2)); }
    else {
      let last = 0, m;
      TOKEN.lastIndex = 0;
      while ((m = TOKEN.exec(line))) {
        push(line.slice(last, m.index));
        if (m[1]) push(m[0], { color: COL.muted });
        else if (m[2]) push(m[0], { color: COL.codeStr });
        else if (m[3]) push(m[0], { color: COL.pink, bold: true });
        else if (m[4]) push(m[0], { color: COL.pink, bold: true });
        else push(m[0], { color: COL.codeType });
        last = m.index + m[0].length;
      }
      push(line.slice(last));
    }
    if (!runs.length) runs.push({ text: " ", options: { color: COL.white } });
    if (i < lines.length - 1) runs[runs.length - 1].options.breakLine = true;
    all.push(...runs);
  });
  return all;
}
// code text at r (no background). px font. badges: {lineIndex: "1"} drawn at right edge
function code(slide, src, r, o = {}) {
  const px = o.px ?? 24, lh = o.lh ?? 1.5;
  slide.addText(highlight(src), {
    x: IN(r.x), y: IN(r.y), w: IN(r.w), h: IN(r.h), fontFace: MONO, fontSize: px / 2,
    margin: 0, valign: "top", lineSpacing: (px / 2) * lh, isTextBox: true, fit: "none",
    objectName: o.name,
  });
  if (o.badges) {
    for (const [li, s] of Object.entries(o.badges)) {
      const bs = Math.min(34, px * lh - 4);
      badge(slide, s, r.x + r.w - bs, r.y + li * px * lh + (px * lh - bs) / 2 - 1, bs, COL.pink, COL.pink);
    }
  }
}
// terminal / code box with optional bar title
function term(slide, src, r, o = {}) {
  rect(slide, r, COL.black, { shadow: o.shadow === false ? false : true });
  let y = r.y;
  if (o.bar) {
    text(slide, o.bar, { x: r.x + 18, y: r.y + 8, w: r.w - 36, h: 26 }, { px: 18, color: COL.muted, mono: true });
    slide.addShape("line", { x: IN(r.x), y: IN(r.y + 40), w: IN(r.w), h: 0, line: { color: "444444", width: 1.5 } });
    y = r.y + 40;
  }
  code(slide, src, { x: r.x + 22, y: y + 16, w: r.w - 44, h: r.y + r.h - y - 26 }, o);
}

// ---------- composite ----------
function box(slide, r, fill = COL.white, shadow = true) { rect(slide, r, fill, { shadow }); }
// numbered notes: items [{n, h, p}] stacked; returns bottom y
function notes(slide, items, x, y, w, o = {}) {
  const hpx = o.hpx ?? 28, ppx = o.ppx ?? 24, gap = o.gap ?? 24;
  for (const it of items) {
    badge(slide, it.n, x, y + 2, 40);
    const tw = w - 60;
    text(slide, it.h, { x: x + 60, y, w: tw, h: hpx * 1.4 }, { px: hpx, bold: true });
    const lines = it.p ? Math.ceil(textW(it.p.replace(/\*\*|`/g, ""), ppx) / tw + 0.05) + (it.p.split("\n").length - 1) : 0;
    const ph = lines * ppx * 1.55;
    if (it.p) text(slide, md2runs(it.p), { x: x + 60, y: y + hpx * 1.45, w: tw, h: ph + 6 }, { px: ppx, lh: 1.3 });
    y += hpx * 1.45 + ph + gap;
  }
  return y;
}
// task card: header + lettered steps
function task(slide, title, steps, x, y, w, o = {}) {
  const px = o.px ?? 24, lh = 1.5;
  const tw = w - 44 - 52;
  const hs = steps.map((s) => {
    const lines = s.split("\n").reduce((a, l) => a + Math.max(1, Math.ceil(textW(l.replace(/\*\*|`/g, ""), px) / tw + 0.05)), 0);
    return Math.max(40, lines * px * lh);
  });
  const h = 46 + 20 + hs.reduce((a, b) => a + b, 0) + (steps.length - 1) * 12 + 18;
  rect(slide, { x, y, w, h }, COL.white, { shadow: true });
  text(slide, title, { x, y, w, h: 46 }, { px: 22, bold: true, color: COL.white, fill: COL.black, valign: "middle" });
  // left padding inside header
  let yy = y + 46 + 20;
  steps.forEach((s, i) => {
    badge(slide, "ABCDEFG"[i], x + 22, yy + 2, 36, COL.green);
    text(slide, md2runs(s), { x: x + 22 + 52, y: yy, w: tw, h: hs[i] }, { px, lh: 1.3 });
    yy += hs[i] + 12;
  });
  return { h };
}

// ---------- icons ----------
async function iconPng(name, color = "#0A0A0A") {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(md[name], { color, size: 256 }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}
const ICONS = {};
async function loadIcons(names) { for (const n of names) ICONS[n] = await iconPng(n); }
function ibox(slide, name, x, y, size = 96, fill = COL.white) {
  rect(slide, { x, y, w: size, h: size }, fill, { shadow: true });
  const s = size * 0.66;
  slide.addImage({ data: ICONS[name], x: IN(x + (size - s) / 2), y: IN(y + (size - s) / 2), w: IN(s), h: IN(s), altText: name });
}

module.exports = { IN, FONT, MONO, COL, LINE, FRAME, HEAD_H, BODY, grid, inset, textW, rect, text, md2runs, label, badge, code, term, box, notes, task, loadIcons, ibox };
