const M3 = [
  "JAN",
  "FEB",
  "MAR",
  "APR",
  "MAY",
  "JUN",
  "JUL",
  "AUG",
  "SEP",
  "OCT",
  "NOV",
  "DEC",
];
const MN = [
  "january",
  "february",
  "march",
  "april",
  "may",
  "june",
  "july",
  "august",
  "september",
  "october",
  "november",
  "december",
];
const DOW = ["SU", "M", "T", "W", "TH", "F", "S"];
const STOP = ["JR", "SR", "II", "III", "IV", "NONE", "NA"];
const att = [],
  sf = [];
let month = null,
  lateMode = "p";
const $ = (s) => document.querySelector(s);
const esc = (s) =>
  String(s).replace(
    /[&<>"]/g,
    (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c],
  );
const col = (i) => XLSX.utils.encode_col(i);

// ---------- loading ----------
$("#drop").onclick = () => $("#fi").click();
$("#fi").onchange = (e) => load([...e.target.files]);
["dragover", "dragenter"].forEach((ev) =>
  $("#drop").addEventListener(ev, (e) => {
    e.preventDefault();
    $("#drop").classList.add("on");
  }),
);
["dragleave", "drop"].forEach((ev) =>
  $("#drop").addEventListener(ev, (e) => {
    e.preventDefault();
    $("#drop").classList.remove("on");
  }),
);
$("#drop").addEventListener("drop", (e) => load([...e.dataTransfer.files]));
async function load(files) {
  for (const f of files) {
    try {
      const buf = await f.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      if (
        wb.SheetNames.includes("Masterlist") &&
        wb.SheetNames.includes("JUN")
      ) {
        const i = sf.findIndex((x) => x.name === f.name);
        if (i >= 0) sf.splice(i, 1);
        sf.push({
          name: f.name,
          buf,
          wb,
          sy: parseInt(String((wb.Sheets.Info?.B6 || {}).v || "2026")) || 2026,
        });
      } else {
        const i = att.findIndex((x) => x.name === f.name);
        if (i >= 0) att.splice(i, 1);
        att.push({
          name: f.name,
          wb,
          sheets: wb.SheetNames.map((n) => parseAtt(wb.Sheets[n], n)).filter(
            (s) => s.students.length,
          ),
        });
      }
    } catch (e) {
      console.error(e);
    }
  }
  detectMonth();
  render();
}

// ---------- parsing ----------
function parseAtt(ws, name) {
  const R = XLSX.utils.sheet_to_json(ws, {
    header: 1,
    defval: null,
    raw: true,
  });
  const dates = [];
  const hdr = R[1] || [],
    dw = R[2] || [];
  for (let c = 2; c < hdr.length; c++) {
    if (typeof hdr[c] !== "number") break;
    dates.push({ ci: c, day: hdr[c], dow: String(dw[c] || "").toUpperCase() });
  }
  const students = [];
  for (let r = 3; r < R.length; r++) {
    const nm = R[r][1];
    if (typeof nm !== "string" || nm.startsWith("<") || /^combined/i.test(nm))
      continue;
    students.push({ name: nm, row: R[r] });
  }
  return { name, dates, students };
}
function norm(s) {
  return String(s)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/[^A-Z, ]/g, " ");
}
function pn(s) {
  const [l, g = ""] = norm(s).split(",");
  return {
    l: l.replace(/\s+/g, ""),
    g: g.split(/\s+/).filter((t) => t && !STOP.includes(t)),
  };
}
function lev(a, b) {
  const d = Array.from({ length: a.length + 1 }, (_, i) => [i]);
  for (let j = 1; j <= b.length; j++) d[0][j] = j;
  for (let i = 1; i <= a.length; i++)
    for (let j = 1; j <= b.length; j++)
      d[i][j] = Math.min(
        d[i - 1][j] + 1,
        d[i][j - 1] + 1,
        d[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1),
      );
  return d[a.length][b.length];
}
function score(m, a) {
  if (m.l !== a.l)
    return lev(m.l, a.l) <= 1 && m.g.length && m.g.every((t) => a.g.includes(t))
      ? 1.5
      : 0;
  if (m.g.every((t) => a.g.includes(t))) return 2 + m.g.length;
  return lev(m.g.join(" "), a.g.slice(0, m.g.length).join(" ")) <= 2 ? 1 : 0;
}
function masterNames(sfile) {
  const ws = sfile.wb.Sheets.Masterlist,
    out = [];
  const add = (from, to, base) => {
    for (let r = from; r <= to; r++) {
      const c = ws["B" + r];
      if (c && typeof c.v === "string" && c.v.trim())
        out.push({ name: c.v.trim(), row: base + (r - from) });
    }
  };
  add(4, 53, 13);
  add(55, 104, 64);
  return out;
}
function autoMap(masters, sheet) {
  const used = new Set(),
    map = [];
  const pa = sheet.students.map((s) => pn(s.name));
  masters.forEach((m) => {
    const pm = pn(m.name);
    let best = -1,
      bs = 0;
    pa.forEach((a, i) => {
      if (used.has(i)) return;
      const s = score(pm, a);
      if (s > bs) {
        bs = s;
        best = i;
      }
    });
    if (best >= 0) used.add(best);
    map.push(best);
  });
  return map;
}
function detectMonth() {
  month = null;
  for (const a of att) {
    const m = MN.findIndex((n) => a.name.toLowerCase().includes(n));
    if (m >= 0) {
      month = M3[m];
      break;
    }
  }
  if (!month && att.length && sf.length) {
    const sy = sf[0].sy,
      cand = [];
    for (let m = 0; m < 12; m++) {
      if (m >= 4 && m < 5) continue;
      const y = m >= 5 ? sy : sy + 1;
      if (
        att.every((a) =>
          a.sheets.every((s) =>
            s.dates.every((d) => DOW[new Date(y, m, d.day).getDay()] === d.dow),
          ),
        )
      )
        cand.push(M3[m]);
    }
    month = cand.find((c) => sf[0].wb.SheetNames.includes(c)) || null;
  }
}

// ---------- pairing & planning ----------
function buildPairs() {
  const pairs = [];
  const sheets = att.flatMap((a) =>
    a.sheets.map((s) => ({ ...s, file: a.name })),
  );
  sf.forEach((f) => {
    const masters = masterNames(f);
    let best = null;
    sheets.forEach((s) => {
      const map = autoMap(masters, s);
      const n = map.filter((x) => x >= 0).length;
      if (!best || n > best.n) best = { s, map, n };
    });
    if (best && best.n >= Math.max(1, masters.length / 2))
      pairs.push({ f, masters, sheet: best.s, map: best.map });
  });
  return pairs;
}
let pairs = [];
function plan(p) {
  const ws = p.f.wb.Sheets[month];
  const res = {
    ok: !!ws,
    changes: [],
    cols: [],
    skipped: [],
    future: [],
    bad: [],
    prot: 0,
    blank: 0,
    absent: 0,
    late: 0,
    total: 0,
    grid: [],
  };
  if (!ws) return res;
  const m = M3.indexOf(month),
    y = m >= 5 ? p.f.sy : p.f.sy + 1,
    today = new Date();
  const target = {};
  for (let c = 6; c <= 30; c++) {
    const d = ws[col(c) + "10"],
      w = ws[col(c) + "11"];
    if (d && typeof d.v === "number")
      target[d.v] = { c, dow: String(w?.v || "").toUpperCase() };
  }
  p.sheet.dates.forEach((d) => {
    const t = target[d.day],
      actual = DOW[new Date(y, m, d.day).getDay()];
    if (actual !== d.dow || (t && t.dow !== d.dow)) {
      res.bad.push(d.day);
      return;
    }
    if (!t) {
      res.skipped.push(`${d.day} (${d.dow})`);
      return;
    }
    if (new Date(y, m, d.day) > today) {
      res.future.push(d.day);
      return;
    }
    res.cols.push({ ...d, c: t.c });
  });
  p.masters.forEach((mm, mi) => {
    const si = p.map[mi],
      stu = si >= 0 ? p.sheet.students[si] : null,
      cells = [];
    res.cols.forEach((d) => {
      if (!stu) {
        cells.push(null);
        return;
      }
      const raw = String(stu.row[d.ci] ?? "")
        .trim()
        .toUpperCase();
      if (!["X", "L", "P"].includes(raw)) {
        res.blank++;
        cells.push(null);
        return;
      }
      const mark =
        raw === "X" ? "x" : raw === "L" ? (lateMode === "h" ? "h" : "") : "";
      const addr = col(d.c) + mm.row,
        cur = String(ws[addr]?.v ?? "").trim();
      res.total++;
      if (raw === "X") res.absent++;
      if (raw === "L") res.late++;
      if (cur && !/^[xh]$/i.test(cur)) {
        res.prot++;
        cells.push({ t: "🔒", k: "p" });
        return;
      }
      if (cur.toLowerCase() !== mark) res.changes.push({ addr, val: mark });
      cells.push({ t: mark || "·", k: mark || "p" });
    });
    res.grid.push({ name: mm.name, cells, matched: !!stu });
  });
  return res;
}

// ---------- rendering ----------
function render() {
  $("#chips").innerHTML = [
    ...att.map((a) => `<span class="chip">📄 ${esc(a.name)}</span>`),
    ...sf.map((f) => `<span class="chip">📘 ${esc(f.name)}</span>`),
  ].join("");
  const out = $("#out");
  if (!att.length || !sf.length) {
    out.innerHTML = `<div class="note">${!att.length ? "Add the attendance export. " : ""}${!sf.length ? "Add at least one SF2 file." : ""}</div>`;
    return;
  }
  pairs = buildPairs();
  const months = sf[0].wb.SheetNames.filter((n) => M3.includes(n));
  let h = `<div class="card row"><label>Month to fill: <select id="mo">${months.map((n) => `<option ${n === month ? "selected" : ""}>${n}</option>`).join("")}</select></label>
  <label>Late (L) counts as: <select id="lm"><option value="p" ${lateMode === "p" ? "selected" : ""}>Present (blank)</option><option value="h" ${lateMode === "h" ? "selected" : ""}>Half-day (h)</option></select></label></div>`;
  if (!month)
    h += `<div class="note w">Couldn’t detect the month from the file name — please pick it above.</div>`;
  if (!pairs.length)
    h += `<div class="note e">Couldn’t match the attendance names to any SF2 masterlist.</div>`;
  pairs.forEach((p, pi) => {
    const r = plan(p),
      un = p.masters.filter((_, i) => p.map[i] < 0);
    const rate = r.total ? r.absent / r.total : 0;
    h += `<div class="card"><h2>${esc(p.sheet.name)} → ${esc(p.f.name)} · ${esc(month || "?")} tab</h2>`;
    if (!r.ok) {
      h += `<div class="note e">This SF2 file has no “${esc(month)}” tab.</div></div>`;
      return;
    }
    h += `<div class="stats"><span><b>${p.map.filter((x) => x >= 0).length}/${p.masters.length}</b> students matched</span><span><b>${r.cols.length}</b> class days filled</span><span><b>${r.absent}</b> absences (x)</span><span><b>${r.late}</b> late</span><span><b>${r.changes.length}</b> cells to change</span></div>`;
    if (r.skipped.length)
      h += `<div class="note w">Skipped (no column on the SF2 for these dates; SF2 is Mon–Fri only): ${r.skipped.join(", ")}</div>`;
    if (r.bad.length)
      h += `<div class="note e">Dates ${r.bad.join(", ")} don’t fall on the stated weekday in ${esc(month)} — is the month right?</div>`;
    if (r.future.length)
      h += `<div class="note">Not filled yet (dates still ahead): ${r.future.join(", ")}</div>`;
    if (r.blank)
      h += `<div class="note">${r.blank} empty cell(s) in the export (no P, L or X) were left untouched on the SF2.</div>`;
    if (r.prot)
      h += `<div class="note w">${r.prot} cell(s) hold other text (e.g. HOLIDAY/SUSPENDED) and are left as they are (🔒).</div>`;
    if (rate > 0.85 && r.total > 20)
      h += `<div class="note w">${Math.round(rate * 100)}% of the marks in this export are absences — check the export before using it.</div>`;
    if (un.length)
      h += `<div class="note w">Not matched: ${un
        .map((m) => {
          const used = new Set(p.map);
          return `${esc(m.name)} <select data-p="${pi}" data-m="${p.masters.indexOf(m)}"><option value="-1">— none —</option>${p.sheet.students.map((s, i) => (used.has(i) ? "" : `<option value="${i}">${esc(s.name)}</option>`)).join("")}</select>`;
        })
        .join("<br>")}</div>`;
    h += `<div class="sc"><table><tr><th>Learner</th>${r.cols.map((d) => `<th>${d.day}<br><span class="mut">${d.dow}</span></th>`).join("")}</tr>${r.grid.map((g) => `<tr><td>${esc(g.name)}</td>${g.cells.map((c) => (c ? `<td class="${c.k}">${c.t}</td>` : '<td class="p">?</td>')).join("")}</tr>`).join("")}</table></div>
    <p style="margin-top:12px"><button class="pri" data-dl="${pi}">Download updated ${esc(p.f.name)}</button> <span class="mut" id="st${pi}"></span></p></div>`;
  });
  out.innerHTML = h;
  $("#mo").onchange = (e) => {
    month = e.target.value;
    render();
  };
  $("#lm").onchange = (e) => {
    lateMode = e.target.value;
    render();
  };
  out.querySelectorAll("select[data-p]").forEach(
    (s) =>
      (s.onchange = (e) => {
        const p = pairs[+s.dataset.p];
        p.map[+s.dataset.m] = +s.value;
        renderKeep();
      }),
  );
  out
    .querySelectorAll("[data-dl]")
    .forEach((b) => (b.onclick = () => download(+b.dataset.dl)));
}
function renderKeep() {
  const keep = pairs;
  const o = buildPairs;
  buildPairs = () => keep;
  render();
  buildPairs = o;
}

// ---------- writing ----------
async function download(pi) {
  const p = pairs[pi],
    st = $("#st" + pi),
    r = plan(p);
  try {
    const zip = await JSZip.loadAsync(p.f.buf);
    const wbx = await zip.file("xl/workbook.xml").async("string");
    const rels = await zip.file("xl/_rels/workbook.xml.rels").async("string");
    const rid = (wbx.match(
      new RegExp(`<sheet [^>]*name="${month}"[^>]*r:id="([^"]+)"`),
    ) || [])[1];
    const tgt = (rels.match(
      new RegExp(`<Relationship [^>]*Id="${rid}"[^>]*Target="([^"]+)"`),
    ) ||
      rels.match(
        new RegExp(`<Relationship [^>]*Target="([^"]+)"[^>]*Id="${rid}"`),
      ) ||
      [])[1];
    if (!tgt) throw new Error("month tab not found in package");
    const path = "xl/" + tgt.replace(/^\/?(xl\/)?/, "");
    let xml = await zip.file(path).async("string");
    let miss = 0;
    r.changes.forEach((ch) => {
      const re = new RegExp(
        `<c r="${ch.addr}"([^>]*?)(?:/>|>(?:(?!</c>)[\\s\\S])*</c>)`,
      );
      const m = xml.match(re);
      if (!m) {
        miss++;
        return;
      }
      const s = (m[1].match(/\ss="(\d+)"/) || [])[1];
      const sa = s ? ` s="${s}"` : "";
      xml = xml.replace(
        re,
        ch.val
          ? `<c r="${ch.addr}"${sa} t="inlineStr"><is><t>${ch.val}</t></is></c>`
          : `<c r="${ch.addr}"${sa}/>`,
      );
    });
    // drop stale cached results of formulas (totals) so they recalculate on open
    xml = xml.replace(
      /(<c r="([A-Z]+)\d+"[^>]*>)(<f[^>]*\/>|<f[^>]*>[\s\S]*?<\/f>)<v>[^<]*<\/v>/g,
      (a, o, c, f) =>
        c === "A" || c === "B" ? a : o.replace(/\st="[^"]*"/, "") + f,
    );
    zip.file(path, xml);
    zip.file(
      "xl/workbook.xml",
      /<calcPr/.test(wbx)
        ? wbx.replace(
            /<calcPr([^>]*?)\/?>/,
            (a, at) =>
              `<calcPr${at.replace(/\sfullCalcOnLoad="[^"]*"/, "")} fullCalcOnLoad="1"/>`,
          )
        : wbx.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>'),
    );
    const blob = await zip.generateAsync({
      type: "blob",
      compression: "DEFLATE",
      mimeType:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = p.f.name.replace(/\.xlsx$/i, "") + "_updated.xlsx";
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 10000);
    st.textContent = `Saved. ${r.changes.length - miss} cells updated${miss ? `, ${miss} not found` : ""}.`;
  } catch (e) {
    st.textContent = "Failed: " + (e.message || e);
  }
}
