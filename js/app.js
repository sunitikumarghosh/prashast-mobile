/* PRASHAST Teacher Onboarding Monitor - SSM, Purba Bardhaman
 * Static app (GitHub Pages). Daily snapshots are stored in IndexedDB and can
 * also be published for everyone through data/manifest.json in the repo. */
'use strict';

const HEADER = 'Prashast On boarding status :: SSM, Purba Bardhaman';
const CREDIT = 'An in-house production of SSM, Purba Bardhaman';
const CAT = { full: '100%', part: '1-99%', zero: '0%', na: 'No teachers' };
const NUM_TYPES = new Set(['idx', 'n', 'p', 'd', 'dp']);

const state = { snaps: [], cur: '', prev: '', charts: [], cache: {} };

// ---------------------------------------------------------------- helpers
const $ = (s, r = document) => r.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const h = html => { const t = document.createElement('template'); t.innerHTML = html.trim(); return t.content.firstElementChild; };
const num = v => { const x = Number(String(v ?? '').replace(/,/g, '')); return isFinite(x) ? x : 0; };
const pct = (o, t) => (t ? (o * 100) / t : 0);
const round2 = n => Math.round(n * 100) / 100;
const fmtN = n => (n == null || n === '' ? '' : Number(n).toLocaleString('en-IN'));
const fmtP = n => (n == null ? '' : round2(n).toFixed(2));
const fmtD = d => (d ? d.split('-').reverse().join('-') : '');
const signed = (n, dec) => { if (n == null) return ''; const v = dec ? round2(n).toFixed(2) : fmtN(n); return n > 0 ? '+' + v : v; };
const pctCls = p => (p >= 80 ? 'good' : p >= 50 ? 'mid' : 'bad');
const slug = s => s.replace(/[^a-z0-9]+/gi, '_').replace(/^_|_$/g, '');
const byDate = (a, b) => a.date.localeCompare(b.date);

function deltaHtml(n, dec, invert) {
    if (n == null) return '<span class="muted">&ndash;</span>';
    const r = round2(n);
    if (r === 0) return '<span class="muted">0</span>';
    const good = (r > 0) !== !!invert;
    return `<span class="${good ? 'up' : 'down'}">${r > 0 ? '&#9650;' : '&#9660;'} ${signed(n, dec)}</span>`;
}

function toast(msg, err) {
    const d = document.createElement('div');
    d.textContent = msg;
    if (err) d.className = 'err';
    $('#toast').append(d);
    setTimeout(() => d.remove(), err ? 7000 : 3500);
}

function dateFromName(name) {
    let m = name.match(/(20\d{2})[-_.](\d{2})[-_.](\d{2})/);
    if (m) return `${m[1]}-${m[2]}-${m[3]}`;
    m = name.match(/(\d{2})[-_.](\d{2})[-_.](20\d{2})/);
    return m ? `${m[3]}-${m[2]}-${m[1]}` : null;
}

function download(blob, name) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.append(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
}

// ---------------------------------------------------------------- storage (IndexedDB, falls back to memory)
const DB = {
    db: null,
    mem: new Map(),
    open() {
        return new Promise(res => {
            setTimeout(res, 3000); // some browsers/private modes never answer; fall back to memory
            try {
                const r = indexedDB.open('prashast-monitor', 1);
                r.onupgradeneeded = () => r.result.createObjectStore('snapshots', { keyPath: 'date' });
                r.onsuccess = () => { this.db = r.result; res(); };
                r.onerror = () => res();
            } catch (e) { res(); }
        });
    },
    tx(mode, fn) {
        return new Promise((res, rej) => {
            const t = this.db.transaction('snapshots', mode);
            const req = fn(t.objectStore('snapshots'));
            t.oncomplete = () => res(req && req.result);
            t.onerror = () => rej(t.error);
        });
    },
    all() { return this.db ? this.tx('readonly', s => s.getAll()) : Promise.resolve([...this.mem.values()]); },
    put(v) { if (!this.db) { this.mem.set(v.date, v); return Promise.resolve(); } return this.tx('readwrite', s => s.put(v)); },
    del(k) { if (!this.db) { this.mem.delete(k); return Promise.resolve(); } return this.tx('readwrite', s => s.delete(k)); },
    clear() { if (!this.db) { this.mem.clear(); return Promise.resolve(); } return this.tx('readwrite', s => s.clear()); },
};

// ---------------------------------------------------------------- Excel parsing
function parseWorkbook(buf) {
    const wb = XLSX.read(buf, { type: 'array' });
    const names = [...wb.SheetNames].sort((a, b) => /all schools/i.test(b) - /all schools/i.test(a));
    for (const name of names) {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '', raw: true });
        const hi = rows.findIndex(r => r.some(c => /school\s*name/i.test(String(c))));
        if (hi < 0) continue;
        const H = rows[hi].map(c => String(c).trim());
        const find = (...tests) => H.findIndex(x => tests.every(t => t(x)));
        const iBlock = find(x => /block/i.test(x), x => !/s\.?\s*no/i.test(x));
        const iSchool = find(x => /school\s*name/i.test(x));
        const iUdise = find(x => /udise/i.test(x), x => !/school\s*name/i.test(x));
        const iTotal = find(x => /total/i.test(x));
        const iOnb = find(x => /onboarded/i.test(x), x => !/%|percent/i.test(x));
        if (iTotal < 0 || iOnb < 0) throw new Error(`Sheet "${name}": "Total Teachers" / "Teachers Onboarded" columns not found`);
        const schools = [];
        for (const r of rows.slice(hi + 1)) {
            let n = String(r[iSchool] ?? '').trim();
            if (!n) continue;
            let u = iUdise >= 0 ? String(r[iUdise] ?? '').trim() : '';
            const k = n.lastIndexOf('|');
            if (!u && k >= 0) { u = n.slice(k + 1).trim(); n = n.slice(0, k).trim(); }
            const b = iBlock >= 0 ? String(r[iBlock] ?? '').trim() : '';
            schools.push({ b: b || 'UNKNOWN', n, u, t: num(r[iTotal]), o: num(r[iOnb]) });
        }
        if (schools.length) return schools;
    }
    throw new Error('No sheet with a "School Name" column was found');
}

// ---------------------------------------------------------------- calculations
function category(x) {
    if (!x.t) return 'na';
    if (x.o >= x.t) return 'full';
    if (x.o <= 0) return 'zero';
    return 'part';
}
const blank = block => ({ block, schools: 0, t: 0, o: 0, full: 0, part: 0, zero: 0, na: 0 });
function addTo(a, x) { a.schools++; a.t += x.t; a.o += x.o; a[x.c]++; }
function finish(a) { a.p = pct(a.o, a.t); a.pend = Math.max(0, a.t - a.o); return a; }

function enrich(s) {
    if (!s || s._e) return s;
    const blocks = new Map();
    const T = blank('District Total');
    s.schools.forEach(x => {
        x.c = category(x);
        x.p = pct(x.o, x.t);
        x.pend = Math.max(0, x.t - x.o);
        x.k = x.u || x.b + '|' + x.n;
        if (!blocks.has(x.b)) blocks.set(x.b, blank(x.b));
        addTo(blocks.get(x.b), x);
        addTo(T, x);
    });
    s.blocks = [...blocks.values()].map(finish).sort((a, b) => a.block.localeCompare(b.block));
    s.total = finish(T);
    s.byKey = new Map(s.schools.map(x => [x.k, x]));
    s.byBlock = new Map(s.blocks.map(b => [b.block, b]));
    Object.defineProperty(s, '_e', { value: true, enumerable: false });
    return s;
}

const snap = d => enrich(state.snaps.find(s => s.date === d));
const cur = () => snap(state.cur);
const prev = () => (state.prev ? snap(state.prev) : null);
const blockNames = () => cur().blocks.map(b => b.block);

function memo(key, fn) {
    if (!(key in state.cache)) state.cache[key] = fn();
    return state.cache[key];
}

// Blocks ranked by onboarded %, with change vs P and Leading/Moderate/Lagging status (top/bottom quarter)
function rankedBlocks(S, P) {
    return memo(`rank|${S.date}|${P ? P.date : ''}`, () => {
        const rows = S.blocks.map(b => {
            const pb = P && P.byBlock.get(b.block);
            return {
                ...b,
                dO: pb ? b.o - pb.o : null,
                dP: pb ? b.p - pb.p : null,
                dFull: pb ? b.full - pb.full : null,
                dZero: pb ? b.zero - pb.zero : null,
            };
        }).sort((a, b) => b.p - a.p || b.o - a.o);
        const n = rows.length, q = Math.ceil(n / 4);
        rows.forEach((r, i) => {
            r.rank = i + 1;
            r.status = i < q ? 'Leading' : i >= n - q ? 'Lagging' : 'Moderate';
        });
        return rows;
    });
}

function schoolRows(S, P) {
    return memo(`sch|${S.date}|${P ? P.date : ''}`, () => S.schools.map(x => {
        const a = P && P.byKey.get(x.k);
        return { ...x, oPrev: a ? a.o : null, pPrev: a ? a.p : null, dO: a ? x.o - a.o : null, cPrev: a ? a.c : null };
    }));
}

function totalsRow(rows, label = 'District Total') {
    const T = blank(label);
    rows.forEach(r => { T.schools += r.schools ?? 1; T.t += r.t; T.o += r.o; ['full', 'part', 'zero', 'na'].forEach(k => { T[k] += r[k] || 0; }); });
    finish(T);
    T.dO = rows.some(r => r.dO != null) ? rows.reduce((s, r) => s + (r.dO || 0), 0) : null;
    return T;
}

// ---------------------------------------------------------------- column definitions
const C = {
    idx: { k: '_i', h: '#', type: 'idx' },
    rank: { k: 'rank', h: 'Rank', type: 'n' },
    block: { k: 'block', h: 'Block', cls: 'nowrap' },
    sb: { k: 'b', h: 'Block', cls: 'nowrap' },
    school: { k: 'n', h: 'School Name', cls: 'wrap' },
    udise: { k: 'u', h: 'UDISE Code', cls: 'nowrap' },
    schools: { k: 'schools', h: 'Schools', type: 'n' },
    t: { k: 't', h: 'Total Teachers', type: 'n' },
    o: { k: 'o', h: 'Onboarded', type: 'n' },
    pend: { k: 'pend', h: 'Pending', type: 'n' },
    p: { k: 'p', h: 'Onboarded %', type: 'p' },
    dO: { k: 'dO', h: 'Change (Onb.)', type: 'd' },
    dP: { k: 'dP', h: 'Change (% pt)', type: 'dp' },
    full: { k: 'full', h: '100% Sch.', type: 'n' },
    part: { k: 'part', h: '1-99% Sch.', type: 'n' },
    zero: { k: 'zero', h: '0% Sch.', type: 'n' },
    status: { k: 'status', h: 'Status', type: 'status' },
    cat: { k: 'c', h: 'Category', type: 'cat' },
};
const col = (base, extra) => ({ ...base, ...extra });

const val = (c, r) => (c.v ? c.v(r) : r[c.k]);

function cellHtml(c, r, i) {
    if (c.type === 'idx') return i + 1;
    const v = val(c, r);
    switch (c.type) {
        case 'n': return fmtN(v);
        case 'p': return v == null ? '' : `<span class="pcell"><span class="pbar"><span class="c-${pctCls(v)}" style="width:${Math.min(100, v)}%"></span></span>${fmtP(v)}</span>`;
        case 'd': return deltaHtml(v, false, c.invert);
        case 'dp': return deltaHtml(v, true, c.invert);
        case 'cat': return v ? `<span class="badge ${v}">${CAT[v]}</span>` : '';
        case 'status': return v ? `<span class="badge ${v}">${v}</span>` : '';
        default: return esc(c.show ? c.show(v) : v);
    }
}

// Value written to Excel (raw numbers) or PDF (formatted text)
function cellOut(c, r, i, pdf) {
    if (c.type === 'idx') return i + 1;
    const v = val(c, r);
    if (v == null) return '';
    switch (c.type) {
        case 'p': return pdf ? fmtP(v) : round2(v);
        case 'dp': return pdf ? signed(v, true) : round2(v);
        case 'd': return pdf ? signed(v) : v;
        case 'n': return pdf ? fmtN(v) : v;
        case 'cat': return CAT[v] || v;
        default: return c.show ? c.show(v) : v;
    }
}

// ---------------------------------------------------------------- table component
function makeTable({ cols, rows, sort = null, dir = 'desc', pageSize = 0, onRow = null, foot = null, empty = 'No records found' }) {
    const wrap = h('<div><div class="table-wrap"><table><thead></thead><tbody></tbody><tfoot></tfoot></table></div><div class="pager"></div></div>');
    const thead = $('thead', wrap), tbody = $('tbody', wrap), tfoot = $('tfoot', wrap), pager = $('.pager', wrap);
    let sk = sort, sd = dir, page = 0;

    const sorted = () => {
        if (!sk) return rows;
        const c = cols.find(x => x.k === sk);
        return [...rows].sort((a, b) => {
            const x = val(c, a), y = val(c, b);
            const r = NUM_TYPES.has(c.type)
                ? (x ?? -1e15) - (y ?? -1e15)
                : String(x ?? '').localeCompare(String(y ?? ''), 'en', { numeric: true });
            return sd === 'asc' ? r : -r;
        });
    };
    const tdCls = c => [NUM_TYPES.has(c.type) ? 'num' : '', c.cls || ''].join(' ').trim();

    function draw() {
        const all = sorted();
        const pages = pageSize ? Math.max(1, Math.ceil(all.length / pageSize)) : 1;
        page = Math.min(page, pages - 1);
        const start = pageSize ? page * pageSize : 0;
        const list = pageSize ? all.slice(start, start + pageSize) : all;

        thead.innerHTML = '<tr>' + cols.map(c => {
            const cls = [NUM_TYPES.has(c.type) ? 'num' : '', c.type === 'idx' ? '' : 'sortable', sk === c.k ? 'sorted ' + sd : ''].join(' ');
            return `<th data-k="${c.k}" class="${cls}">${esc(c.h)}</th>`;
        }).join('') + '</tr>';

        tbody.innerHTML = list.length
            ? list.map((r, i) => `<tr data-i="${start + i}" class="${onRow ? 'clickable' : ''}">${cols.map(c => `<td class="${tdCls(c)}">${cellHtml(c, r, start + i)}</td>`).join('')}</tr>`).join('')
            : `<tr><td colspan="${cols.length}" class="empty">${empty}</td></tr>`;

        tfoot.innerHTML = foot && list.length
            ? '<tr>' + cols.map(c => `<td class="${tdCls(c)}">${c.type === 'idx' ? '' : cellHtml(c, foot, 0)}</td>`).join('') + '</tr>'
            : '';

        pager.innerHTML = `<span>${fmtN(all.length)} record${all.length === 1 ? '' : 's'}</span>` + (pages > 1
            ? `<span class="pg"><button class="btn sm" data-p="-1" ${page === 0 ? 'disabled' : ''}>&lsaquo; Prev</button> Page ${page + 1} of ${pages} <button class="btn sm" data-p="1" ${page >= pages - 1 ? 'disabled' : ''}>Next &rsaquo;</button></span>`
            : '');
    }

    thead.addEventListener('click', e => {
        const th = e.target.closest('th.sortable');
        if (!th) return;
        const k = th.dataset.k;
        if (sk === k) sd = sd === 'asc' ? 'desc' : 'asc';
        else { sk = k; sd = NUM_TYPES.has(cols.find(c => c.k === k).type) ? 'desc' : 'asc'; }
        page = 0;
        draw();
    });
    pager.addEventListener('click', e => {
        const b = e.target.closest('[data-p]');
        if (!b) return;
        page += Number(b.dataset.p);
        draw();
        wrap.scrollIntoView({ block: 'start', behavior: 'smooth' });
    });
    if (onRow) tbody.addEventListener('click', e => {
        const tr = e.target.closest('tr[data-i]');
        if (tr) onRow(sorted()[Number(tr.dataset.i)]);
    });

    draw();
    return {
        el: wrap,
        cols,
        get foot() { return foot; },
        rows: sorted,
        setRows(r, f) { rows = r; foot = f === undefined ? foot : f; page = 0; draw(); },
    };
}

// ---------------------------------------------------------------- export (Excel / PDF)
function exportExcel(spec) {
    const { title, cols, rows, foot, sub } = spec();
    const n = cols.length;
    const aoa = [
        [HEADER],
        [title],
        [`Data as on: ${fmtD(state.cur)}${sub ? '  |  ' + sub : ''}`],
        [`Generated on: ${new Date().toLocaleString('en-IN')}  |  ${CREDIT}`],
        [],
        cols.map(c => c.h),
        ...rows.map((r, i) => cols.map(c => cellOut(c, r, i, false))),
    ];
    if (foot) aoa.push(cols.map((c, i) => (c.type === 'idx' ? (i === 0 ? 'Total' : '') : cellOut(c, foot, 0, false))));
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!merges'] = [0, 1, 2, 3].map(r => ({ s: { r, c: 0 }, e: { r, c: Math.max(0, n - 1) } }));
    ws['!cols'] = cols.map((c, j) => {
        let w = c.h.length;
        for (let i = 6; i < Math.min(aoa.length, 400); i++) w = Math.max(w, String(aoa[i][j] ?? '').length);
        return { wch: Math.min(60, w + 2) };
    });
    if (rows.length) ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 5, c: 0 }, e: { r: 5 + rows.length, c: n - 1 } }) };
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, slug(title).slice(0, 31) || 'Report');
    XLSX.writeFile(wb, `${slug(title)}_${state.cur}.xlsx`);
}

function exportPDF(spec) {
    const { title, cols, rows, foot, sub } = spec();
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: cols.length > 6 ? 'landscape' : 'portrait', unit: 'mm', format: 'a4' });
    const W = doc.internal.pageSize.getWidth(), H = doc.internal.pageSize.getHeight();
    const TP = '{total_pages}';
    const generated = new Date().toLocaleString('en-IN');
    const colStyles = {};
    cols.forEach((c, i) => { if (NUM_TYPES.has(c.type)) colStyles[i] = { halign: 'right' }; });

    doc.autoTable({
        head: [cols.map(c => c.h)],
        body: rows.map((r, i) => cols.map(c => cellOut(c, r, i, true))),
        foot: foot ? [cols.map((c, i) => (c.type === 'idx' ? (i === 0 ? 'Total' : '') : cellOut(c, foot, 0, true)))] : undefined,
        showFoot: 'lastPage',
        startY: 34,
        margin: { top: 34, left: 8, right: 8, bottom: 14 },
        styles: { fontSize: cols.length > 10 ? 7 : 8, cellPadding: 1.6, overflow: 'linebreak', lineColor: [220, 222, 235], lineWidth: 0.1 },
        headStyles: { fillColor: [26, 35, 126], textColor: 255, fontStyle: 'bold' },
        footStyles: { fillColor: [232, 234, 246], textColor: 20, fontStyle: 'bold' },
        alternateRowStyles: { fillColor: [246, 247, 252] },
        columnStyles: colStyles,
        didParseCell: d => {
            if (d.section !== 'body' && NUM_TYPES.has(cols[d.column.index].type)) d.cell.styles.halign = 'right';
        },
        didDrawPage: () => {
            doc.setFont('helvetica', 'bold');
            doc.setFontSize(13);
            doc.setTextColor(26, 35, 126);
            doc.text(HEADER, W / 2, 11, { align: 'center' });
            doc.setFontSize(11);
            doc.setTextColor(30);
            doc.text(title, W / 2, 18, { align: 'center' });
            doc.setFont('helvetica', 'normal');
            doc.setFontSize(8.5);
            doc.text(`Data as on: ${fmtD(state.cur)}${sub ? '   |   ' + sub : ''}`, W / 2, 24, { align: 'center' });
            doc.setDrawColor(26, 35, 126);
            doc.setLineWidth(0.5);
            doc.line(8, 28, W - 8, 28);
            doc.setFontSize(7.5);
            doc.setTextColor(110);
            doc.text(`Generated on ${generated}`, 8, H - 6);
            doc.text(CREDIT, W / 2, H - 6, { align: 'center' });
            doc.text(`Page ${doc.internal.getNumberOfPages()} of ${TP}`, W - 8, H - 6, { align: 'right' });
        },
    });
    if (typeof doc.putTotalPages === 'function') doc.putTotalPages(TP);
    doc.save(`${slug(title)}_${state.cur}.pdf`);
}

function exportBtns(spec) {
    const d = h('<div class="actions"><button class="btn xls">&#128215; Excel</button><button class="btn pdf">&#128213; PDF</button></div>');
    const run = fn => () => {
        try { fn(spec); } catch (e) { console.error(e); toast('Export failed: ' + e.message, true); }
    };
    d.children[0].onclick = run(exportExcel);
    d.children[1].onclick = run(exportPDF);
    return d;
}
// Exported reports are always ordered by block name, then school name
const cmpText = (a, b) => String(a ?? '').localeCompare(String(b ?? ''), 'en', { numeric: true, sensitivity: 'base' });
function exportOrder(rows) {
    if (!rows.length) return rows;
    if ('b' in rows[0]) return [...rows].sort((x, y) => cmpText(x.b, y.b) || cmpText(x.n, y.n));
    if ('block' in rows[0]) return [...rows].sort((x, y) => cmpText(x.block, y.block));
    return rows;
}
const tableSpec = (tbl, title, sub) => () => ({
    title: typeof title === 'function' ? title() : title,
    sub: typeof sub === 'function' ? sub() : sub,
    cols: tbl.cols,
    rows: exportOrder(tbl.rows()),
    foot: tbl.foot,
});

// ---------------------------------------------------------------- layout helpers
function pageHead(v, title, sub, actions) {
    const d = h(`<div class="page-head"><div><h1></h1>${sub ? `<p class="sub">${sub}</p>` : ''}</div></div>`);
    $('h1', d).textContent = title;
    if (actions) d.append(actions);
    v.append(d);
    return d;
}
function card(parent, title, actions, flush) {
    const c = h(`<section class="card"><div class="card-head"><h2></h2></div><div class="card-body${flush ? ' flush' : ''}"></div></section>`);
    $('h2', c).textContent = title;
    if (actions) $('.card-head', c).append(actions);
    parent.append(c);
    return $('.card-body', c);
}
function tableCard(parent, title, opts, sub) {
    const tbl = makeTable(opts);
    const body = card(parent, title, exportBtns(tableSpec(tbl, title, sub)), true);
    body.append(tbl.el);
    return tbl;
}
const subLine = (S, P) => `Data as on <b>${fmtD(S.date)}</b>${P ? ` &middot; changes compared with <b>${fmtD(P.date)}</b>` : ''}`;
const plainSub = (S, P) => (P ? `Compared with: ${fmtD(P.date)}` : '');

function chart(parent, height, config) {
    const box = h(`<div class="chart-box" style="height:${height}px"><canvas></canvas></div>`);
    parent.append(box);
    config.options = Object.assign({ responsive: true, maintainAspectRatio: false, animation: { duration: 300 } }, config.options);
    state.charts.push(new Chart($('canvas', box), config));
}
const COLORS = { good: '#2e7d32', mid: '#e08600', bad: '#c62828', na: '#9e9e9e', brand: '#1a237e', brand2: '#3949ab' };
const statusColor = s => (s === 'Leading' ? COLORS.good : s === 'Lagging' ? COLORS.bad : COLORS.mid);

const ICONS = {
    teacher: '<circle cx="12" cy="8" r="4"/><path d="M4 21a8 8 0 0 1 16 0"/>',
    check: '<circle cx="12" cy="12" r="9"/><path d="m8 12 3 3 5-6"/>',
    clock: '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
    school: '<path d="M3 21h18M5 21V10l7-5 7 5v11M10 21v-5h4v5"/>',
    full: '<path d="M20 6 9 17l-5-5"/>',
    part: '<circle cx="12" cy="12" r="9"/><path d="M12 3v9h9"/>',
    zero: '<circle cx="12" cy="12" r="9"/><path d="M5.6 5.6l12.8 12.8"/>',
    up: '<path d="m3 17 6-6 4 4 8-8"/><path d="M14 7h7v7"/>',
    down: '<path d="m3 7 6 6 4-4 8 8"/><path d="M14 17h7v-7"/>',
    info: '<circle cx="12" cy="12" r="9"/><path d="M12 11v5M12 8h.01"/>',
    flag: '<path d="M12 3 2 20h20L12 3Z"/><path d="M12 10v4M12 17h.01"/>',
};
const icon = n => `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ICONS[n]}</svg>`;

function statCard({ label, value, delta, invert, pctDelta, foot, ic, tone, href }) {
    const d = delta == null ? '' : `<span class="stat-delta">${deltaHtml(delta, pctDelta, invert)} <small>since ${fmtD(state.prev)}</small></span>`;
    return `<a class="stat ${tone}"${href ? ` href="${href}"` : ''}>
        <span class="stat-icon">${icon(ic)}</span>
        <span class="stat-body"><span class="stat-label">${label}</span><span class="stat-val">${value}</span>${d}<span class="stat-foot">${foot}</span></span>
    </a>`;
}

function pageDashboard(v) {
    const S = cur(), P = prev(), T = S.total, PT = P && P.total;
    const d = k => (PT ? T[k] - PT[k] : null);
    const ranked = rankedBlocks(S, P);
    const count = s => ranked.filter(b => b.status === s).length;
    pageHead(v, 'District Dashboard', subLine(S, P));

    // hero: district progress
    const chg = PT ? `<span class="hero-chip ${T.p >= PT.p ? 'pos' : 'neg'}">${icon(T.p >= PT.p ? 'up' : 'down')} ${signed(T.p - PT.p, true)} pt &middot; ${signed(T.o - PT.o)} teachers since ${fmtD(P.date)}</span>` : '';
    v.append(h(`<section class="hero">
        <div class="ring" style="--p:${Math.min(100, T.p)}"><div class="ring-in"><b>${fmtP(T.p)}%</b><small>onboarded</small></div></div>
        <div class="hero-main">
            <div class="hero-label">District teacher onboarding &middot; Purba Bardhaman</div>
            <div class="hero-title"><b>${fmtN(T.o)}</b> of ${fmtN(T.t)} teachers onboarded</div>
            <div class="hero-bar"><span style="width:${Math.min(100, T.p)}%"></span></div>
            <div class="hero-meta">${fmtN(T.pend)} teachers still to be onboarded across ${fmtN(T.schools)} schools</div>
            ${chg}
        </div>
        <div class="hero-stats">
            <div><b>${S.blocks.length}</b><span>Blocks</span></div>
            <div><b>${fmtN(T.schools)}</b><span>Schools</span></div>
            <div><b class="g">${count('Leading')}</b><span>Leading blocks</span></div>
            <div><b class="r">${count('Lagging')}</b><span>Lagging blocks</span></div>
        </div>
    </section>`));

    // teacher KPIs
    v.append(h(`<div class="stat-grid">
        ${statCard({ label: 'Total Teachers', value: fmtN(T.t), delta: d('t'), foot: `in ${fmtN(T.schools)} schools`, ic: 'teacher', tone: 'blue' })}
        ${statCard({ label: 'Teachers Onboarded', value: fmtN(T.o), delta: d('o'), foot: `${fmtP(T.p)}% of all teachers`, ic: 'check', tone: 'green' })}
        ${statCard({ label: 'Pending Teachers', value: fmtN(T.pend), delta: d('pend'), invert: true, foot: `${fmtP(100 - T.p)}% yet to onboard`, ic: 'clock', tone: 'red' })}
        ${statCard({ label: 'Onboarded %', value: fmtP(T.p) + '%', delta: d('p'), pctDelta: true, foot: `best block ${fmtP(ranked[0].p)}%`, ic: 'up', tone: 'indigo' })}
    </div>`));

    // school categories
    const catCard = (key, label, ic, invert) => {
        const n = T[key], share = pct(n, T.schools);
        const dd = d(key);
        return `<a class="cat-card ${key}" href="#/schools/${key}">
            <div class="cat-top"><span class="cat-icon">${icon(ic)}</span><span class="cat-label">${label}</span><span class="cat-link">View list &rsaquo;</span></div>
            <div class="cat-val">${fmtN(n)} <small>schools</small></div>
            <div class="cat-bar"><span style="width:${share}%"></span></div>
            <div class="cat-foot"><span>${fmtP(share)}% of schools</span>${dd == null ? '' : deltaHtml(dd, false, invert)}</div>
        </a>`;
    };
    v.append(h(`<div class="cat-grid">
        ${catCard('full', '100% Onboarded Schools', 'full')}
        ${catCard('part', 'Partially Onboarded (1-99%)', 'part')}
        ${catCard('zero', '0% Onboarded Schools', 'zero', true)}
    </div>`));

    // key highlights
    const best = ranked[0], worst = ranked[ranked.length - 1];
    const notes = [
        ['up', `Highest onboarding: <b>${esc(best.block)}</b> at ${fmtP(best.p)}%; lowest: <b>${esc(worst.block)}</b> at ${fmtP(worst.p)}%.`],
        ['zero', `<b>${fmtN(T.zero)}</b> schools (${fmtP(pct(T.zero, T.schools))}%) have not onboarded any teacher. <a href="#/schools/zero">See list &rsaquo;</a>`],
        ['flag', `Most pending teachers: <b>${esc([...ranked].sort((a, b) => b.pend - a.pend)[0].block)}</b> (${fmtN([...ranked].sort((a, b) => b.pend - a.pend)[0].pend)} pending).`],
    ];
    if (P) {
        const sr = schoolRows(S, P);
        const newFull = sr.filter(x => x.c === 'full' && x.cPrev && x.cPrev !== 'full').length;
        const outZero = sr.filter(x => x.cPrev === 'zero' && x.c !== 'zero').length;
        const improved = sr.filter(x => x.dO > 0).length;
        const regressed = sr.filter(x => x.dO < 0).length;
        const still = ranked.filter(b => b.dO === 0).map(b => b.block);
        const gainer = [...ranked].sort((a, b) => (b.dO || 0) - (a.dO || 0))[0];
        notes.unshift(['check', `Since ${fmtD(P.date)}: <b>${fmtN(improved)}</b> schools improved, <b>${fmtN(newFull)}</b> became 100% and <b>${fmtN(outZero)}</b> moved out of 0%. <a href="#/compare">Details &rsaquo;</a>`]);
        if (gainer && gainer.dO > 0) notes.push(['up', `Top gainer since ${fmtD(P.date)}: <b>${esc(gainer.block)}</b> (${signed(gainer.dO)} teachers).`]);
        if (still.length) notes.push(['info', `No progress since ${fmtD(P.date)} in ${still.length} block(s): ${still.map(esc).join(', ')}.`]);
        if (regressed) notes.push(['down', `<span class="down">${fmtN(regressed)} school(s) show fewer onboarded teachers than on ${fmtD(P.date)}</span> &mdash; please verify.`]);
    }
    card(v, 'Key Highlights').append(h(`<ul class="highlights">${notes.map(([i, t]) => `<li><span class="hl-icon">${icon(i)}</span><span>${t}</span></li>`).join('')}</ul>`));

    // charts
    const g = h('<div class="grid-2"></div>');
    v.append(g);
    chart(card(g, 'Schools by Onboarding Category'), 260, {
        type: 'doughnut',
        data: {
            labels: ['100%', '1-99%', '0%', 'No teachers'],
            datasets: [{ data: [T.full, T.part, T.zero, T.na], backgroundColor: [COLORS.good, COLORS.mid, COLORS.bad, COLORS.na], borderWidth: 2, borderColor: '#fff' }],
        },
        options: { plugins: { legend: { position: window.innerWidth < 600 ? 'bottom' : 'right', labels: { usePointStyle: true, padding: 14 } } }, cutout: '64%' },
    });
    const bands = [['Below 40%', 0, 40], ['40-60%', 40, 60], ['60-80%', 60, 80], ['80% and above', 80, 101]];
    chart(card(g, 'Blocks by Onboarding Level'), 260, {
        type: 'bar',
        data: {
            labels: bands.map(b => b[0]),
            datasets: [{ label: 'Blocks', data: bands.map(([, lo, hi]) => ranked.filter(b => b.p >= lo && b.p < hi).length), backgroundColor: [COLORS.bad, COLORS.mid, COLORS.brand2, COLORS.good], borderRadius: 6, maxBarThickness: 56 }],
        },
        options: { plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true, ticks: { precision: 0 }, title: { display: true, text: 'Number of blocks' } }, x: { grid: { display: false } } } },
    });
    if (state.snaps.length > 1) {
        const all = state.snaps.map(s => enrich(s));
        chart(card(v, 'District Daily Trend'), 280, {
            type: 'line',
            data: {
                labels: all.map(s => fmtD(s.date)),
                datasets: [
                    { label: 'Onboarded %', data: all.map(s => round2(s.total.p)), borderColor: COLORS.brand, backgroundColor: 'rgba(26,35,126,.10)', fill: true, tension: .3, yAxisID: 'y', pointRadius: 3 },
                    { label: '0% schools', data: all.map(s => s.total.zero), borderColor: COLORS.bad, tension: .3, yAxisID: 'y1', pointRadius: 3 },
                    { label: '100% schools', data: all.map(s => s.total.full), borderColor: COLORS.good, tension: .3, yAxisID: 'y1', pointRadius: 3 },
                ],
            },
            options: { interaction: { mode: 'index', intersect: false }, plugins: { legend: { labels: { usePointStyle: true } } }, scales: { y: { ticks: { callback: x => x + '%' } }, y1: { position: 'right', grid: { drawOnChartArea: false }, title: { display: true, text: 'Schools' } } } },
        });
    }

    const g2 = h('<div class="grid-2" style="margin-top:16px"></div>');
    v.append(g2);
    const mini = [C.rank, C.block, C.o, C.p, col(C.dO, { h: 'Change' })];
    const go = r => { location.hash = '#/schools?block=' + encodeURIComponent(r.block); };
    tableCard(g2, 'Top 5 Leading Blocks', { cols: mini, rows: ranked.slice(0, 5), onRow: go }, plainSub(S, P));
    tableCard(g2, 'Bottom 5 Lagging Blocks', { cols: mini, rows: ranked.slice(-5).reverse(), onRow: go }, plainSub(S, P));
}

function pageBlocks(v) {
    const S = cur(), P = prev();
    const rows = rankedBlocks(S, P);
    const cols = [C.rank, C.block, C.schools, C.t, C.o, col(C.pend, { invert: true }), C.p, C.dO, C.dP, C.full, C.part, C.zero, C.status];
    const foot = Object.assign(totalsRow(rows), { dP: P ? S.total.p - P.total.p : null });
    const tbl = makeTable({ cols, rows, foot, onRow: r => { location.hash = '#/schools?block=' + encodeURIComponent(r.block); } });
    pageHead(v, 'Block-wise Onboarding Status', subLine(S, P) + ' &middot; tap a block to see its schools', exportBtns(tableSpec(tbl, 'Block-wise Onboarding Status', plainSub(S, P))));
    v.append(h(`<p class="help">Status: <span class="badge Leading">Leading</span> = top quarter of blocks by onboarded %, <span class="badge Lagging">Lagging</span> = bottom quarter, <span class="badge Moderate">Moderate</span> = the rest.</p>`));
    card(v, 'All Blocks', null, true).append(tbl.el);

    const byName = [...rows].sort((a, b) => a.block.localeCompare(b.block));
    chart(card(v, 'School categories by block'), Math.max(320, rows.length * 20 + 60), {
        type: 'bar',
        data: {
            labels: byName.map(b => b.block),
            datasets: [
                { label: '100%', data: byName.map(b => b.full), backgroundColor: COLORS.good },
                { label: '1-99%', data: byName.map(b => b.part), backgroundColor: COLORS.mid },
                { label: '0%', data: byName.map(b => b.zero), backgroundColor: COLORS.bad },
                { label: 'No teachers', data: byName.map(b => b.na), backgroundColor: COLORS.na },
            ],
        },
        options: { indexAxis: 'y', scales: { x: { stacked: true }, y: { stacked: true, ticks: { autoSkip: false, font: { size: 10 } } } }, interaction: { mode: 'index', intersect: false } },
    });
}

const SCHOOL_TITLES = {
    all: 'All Schools',
    full: '100% Onboarded Schools',
    part: 'Partially Onboarded Schools (1-99%)',
    zero: '0% Onboarded Schools',
    na: 'Schools with No Teachers',
};

function pageSchools(v, r) {
    const S = cur(), P = prev();
    const catKey = SCHOOL_TITLES[r.arg] ? r.arg : 'all';
    const block = r.q.get('block') || '';
    let search = '';
    const all = schoolRows(S, P);
    const filtered = () => all.filter(x => (!block || x.b === block) && (catKey === 'all' || x.c === catKey) &&
        (!search || (x.n + ' ' + x.u).toLowerCase().includes(search)));
    const title = SCHOOL_TITLES[catKey] + (block ? ` - ${block} Block` : '');
    const go = (c, b) => { location.hash = `#/schools${c && c !== 'all' ? '/' + c : ''}${b ? '?block=' + encodeURIComponent(b) : ''}`; };

    const cols = [C.idx];
    if (!block) cols.push(C.sb);
    cols.push(C.school, C.udise, C.t, C.o, col(C.pend, { invert: true }), C.p);
    if (P) cols.push({ k: 'oPrev', h: `Onb. on ${fmtD(P.date)}`, type: 'n' }, C.dO);
    if (catKey === 'all') cols.push(C.cat);

    const defaultSort = { zero: ['t', 'desc'], part: ['p', 'asc'] }[catKey] || [null];
    const tbl = makeTable({ cols, rows: filtered(), sort: defaultSort[0], dir: defaultSort[1], pageSize: 50 });
    pageHead(v, title, subLine(S, P), exportBtns(tableSpec(tbl, () => title + (search ? ` (search: ${search})` : ''), plainSub(S, P))));

    const f = h(`<div class="filters">
        <label>Block<select id="fBlock"><option value="">All blocks</option>${blockNames().map(b => `<option${b === block ? ' selected' : ''}>${esc(b)}</option>`).join('')}</select></label>
        <label>Category<select id="fCat">${Object.keys(SCHOOL_TITLES).map(k => `<option value="${k}"${k === catKey ? ' selected' : ''}>${k === 'all' ? 'All categories' : CAT[k] + ' onboarded'}</option>`).join('')}</select></label>
        <label>Search school / UDISE<input type="search" id="fSearch" placeholder="Type to search..."></label>
    </div>`);
    v.append(f);
    $('#fBlock', f).onchange = e => go(catKey, e.target.value);
    $('#fCat', f).onchange = e => go(e.target.value, block);

    const chips = h('<div class="chips"></div>');
    v.append(chips);
    const drawChips = rows => {
        const t = totalsRow(rows);
        chips.innerHTML = `<span class="chip">Schools <b>${fmtN(rows.length)}</b></span><span class="chip">Teachers <b>${fmtN(t.t)}</b></span>
            <span class="chip">Onboarded <b>${fmtN(t.o)}</b></span><span class="chip">Pending <b>${fmtN(t.pend)}</b></span><span class="chip">Onboarded % <b>${fmtP(t.p)}</b></span>`;
    };
    drawChips(tbl.rows());
    let timer;
    $('#fSearch', f).oninput = e => {
        clearTimeout(timer);
        timer = setTimeout(() => { search = e.target.value.trim().toLowerCase(); const rows = filtered(); tbl.setRows(rows); drawChips(rows); }, 200);
    };
    card(v, 'Schools', null, true).append(tbl.el);
}

function pageCompare(v, r) {
    if (state.snaps.length < 2) {
        pageHead(v, 'Day-wise Comparison');
        v.append(h('<div class="card"><div class="card-body">Upload data for at least two dates to compare.</div></div>'));
        return;
    }
    const dates = state.snaps.map(s => s.date);
    const bDate = r.q.get('b') || state.cur;
    const aDate = r.q.get('a') || state.prev || dates[Math.max(0, dates.indexOf(bDate) - 1)];
    const A = snap(aDate), B = snap(bDate);
    const opts = sel => dates.slice().reverse().map(d => `<option value="${d}"${d === sel ? ' selected' : ''}>${fmtD(d)}</option>`).join('');
    const sub = `${fmtD(A.date)} vs ${fmtD(B.date)}`;
    pageHead(v, 'Day-wise Comparison', `Changes from <b>${fmtD(A.date)}</b> to <b>${fmtD(B.date)}</b>`);

    const f = h(`<div class="filters"><label>From date<select id="cA">${opts(aDate)}</select></label><label>To date<select id="cB">${opts(bDate)}</select></label></div>`);
    v.append(f);
    const go = () => { location.hash = `#/compare?a=${$('#cA').value}&b=${$('#cB').value}`; };
    $('#cA', f).onchange = go;
    $('#cB', f).onchange = go;

    const TA = A.total, TB = B.total;
    const k = (label, key, o = {}) => `<div class="kpi ${o.cls || ''}"><span class="kpi-label">${label}</span><span class="kpi-val">${o.pct ? fmtP(TB[key]) + '%' : fmtN(TB[key])}</span><span class="kpi-delta">${deltaHtml(TB[key] - TA[key], o.pct, o.invert)} <small>from ${o.pct ? fmtP(TA[key]) + '%' : fmtN(TA[key])}</small></span></div>`;
    v.append(h(`<div class="kpis">${k('Teachers Onboarded', 'o', { cls: 'good' })}${k('Onboarded %', 'p', { pct: true })}${k('Total Teachers', 't')}${k('100% Schools', 'full', { cls: 'good' })}${k('1-99% Schools', 'part', { cls: 'mid' })}${k('0% Schools', 'zero', { invert: true, cls: 'bad' })}</div>`));

    // block comparison
    const names = [...new Set([...A.blocks, ...B.blocks].map(b => b.block))];
    const rows = names.map(n => {
        const a = A.byBlock.get(n) || finish(blank(n)), b = B.byBlock.get(n) || finish(blank(n));
        return { block: n, t: b.t, oA: a.o, o: b.o, dO: b.o - a.o, pA: a.p, p: b.p, dP: b.p - a.p, fullA: a.full, full: b.full, zeroA: a.zero, zero: b.zero, dZero: b.zero - a.zero };
    });
    const tot = { block: 'District Total', t: TB.t, oA: TA.o, o: TB.o, dO: TB.o - TA.o, pA: TA.p, p: TB.p, dP: TB.p - TA.p, fullA: TA.full, full: TB.full, zeroA: TA.zero, zero: TB.zero, dZero: TB.zero - TA.zero };
    const dA = fmtD(A.date), dB = fmtD(B.date);
    tableCard(v, 'Block-wise Comparison', {
        cols: [C.idx, C.block, C.t, { k: 'oA', h: `Onb. ${dA}`, type: 'n' }, { k: 'o', h: `Onb. ${dB}`, type: 'n' }, C.dO,
            { k: 'pA', h: `% ${dA}`, type: 'p' }, { k: 'p', h: `% ${dB}`, type: 'p' }, C.dP,
            { k: 'fullA', h: `100% ${dA}`, type: 'n' }, { k: 'full', h: `100% ${dB}`, type: 'n' },
            { k: 'zeroA', h: `0% ${dA}`, type: 'n' }, { k: 'zero', h: `0% ${dB}`, type: 'n' }, { k: 'dZero', h: 'Change (0% sch.)', type: 'd', invert: true }],
        rows, sort: 'dO', foot: tot,
    }, sub);

    // school movements
    const sr = schoolRows(B, A);
    const groups = [
        ['improved', 'Schools improved', x => x.dO > 0],
        ['newfull', 'Became 100%', x => x.c === 'full' && x.cPrev && x.cPrev !== 'full'],
        ['outzero', 'Moved out of 0%', x => x.cPrev === 'zero' && x.c !== 'zero'],
        ['stillzero', 'Still 0% on both dates', x => x.c === 'zero' && x.cPrev === 'zero'],
        ['regressed', 'Onboarded count decreased', x => x.dO < 0],
        ['new', 'New in later file', x => x.cPrev == null],
    ].map(([key, label, fn]) => ({ key, label, rows: sr.filter(fn) }));
    let active = groups[0];
    const cols = [C.idx, C.sb, C.school, C.udise, C.t, { k: 'oPrev', h: `Onb. ${dA}`, type: 'n' }, { k: 'o', h: `Onb. ${dB}`, type: 'n' }, C.dO, C.p, C.cat];
    const tbl = makeTable({ cols, rows: active.rows, sort: 'dO', pageSize: 50 });
    const body = card(v, 'School-level Movement', exportBtns(() => ({ ...tableSpec(tbl, `School Movement - ${active.label}`, sub)() })), true);
    const tabs = h(`<div class="tabs">${groups.map((g, i) => `<button class="tab${i === 0 ? ' active' : ''}" data-i="${i}">${g.label}<span class="count">${fmtN(g.rows.length)}</span></button>`).join('')}</div>`);
    tabs.onclick = e => {
        const b = e.target.closest('.tab');
        if (!b) return;
        active = groups[Number(b.dataset.i)];
        tabs.querySelectorAll('.tab').forEach(t => t.classList.toggle('active', t === b));
        tbl.setRows(active.rows);
    };
    body.append(tabs, tbl.el);
}

function pageTrend(v) {
    const all = state.snaps.map(s => enrich(s));
    const rows = all.map((s, i) => {
        const p = i ? all[i - 1].total : null, t = s.total;
        return { date: s.date, schools: t.schools, t: t.t, o: t.o, dO: p ? t.o - p.o : null, p: t.p, dP: p ? t.p - p.p : null, full: t.full, part: t.part, zero: t.zero, pend: t.pend };
    }).reverse();
    pageHead(v, 'Daily Trend', `${all.length} day(s) of data, from <b>${fmtD(all[0].date)}</b> to <b>${fmtD(all[all.length - 1].date)}</b>`);

    if (all.length > 1) {
        chart(card(v, 'District progress by date'), 280, {
            type: 'bar',
            data: {
                labels: all.map(s => fmtD(s.date)),
                datasets: [
                    { type: 'line', label: 'Onboarded %', data: all.map(s => round2(s.total.p)), borderColor: COLORS.brand, backgroundColor: COLORS.brand, yAxisID: 'y1', tension: .25 },
                    { label: 'Teachers onboarded', data: all.map(s => s.total.o), backgroundColor: 'rgba(57,73,171,.35)', yAxisID: 'y' },
                ],
            },
            options: { interaction: { mode: 'index', intersect: false }, scales: { y: { beginAtZero: true }, y1: { position: 'right', grid: { drawOnChartArea: false }, ticks: { callback: x => x + '%' } } } },
        });
    }

    tableCard(v, 'District Date-wise Summary', {
        cols: [{ k: 'date', h: 'Date', cls: 'nowrap', show: fmtD }, C.schools, C.t, C.o, col(C.dO, { h: 'Daily change' }), C.p, C.dP, C.full, C.part, C.zero, col(C.pend, { invert: true })],
        rows,
    });

    // block x date matrix
    const metricBox = card(v, 'Block-wise Trend (by date)');
    const sel = h('<div class="filters"><label>Show<select><option value="p">Onboarded %</option><option value="o">Teachers onboarded</option><option value="zero">0% schools</option><option value="full">100% schools</option></select></label></div>');
    metricBox.append(sel);
    const holder = h('<div></div>');
    metricBox.append(holder);
    const draw = metric => {
        holder.innerHTML = '';
        const last = all[all.length - 1], first = all[0];
        const names = last.blocks.map(b => b.block);
        const type = metric === 'p' ? 'p' : 'n';
        const cols = [C.idx, C.block, ...all.map(s => ({ k: s.date, h: fmtD(s.date), type })),
            { k: 'chg', h: 'Change (first to last)', type: metric === 'p' ? 'dp' : 'd', invert: metric === 'zero' }];
        const rows = names.map(n => {
            const r = { block: n };
            all.forEach(s => { const b = s.byBlock.get(n); r[s.date] = b ? b[metric] : null; });
            r.chg = r[last.date] != null && r[first.date] != null ? r[last.date] - r[first.date] : null;
            return r;
        });
        const label = sel.querySelector('option:checked').textContent;
        const tbl = makeTable({ cols, rows, sort: last.date });
        const bar = h('<div class="actions" style="margin-bottom:8px"></div>');
        bar.append(exportBtns(tableSpec(tbl, `Block-wise Trend - ${label}`, `${fmtD(first.date)} to ${fmtD(last.date)}`)));
        holder.append(bar, tbl.el);
    };
    $('select', sel).onchange = e => draw(e.target.value);
    draw('p');
}

function pageRank(v) {
    const S = cur(), P = prev();
    const ranked = rankedBlocks(S, P);
    const sub = plainSub(S, P);
    pageHead(v, 'Leading & Lagging', subLine(S, P));
    const go = r => { location.hash = '#/schools?block=' + encodeURIComponent(r.block || r.b); };
    const bcols = [C.rank, C.block, C.t, C.o, C.pend, C.p, C.dO];

    const g1 = h('<div class="grid-2"></div>');
    v.append(g1);
    tableCard(g1, 'Top 10 Leading Blocks (by onboarded %)', { cols: bcols, rows: ranked.slice(0, 10), onRow: go }, sub);
    tableCard(g1, 'Bottom 10 Lagging Blocks (by onboarded %)', { cols: bcols, rows: ranked.slice(-10).reverse(), onRow: go }, sub);

    const g2 = h('<div class="grid-2" style="margin-top:16px"></div>');
    v.append(g2);
    if (P) {
        tableCard(g2, `Top Gainers since ${fmtD(P.date)}`, { cols: [C.idx, C.block, C.dO, C.dP, C.p, { k: 'dFull', h: 'New 100% sch.', type: 'd' }], rows: [...ranked].sort((a, b) => b.dO - a.dO).slice(0, 10), onRow: go }, sub);
        tableCard(g2, `Least Progress since ${fmtD(P.date)}`, { cols: [C.idx, C.block, C.dO, C.dP, C.p, C.pend], rows: [...ranked].sort((a, b) => a.dO - b.dO || a.p - b.p).slice(0, 10), onRow: go }, sub);
    }
    tableCard(g2, 'Blocks with Most Pending Teachers', { cols: [C.idx, C.block, C.t, C.o, C.pend, C.p], rows: [...ranked].sort((a, b) => b.pend - a.pend).slice(0, 10), onRow: go }, sub);
    tableCard(g2, 'Blocks with Most 0% Schools', { cols: [C.idx, C.block, C.schools, C.zero, { k: 'zp', h: '% of schools at 0%', type: 'p', v: r => pct(r.zero, r.schools) }, C.p], rows: [...ranked].sort((a, b) => b.zero - a.zero).slice(0, 10), onRow: go }, sub);

    const sr = schoolRows(S, P);
    const scols = [C.idx, C.sb, C.school, C.udise, C.t, C.o, C.pend, C.p];
    const g3 = h('<div class="stack" style="margin-top:16px"></div>');
    v.append(g3);
    tableCard(g3, 'Priority Schools - Most Pending Teachers (Top 50)', { cols: scols, rows: [...sr].sort((a, b) => b.pend - a.pend).slice(0, 50) }, sub);
    tableCard(g3, 'Large Schools at 0% (Top 50 by teachers)', { cols: scols, rows: sr.filter(x => x.c === 'zero').sort((a, b) => b.t - a.t).slice(0, 50) }, sub);
    if (P) tableCard(g3, `Best Performing Schools since ${fmtD(P.date)} (Top 50)`, { cols: [...scols, C.dO], rows: sr.filter(x => x.dO > 0).sort((a, b) => b.dO - a.dO).slice(0, 50) }, sub);
}

function pageData(v) {
    pageHead(v, 'Upload / Manage Data', 'Upload the combined Excel file every day. One file is kept per date.');

    const up = card(v, 'Upload daily Excel');
    up.append(h(`<div class="upload">
        <label class="drop" id="drop">
            <input type="file" id="file" accept=".xlsx,.xls,.csv" multiple hidden>
            <span class="drop-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 16V4M7 9l5-5 5 5"/><path d="M4 16v3a1 1 0 0 0 1 1h14a1 1 0 0 0 1-1v-3"/></svg></span>
            <strong>Drag &amp; drop the Excel file(s) here</strong>
            <span class="help">or</span>
            <span class="btn primary">Browse files</span>
            <span class="help drop-note">.xlsx &middot; e.g. <code>Teacher_Onboarding_School_Wise_All_Blocks_2026-10-28.xlsx</code> &middot; several days can be uploaded together</span>
        </label>
        <div class="upload-opts">
            <div class="field">
                <label for="upDate">Data date <em>(optional)</em></label>
                <input type="date" id="upDate">
            </div>
            <div class="upload-hint">
                <span class="hl-icon">${icon('info')}</span>
                <p>The date is read automatically from the file name (<code>YYYY-MM-DD</code> or <code>DD-MM-YYYY</code>). Set a date here only when uploading a single file whose name has no date, or to override it.</p>
            </div>
        </div>
    </div>`));
    const drop = $('#drop', up), file = $('#file', up);
    file.onchange = () => { handleFiles([...file.files]); file.value = ''; };
    drop.ondragover = e => { e.preventDefault(); drop.classList.add('over'); };
    drop.ondragleave = () => drop.classList.remove('over');
    drop.ondrop = e => { e.preventDefault(); drop.classList.remove('over'); handleFiles([...e.dataTransfer.files]); };

    const rows = state.snaps.map(s => { const e = enrich(s); return { date: s.date, file: s.fileName || '', src: s.source, schools: e.total.schools, t: e.total.t, o: e.total.o, p: e.total.p, full: e.total.full, zero: e.total.zero }; }).reverse();
    const list = card(v, `Uploaded data (${rows.length} day${rows.length === 1 ? '' : 's'})`, null, true);
    list.append(h(`<div class="table-wrap"><table><thead><tr><th>Date</th><th>File</th><th class="num">Schools</th><th class="num">Teachers</th><th class="num">Onboarded</th><th class="num">%</th><th class="num">100%</th><th class="num">0%</th><th></th></tr></thead><tbody>
        ${rows.map(r => `<tr><td class="nowrap"><b>${fmtD(r.date)}</b></td><td>${esc(r.file)} ${r.src === 'repo' ? '<span class="badge repo">from repo</span>' : ''}</td><td class="num">${fmtN(r.schools)}</td><td class="num">${fmtN(r.t)}</td><td class="num">${fmtN(r.o)}</td><td class="num">${fmtP(r.p)}</td><td class="num">${fmtN(r.full)}</td><td class="num">${fmtN(r.zero)}</td>
        <td class="nowrap">${r.src === 'repo' ? '' : `<button class="btn sm danger" data-del="${r.date}">Delete</button>`}</td></tr>`).join('') || '<tr><td colspan="9" class="empty">No data uploaded yet</td></tr>'}
    </tbody></table></div>`));
    list.onclick = async e => {
        const b = e.target.closest('[data-del]');
        if (!b || !confirm(`Delete data of ${fmtD(b.dataset.del)}?`)) return;
        await DB.del(b.dataset.del);
        await reload();
        toast('Deleted');
    };

    const bk = card(v, 'Backup & restore');
    bk.append(h(`<div><p class="help">Uploaded data is saved only in this browser. Download a backup regularly, and restore it on another phone or computer.</p>
        <div class="actions"><button class="btn primary" id="bkDown">&#11015; Download backup (.json)</button>
        <label class="btn">&#11014; Restore backup<input type="file" id="bkUp" accept=".json" hidden></label>
        <button class="btn danger" id="bkClear">Delete all uploaded data</button></div></div>`));
    $('#bkDown', bk).onclick = async () => {
        const local = (await DB.all()).sort(byDate);
        const data = { app: 'prashast-monitor', version: 1, exported: new Date().toISOString(), snapshots: local.map(s => ({ date: s.date, fileName: s.fileName, uploadedAt: s.uploadedAt, schools: s.schools.map(({ b, n, u, t, o }) => ({ b, n, u, t, o })) })) };
        download(new Blob([JSON.stringify(data)], { type: 'application/json' }), `prashast-monitor-backup_${new Date().toISOString().slice(0, 10)}.json`);
    };
    $('#bkUp', bk).onchange = async e => {
        const f = e.target.files[0];
        if (!f) return;
        try {
            const data = JSON.parse(await f.text());
            if (!Array.isArray(data.snapshots)) throw new Error('Not a backup file of this portal');
            for (const s of data.snapshots) await DB.put({ ...s, source: 'upload' });
            await reload();
            toast(`Restored ${data.snapshots.length} day(s)`);
        } catch (err) { toast('Restore failed: ' + err.message, true); }
    };
    $('#bkClear', bk).onclick = async () => {
        if (!confirm('Delete ALL uploaded data from this browser? Download a backup first if needed.')) return;
        await DB.clear();
        await reload();
        toast('All uploaded data deleted');
    };

    const gh = card(v, 'Share data with everyone (GitHub)');
    gh.append(h(`<div class="help" style="font-size:13px">
        <p>Data uploaded here stays on this device. To let every officer who opens the portal see the same history, add the daily files to the repository:</p>
        <ol>
            <li>Put the Excel file in the <code>data/</code> folder of the repository (keep the date in the file name).</li>
            <li>Add the file name to <code>data/manifest.json</code>, for example:<br><code>["Teacher_Onboarding_School_Wise_All_Blocks_2026-10-27.xlsx", "Teacher_Onboarding_School_Wise_All_Blocks_2026-10-28.xlsx"]</code></li>
            <li>Commit. GitHub Pages updates in a minute or two; those dates then show as <span class="badge repo">from repo</span> for everyone.</li>
        </ol>
    </div>`));
}

function emptyState() {
    return h(`<div class="card empty-state"><h2>No data yet</h2>
        <p>Upload the combined school-wise Excel file (the one with the <b>Block Name</b> column) to start monitoring.</p>
        <p><a class="btn primary" href="#/data">&#11014; Upload Excel</a></p>
        <p class="help">${CREDIT}</p></div>`);
}

// ---------------------------------------------------------------- upload handling
async function handleFiles(files) {
    if (!files.length) return;
    const override = $('#upDate') && $('#upDate').value;
    if (override && files.length > 1) { toast('Pick a date only when uploading a single file.', true); return; }
    let lastDate = null;
    for (const f of files) {
        try {
            const date = override || dateFromName(f.name);
            if (!date) throw new Error('no date in the file name; choose the "Data date" and upload again');
            const schools = parseWorkbook(await f.arrayBuffer());
            const exists = state.snaps.find(s => s.date === date && s.source !== 'repo');
            if (exists && !confirm(`Data for ${fmtD(date)} already exists. Replace it with ${f.name}?`)) continue;
            await DB.put({ date, fileName: f.name, uploadedAt: new Date().toISOString(), source: 'upload', schools });
            lastDate = date;
            toast(`${fmtD(date)}: ${fmtN(schools.length)} schools loaded`);
        } catch (e) {
            console.error(e);
            toast(`${f.name}: ${e.message}`, true);
        }
    }
    if (lastDate) await reload(true);
}

// ---------------------------------------------------------------- data loading
async function loadRepo() {
    try {
        const res = await fetch('data/manifest.json', { cache: 'no-store' });
        if (!res.ok) return [];
        const list = await res.json();
        const out = [];
        for (const e of list) {
            const file = typeof e === 'string' ? e : e.file;
            const date = (typeof e === 'object' && e.date) || dateFromName(file);
            if (!file || !date) continue;
            try {
                const r = await fetch('data/' + encodeURIComponent(file), { cache: 'no-store' });
                if (!r.ok) throw new Error('HTTP ' + r.status);
                out.push({ date, fileName: file, source: 'repo', schools: parseWorkbook(await r.arrayBuffer()) });
            } catch (err) { toast(`data/${file}: ${err.message}`, true); }
        }
        return out;
    } catch (e) { return []; }
}

let repoSnaps = null;
async function reload(jumpToLatest) {
    if (repoSnaps === null) repoSnaps = await loadRepo();
    const local = await DB.all();
    const map = new Map(repoSnaps.map(s => [s.date, s]));
    local.forEach(s => map.set(s.date, s)); // uploaded data overrides repo data of the same date
    state.snaps = [...map.values()].sort(byDate);
    state.cache = {};
    const dates = state.snaps.map(s => s.date);
    if (jumpToLatest || !dates.includes(state.cur)) setCur(dates[dates.length - 1] || '');
    else if (state.prev && !dates.includes(state.prev)) setCur(state.cur);
    fillDateSelects();
    render();
}

function setCur(d) {
    state.cur = d;
    const i = state.snaps.findIndex(s => s.date === d);
    state.prev = i > 0 ? state.snaps[i - 1].date : '';
}

function fillDateSelects() {
    const cs = $('#curDate'), ps = $('#prevDate');
    const dates = state.snaps.map(s => s.date).reverse();
    cs.innerHTML = dates.length ? dates.map(d => `<option value="${d}">${fmtD(d)}</option>`).join('') : '<option value="">No data</option>';
    cs.value = state.cur;
    ps.innerHTML = '<option value="">None</option>' + dates.filter(d => d !== state.cur).map(d => `<option value="${d}">${fmtD(d)}</option>`).join('');
    ps.value = state.prev;
}

// ---------------------------------------------------------------- routing
const routes = { dashboard: pageDashboard, blocks: pageBlocks, schools: pageSchools, compare: pageCompare, trend: pageTrend, rank: pageRank, data: pageData };

function parseHash() {
    const raw = location.hash.replace(/^#\/?/, '');
    const [path, q = ''] = raw.split('?');
    const parts = path.split('/').filter(Boolean);
    return { name: routes[parts[0]] ? parts[0] : 'dashboard', arg: parts[1] || '', q: new URLSearchParams(q) };
}

function render() {
    state.charts.forEach(c => c.destroy());
    state.charts = [];
    const r = parseHash();
    const key = r.name + (r.arg ? '/' + r.arg : '');
    document.querySelectorAll('#sidebar a').forEach(a => a.classList.toggle('active', a.dataset.route === key));
    const v = $('#view');
    v.innerHTML = '';
    document.body.classList.remove('menu-open');
    try {
        if (!state.snaps.length && r.name !== 'data') v.append(emptyState());
        else routes[r.name](v, r);
    } catch (e) {
        console.error(e);
        v.append(h(`<div class="card"><div class="card-body">Something went wrong: ${esc(e.message)}</div></div>`));
    }
}

async function init() {
    if (typeof XLSX === 'undefined' || !window.jspdf || typeof Chart === 'undefined') {
        $('#view').innerHTML = '<div class="card"><div class="card-body">Could not load required libraries. Please check the internet connection and reload.</div></div>';
        return;
    }
    $('#year').textContent = new Date().getFullYear();
    $('#menuBtn').onclick =() => document.body.classList.toggle('menu-open');
    $('#backdrop').onclick = () => document.body.classList.remove('menu-open');
    $('#curDate').onchange = e => { setCur(e.target.value); fillDateSelects(); render(); };
    $('#prevDate').onchange = e => { state.prev = e.target.value; render(); };
    window.addEventListener('hashchange', () => { render(); window.scrollTo(0, 0); });
    await DB.open();
    await reload();
}

init();
