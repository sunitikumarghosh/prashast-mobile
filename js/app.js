/* PRASHAST Onboarding Monitor - SSM, Purba Bardhaman
 * Two modules: Teacher Onboarding (teachers per school) and School Onboarding
 * (Completed / Pending status per school).
 * Static app (GitHub Pages). Daily snapshots are stored in IndexedDB and can
 * also be published for everyone through data/manifest.json in the repo. */
'use strict';

const HEADER = 'Prashast On boarding status :: SSM, Purba Bardhaman';
const CREDIT = 'An in-house production of SSM, Purba Bardhaman';
const CAT = { full: '100%', part: '1-99%', zero: '0%', na: 'No teachers', done: 'Completed', pending: 'Pending' };
const NUM_TYPES = new Set(['idx', 'n', 'p', 'd', 'dp']);
const MODS = { teacher: 'Teacher Onboarding', school: 'School Onboarding' };

// state.snaps / cur / prev always refer to the active module (state.mod)
const state = {
    mod: 'teacher',
    mods: { teacher: { snaps: [], cur: '', prev: '' }, school: { snaps: [], cur: '', prev: '' } },
    charts: [], cache: {},
};
['snaps', 'cur', 'prev'].forEach(k => Object.defineProperty(state, k, {
    get() { return this.mods[this.mod][k]; },
    set(v) { this.mods[this.mod][k] = v; },
}));

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
// One object store per module, keyed by date; "meta" holds the uploaded school master (key 'master')
const STORES = { teacher: 'snapshots', school: 'school_snapshots', meta: 'meta' };
const DB = {
    db: null,
    mem: { teacher: new Map(), school: new Map(), meta: new Map() },
    open() {
        return new Promise(res => {
            setTimeout(res, 3000); // some browsers/private modes never answer; fall back to memory
            try {
                const r = indexedDB.open('prashast-monitor', 3);
                r.onupgradeneeded = () => Object.entries(STORES).forEach(([kind, n]) => {
                    if (!r.result.objectStoreNames.contains(n)) r.result.createObjectStore(n, { keyPath: kind === 'meta' ? 'key' : 'date' });
                });
                r.onsuccess = () => { this.db = r.result; res(); };
                r.onerror = () => res();
            } catch (e) { res(); }
        });
    },
    tx(kind, mode, fn) {
        return new Promise((res, rej) => {
            const t = this.db.transaction(STORES[kind], mode);
            const req = fn(t.objectStore(STORES[kind]));
            t.oncomplete = () => res(req && req.result);
            t.onerror = () => rej(t.error);
        });
    },
    all(kind) { return this.db ? this.tx(kind, 'readonly', s => s.getAll()) : Promise.resolve([...this.mem[kind].values()]); },
    put(kind, v) { if (!this.db) { this.mem[kind].set(v.date ?? v.key, v); return Promise.resolve(); } return this.tx(kind, 'readwrite', s => s.put(v)); },
    del(kind, k) { if (!this.db) { this.mem[kind].delete(k); return Promise.resolve(); } return this.tx(kind, 'readwrite', s => s.delete(k)); },
    clear(kind) { if (!this.db) { this.mem[kind].clear(); return Promise.resolve(); } return this.tx(kind, 'readwrite', s => s.clear()); },
};

// ---------------------------------------------------------------- Excel parsing
// Returns { kind: 'teacher' | 'school', schools }. The kind is detected from the columns:
// teacher files have Total Teachers / Teachers Onboarded, school files have a Status column.
const isDone = st => /^(completed?|onboarded|yes)$/i.test(String(st).trim());

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
        const iStatus = find(x => /status/i.test(x));
        const kind = iTotal >= 0 && iOnb >= 0 ? 'teacher' : iStatus >= 0 ? 'school' : null;
        if (!kind) throw new Error(`Sheet "${name}": neither "Total Teachers" / "Teachers Onboarded" nor "Status" columns found`);
        const schools = [];
        for (const r of rows.slice(hi + 1)) {
            let n = String(r[iSchool] ?? '').trim();
            if (!n) continue;
            let u = iUdise >= 0 ? String(r[iUdise] ?? '').trim() : '';
            const k = n.lastIndexOf('|');
            if (!u && k >= 0) { u = n.slice(k + 1).trim(); n = n.slice(0, k).trim(); }
            const b = iBlock >= 0 ? String(r[iBlock] ?? '').trim() : '';
            schools.push(kind === 'teacher'
                ? { b: b || 'UNKNOWN', n, u, t: num(r[iTotal]), o: num(r[iOnb]) }
                : { b: b || 'UNKNOWN', n, u, st: String(r[iStatus] ?? '').trim() || 'Pending' });
        }
        if (schools.length) return { kind, schools };
    }
    throw new Error('No sheet with a "School Name" column was found');
}

// School master (School Search export): UDISE Code, Management, Category, Type.
// Returns [{ u, mg, ct, ty }] or null when the file is not a master file.
function parseMaster(buf) {
    const wb = XLSX.read(buf, { type: 'array' });
    for (const name of wb.SheetNames) {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '', raw: true });
        const hi = rows.findIndex(r => r.some(c => /udise/i.test(String(c))) && r.some(c => /management/i.test(String(c))));
        if (hi < 0) continue;
        const H = rows[hi].map(c => String(c).trim());
        const iU = H.findIndex(x => /udise/i.test(x)), iM = H.findIndex(x => /management/i.test(x));
        const iC = H.findIndex(x => /category/i.test(x)), iT = H.findIndex(x => /^(school\s*)?type$/i.test(x));
        const txt = (r, i) => (i >= 0 ? String(r[i] ?? '').trim() : '');
        const out = rows.slice(hi + 1).map(r => ({ u: txt(r, iU), mg: txt(r, iM), ct: txt(r, iC), ty: txt(r, iT) })).filter(x => x.u);
        if (out.length) return out;
    }
    return null;
}

// ---------------------------------------------------------------- calculations
function category(x) {
    if (!x.t) return 'na';
    if (x.o >= x.t) return 'full';
    if (x.o <= 0) return 'zero';
    return 'part';
}
const blank = block => ({ block, schools: 0, t: 0, o: 0, full: 0, part: 0, zero: 0, na: 0, done: 0, pending: 0 });
function addTo(a, x) { a.schools++; a.t += x.t; a.o += x.o; a[x.c]++; }
function finish(a) { a.p = pct(a.o, a.t); a.pend = Math.max(0, a.t - a.o); return a; }

// In the school module every school counts as t = 1 and o = 1 when its status is Completed,
// so block / district totals, %, ranking and comparisons work the same way as for teachers.
function enrich(s) {
    if (!s || s._e) return s;
    const blocks = new Map();
    const T = blank('District Total');
    s.schools.forEach(x => {
        if (s.kind === 'school') { x.t = 1; x.o = isDone(x.st) ? 1 : 0; x.c = x.o ? 'done' : 'pending'; }
        else x.c = category(x);
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
    key = state.mod + '|' + key;
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
    rows.forEach(r => { T.schools += r.schools ?? 1; T.t += r.t; T.o += r.o; ['full', 'part', 'zero', 'na', 'done', 'pending'].forEach(k => { T[k] += r[k] || 0; }); });
    finish(T);
    T.dO = rows.some(r => r.dO != null) ? rows.reduce((s, r) => s + (r.dO || 0), 0) : null;
    return T;
}

// ---------------------------------------------------------------- school master groups (Management / Category / Type by UDISE)
const GROUPS = { mg: 'Management', ct: 'Category', ty: 'School Type' };
const NO_MASTER = 'Not in school master';
const groupOf = (x, by) => { const m = state.master && state.master.get(x.u); return (m && m[by]) || NO_MASTER; };
const shortCat = c => { const m = String(c || '').match(/\(([^)]+)\)\s*$/); return m ? m[1] : c || ''; };

// Rows shaped like block rows (group name in "block") so the block columns and totals can be reused
function groupRows(S, P, by, block) {
    return memo(`grp|${by}|${block}|${S.date}|${P ? P.date : ''}`, () => {
        const agg = X => {
            const map = new Map();
            X.schools.forEach(x => {
                if (block && x.b !== block) return;
                const g = groupOf(x, by);
                if (!map.has(g)) map.set(g, blank(g));
                addTo(map.get(g), x);
            });
            map.forEach(finish);
            return map;
        };
        const now = agg(S), before = P ? agg(P) : null;
        return [...now.values()].map(g => {
            const pg = before && before.get(g.block);
            return { ...g, dO: pg ? g.o - pg.o : null, dP: pg ? g.p - pg.p : null };
        }).sort((a, b) => b.schools - a.schools);
    });
}

// School list filters by group (?mg=..&ct=..&ty=..)
const groupParams = r => Object.keys(GROUPS).filter(k => r.q.get(k)).map(k => [k, r.q.get(k)]);
const inGroups = (x, gf) => gf.every(([k, val]) => groupOf(x, k) === val);
const groupTitle = gf => gf.map(([, val]) => ` - ${val}`).join('');
function listHash(base, cat, block, gf) {
    const q = new URLSearchParams();
    if (block) q.set('block', block);
    gf.forEach(([k, val]) => { if (val) q.set(k, val); });
    const s = q.toString();
    return `${base}schools${cat && cat !== 'all' ? '/' + cat : ''}${s ? '?' + s : ''}`;
}
function groupSelects(S, gf) {
    if (!state.master) return '';
    return Object.keys(GROUPS).map(k => {
        const sel = (gf.find(g => g[0] === k) || [])[1] || '';
        const vals = [...new Set(S.schools.map(x => groupOf(x, k)))].sort();
        return `<label>${GROUPS[k]}<select data-grp="${k}"><option value="">All</option>${vals.map(x => `<option${x === sel ? ' selected' : ''}>${esc(x)}</option>`).join('')}</select></label>`;
    }).join('');
}
function wireGroupSelects(f, base, cat, block) {
    f.querySelectorAll('[data-grp]').forEach(s => {
        s.onchange = () => {
            const gf = [...f.querySelectorAll('[data-grp]')].map(x => [x.dataset.grp, x.value]).filter(g => g[1]);
            location.hash = listHash(base, cat, block, gf);
        };
    });
}
const groupCols = () => (state.master ? [
    { k: 'mg', h: 'Management', cls: 'wrap', v: x => groupOf(x, 'mg') },
    { k: 'ct', h: 'Category', cls: 'nowrap', v: x => shortCat(groupOf(x, 'ct')) },
] : []);

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
    groupSummary(v, S, P);
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
    const gf = groupParams(r);
    const filtered = () => all.filter(x => (!block || x.b === block) && (catKey === 'all' || x.c === catKey) && inGroups(x, gf) &&
        (!search || (x.n + ' ' + x.u).toLowerCase().includes(search)));
    const title = SCHOOL_TITLES[catKey] + (block ? ` - ${block} Block` : '') + groupTitle(gf);
    const go = (c, b) => { location.hash = listHash('#/', c, b, gf); };

    const cols = [C.idx];
    if (!block) cols.push(C.sb);
    cols.push(C.school, C.udise, ...groupCols(), C.t, C.o, col(C.pend, { invert: true }), C.p);
    if (P) cols.push({ k: 'oPrev', h: `Onb. on ${fmtD(P.date)}`, type: 'n' }, C.dO);
    if (catKey === 'all') cols.push(C.cat);

    const defaultSort = { zero: ['t', 'desc'], part: ['p', 'asc'] }[catKey] || [null];
    const tbl = makeTable({ cols, rows: filtered(), sort: defaultSort[0], dir: defaultSort[1], pageSize: 50 });
    pageHead(v, title, subLine(S, P), exportBtns(tableSpec(tbl, () => title + (search ? ` (search: ${search})` : ''), plainSub(S, P))));

    const f = h(`<div class="filters">
        <label>Block<select id="fBlock"><option value="">All blocks</option>${blockNames().map(b => `<option${b === block ? ' selected' : ''}>${esc(b)}</option>`).join('')}</select></label>
        <label>Onboarding<select id="fCat">${Object.keys(SCHOOL_TITLES).map(k => `<option value="${k}"${k === catKey ? ' selected' : ''}>${k === 'all' ? 'All levels' : CAT[k] + ' onboarded'}</option>`).join('')}</select></label>
        ${groupSelects(S, gf)}
        <label>Search school / UDISE<input type="search" id="fSearch" placeholder="Type to search..."></label>
    </div>`);
    v.append(f);
    $('#fBlock', f).onchange = e => go(catKey, e.target.value);
    $('#fCat', f).onchange = e => go(e.target.value, block);
    wireGroupSelects(f, '#/', catKey, block);

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
    pageHead(v, 'Upload / Manage Data', 'Upload the combined Excel files every day. One file is kept per date for each module; the module is detected from the file.');

    const up = card(v, 'Upload daily Excel');
    up.append(h(`<div class="upload">
        <label class="drop" id="drop">
            <input type="file" id="file" accept=".xlsx,.xls,.csv" multiple hidden>
            <span class="drop-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 16V4M7 9l5-5 5 5"/><path d="M4 16v3a1 1 0 0 0 1 1h14a1 1 0 0 0 1-1v-3"/></svg></span>
            <strong>Drag &amp; drop the Excel file(s) here</strong>
            <span class="help">or</span>
            <span class="btn primary">Browse files</span>
            <span class="help drop-note">.xlsx &middot; <code>Teacher_Onboarding_School_Wise_All_Blocks_2026-10-28.xlsx</code> or <code>Schools_Onboarding_Status_All_Blocks_2026-10-29.xlsx</code>, and the School Search <code>.csv</code> (school master) &middot; several files can be uploaded together</span>
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

    const delBtn = (kind, r) => (r.src === 'repo' ? '' : `<button class="btn sm danger" data-kind="${kind}" data-del="${r.date}">Delete</button>`);
    const fileCell = r => `${esc(r.file)} ${r.src === 'repo' ? '<span class="badge repo">from repo</span>' : ''}`;
    const listed = kind => state.mods[kind].snaps.map(s => { const e = enrich(s); return { date: s.date, file: s.fileName || '', src: s.source, ...e.total }; }).reverse();

    const tRows = listed('teacher');
    const tList = card(v, `Teacher Onboarding data (${tRows.length} day${tRows.length === 1 ? '' : 's'})`, null, true);
    tList.append(h(`<div class="table-wrap"><table><thead><tr><th>Date</th><th>File</th><th class="num">Schools</th><th class="num">Teachers</th><th class="num">Onboarded</th><th class="num">%</th><th class="num">100%</th><th class="num">0%</th><th></th></tr></thead><tbody>
        ${tRows.map(r => `<tr><td class="nowrap"><b>${fmtD(r.date)}</b></td><td>${fileCell(r)}</td><td class="num">${fmtN(r.schools)}</td><td class="num">${fmtN(r.t)}</td><td class="num">${fmtN(r.o)}</td><td class="num">${fmtP(r.p)}</td><td class="num">${fmtN(r.full)}</td><td class="num">${fmtN(r.zero)}</td>
        <td class="nowrap">${delBtn('teacher', r)}</td></tr>`).join('') || '<tr><td colspan="9" class="empty">No data uploaded yet</td></tr>'}
    </tbody></table></div>`));

    const sRows = listed('school');
    const sList = card(v, `School Onboarding data (${sRows.length} day${sRows.length === 1 ? '' : 's'})`, null, true);
    sList.append(h(`<div class="table-wrap"><table><thead><tr><th>Date</th><th>File</th><th class="num">Schools</th><th class="num">Completed</th><th class="num">Pending</th><th class="num">%</th><th></th></tr></thead><tbody>
        ${sRows.map(r => `<tr><td class="nowrap"><b>${fmtD(r.date)}</b></td><td>${fileCell(r)}</td><td class="num">${fmtN(r.schools)}</td><td class="num">${fmtN(r.o)}</td><td class="num">${fmtN(r.pend)}</td><td class="num">${fmtP(r.p)}</td>
        <td class="nowrap">${delBtn('school', r)}</td></tr>`).join('') || '<tr><td colspan="7" class="empty">No data uploaded yet</td></tr>'}
    </tbody></table></div>`));

    const onDel = async e => {
        const b = e.target.closest('[data-del]');
        if (!b || !confirm(`Delete ${MODS[b.dataset.kind]} data of ${fmtD(b.dataset.del)}?`)) return;
        await DB.del(b.dataset.kind, b.dataset.del);
        await reload();
        toast('Deleted');
    };
    tList.onclick = onDel;
    sList.onclick = onDel;

    const mi = state.masterInfo;
    const mBox = card(v, 'School master (Management / Category / Type)');
    const unmatched = kind => { const s = state.mods[kind].snaps; const last = s[s.length - 1]; return last ? last.schools.filter(x => !state.master || !state.master.has(x.u)).length : null; };
    const um = ['teacher', 'school'].map(k => [k, unmatched(k)]).filter(([, n]) => n != null);
    mBox.append(h(`<div class="help" style="font-size:13px">
        <p>The <b>Management</b>, <b>Category</b> and <b>Type</b> of each school come from the School Search export
        (columns UDISE Code, Management, Category, Type), matched by UDISE code. Upload a newer CSV here at any time; it replaces the one in use in this browser.</p>
        ${mi ? `<p>In use: <b>${esc(mi.fileName)}</b> ${mi.source === 'repo' ? '<span class="badge repo">from repo</span>' : `(uploaded ${new Date(mi.uploadedAt).toLocaleString('en-IN')})`} &middot; ${fmtN(mi.count)} schools</p>
            <p>${um.map(([k, n]) => `${MODS[k]} (latest day): ${n ? `<span class="down">${fmtN(n)} school(s) not found in the master</span>` : 'all schools matched'}`).join(' &middot; ')}</p>
            ${mi.source === 'repo' ? '' : '<p><button class="btn sm danger" id="mDel">Remove uploaded master</button></p>'}`
        : '<p class="down">No school master loaded. Upload the School Search CSV to enable the Management / Category reports.</p>'}
    </div>`));
    if ($('#mDel', mBox)) $('#mDel', mBox).onclick = async () => {
        if (!confirm('Remove the uploaded school master from this browser?')) return;
        await DB.del('meta', 'master');
        await reload();
        toast('School master removed');
    };

    const bk = card(v, 'Backup & restore');
    bk.append(h(`<div><p class="help">Uploaded data is saved only in this browser. Download a backup regularly, and restore it on another phone or computer.</p>
        <div class="actions"><button class="btn primary" id="bkDown">&#11015; Download backup (.json)</button>
        <label class="btn">&#11014; Restore backup<input type="file" id="bkUp" accept=".json" hidden></label>
        <button class="btn danger" id="bkClear">Delete all uploaded data</button></div></div>`));
    // Backup v2: "snapshots" = teacher data (same as v1), "schoolSnapshots" = school data
    $('#bkDown', bk).onclick = async () => {
        const meta = s => ({ date: s.date, fileName: s.fileName, uploadedAt: s.uploadedAt });
        const teacher = (await DB.all('teacher')).sort(byDate);
        const school = (await DB.all('school')).sort(byDate);
        const data = {
            app: 'prashast-monitor', version: 2, exported: new Date().toISOString(),
            snapshots: teacher.map(s => ({ ...meta(s), schools: s.schools.map(({ b, n, u, t, o }) => ({ b, n, u, t, o })) })),
            schoolSnapshots: school.map(s => ({ ...meta(s), schools: s.schools.map(({ b, n, u, st }) => ({ b, n, u, st })) })),
            master: (await DB.all('meta')).find(x => x.key === 'master') || null,
        };
        download(new Blob([JSON.stringify(data)], { type: 'application/json' }), `prashast-monitor-backup_${new Date().toISOString().slice(0, 10)}.json`);
    };
    $('#bkUp', bk).onchange = async e => {
        const f = e.target.files[0];
        if (!f) return;
        try {
            const data = JSON.parse(await f.text());
            if (!Array.isArray(data.snapshots)) throw new Error('Not a backup file of this portal');
            const school = Array.isArray(data.schoolSnapshots) ? data.schoolSnapshots : [];
            for (const s of data.snapshots) await DB.put('teacher', { ...s, source: 'upload' });
            for (const s of school) await DB.put('school', { ...s, source: 'upload' });
            if (data.master) await DB.put('meta', data.master);
            await reload();
            toast(`Restored ${data.snapshots.length} teacher day(s) and ${school.length} school day(s)`);
        } catch (err) { toast('Restore failed: ' + err.message, true); }
    };
    $('#bkClear', bk).onclick = async () => {
        if (!confirm('Delete ALL uploaded data (teacher and school onboarding) from this browser? Download a backup first if needed.')) return;
        await DB.clear('teacher');
        await DB.clear('school');
        await DB.clear('meta');
        await reload();
        toast('All uploaded data deleted');
    };

    const gh = card(v, 'Share data with everyone (GitHub)');
    gh.append(h(`<div class="help" style="font-size:13px">
        <p>Data uploaded here stays on this device. To let every officer who opens the portal see the same history, add the daily files to the repository:</p>
        <ol>
            <li>Put the Excel file in the <code>data/</code> folder of the repository (keep the date in the file name).</li>
            <li>Add the file name to <code>data/manifest.json</code>. Teacher and school files go in the same list; the module is detected from the columns, for example:<br><code>["Teacher_Onboarding_School_Wise_All_Blocks_2026-10-28.xlsx", "Schools_Onboarding_Status_All_Blocks_2026-10-29.xlsx"]</code></li>
            <li>Commit. GitHub Pages updates in a minute or two; those dates then show as <span class="badge repo">from repo</span> for everyone.</li>
        </ol>
    </div>`));
}

// ================================================================ School Onboarding module
// Each school: t = 1, o = 1 if Completed. Block rows: schools, o (onboarded), pend, p.
const SC = {
    schools: col(C.schools, { h: 'Total Schools' }),
    o: col(C.o, { h: 'Schools Onboarded' }),
    pend: col(C.pend, { h: 'Pending Schools', invert: true }),
    p: C.p,
    dO: col(C.dO, { h: 'Change (Onb.)' }),
    st: col(C.cat, { h: 'Status' }),
};
const ST = { all: 'All Schools', done: 'Onboarded Schools (Completed)', pending: 'Pending Schools' };
const sTitle = t => `School Onboarding - ${t}`;
const sGo = r => { location.hash = '#/s/schools?block=' + encodeURIComponent(r.block || r.b); };

function sDashboard(v) {
    const S = cur(), P = prev(), T = S.total, PT = P && P.total;
    const d = k => (PT ? T[k] - PT[k] : null);
    const ranked = rankedBlocks(S, P);
    const count = s => ranked.filter(b => b.status === s).length;
    pageHead(v, 'School Onboarding Dashboard', subLine(S, P));

    const chg = PT ? `<span class="hero-chip ${T.p >= PT.p ? 'pos' : 'neg'}">${icon(T.p >= PT.p ? 'up' : 'down')} ${signed(T.p - PT.p, true)} pt &middot; ${signed(T.o - PT.o)} schools since ${fmtD(P.date)}</span>` : '';
    v.append(h(`<section class="hero">
        <div class="ring" style="--p:${Math.min(100, T.p)}"><div class="ring-in"><b>${fmtP(T.p)}%</b><small>onboarded</small></div></div>
        <div class="hero-main">
            <div class="hero-label">District school onboarding &middot; Purba Bardhaman</div>
            <div class="hero-title"><b>${fmtN(T.o)}</b> of ${fmtN(T.schools)} schools onboarded</div>
            <div class="hero-bar"><span style="width:${Math.min(100, T.p)}%"></span></div>
            <div class="hero-meta">${fmtN(T.pend)} schools still pending across ${S.blocks.length} blocks</div>
            ${chg}
        </div>
        <div class="hero-stats">
            <div><b>${S.blocks.length}</b><span>Blocks</span></div>
            <div><b>${fmtN(T.schools)}</b><span>Schools</span></div>
            <div><b class="g">${count('Leading')}</b><span>Leading blocks</span></div>
            <div><b class="r">${count('Lagging')}</b><span>Lagging blocks</span></div>
        </div>
    </section>`));

    v.append(h(`<div class="stat-grid">
        ${statCard({ label: 'Total Schools', value: fmtN(T.schools), delta: d('schools'), foot: `in ${S.blocks.length} blocks`, ic: 'school', tone: 'blue', href: '#/s/schools' })}
        ${statCard({ label: 'Schools Onboarded', value: fmtN(T.o), delta: d('o'), foot: `status "Completed"`, ic: 'check', tone: 'green', href: '#/s/schools/done' })}
        ${statCard({ label: 'Pending Schools', value: fmtN(T.pend), delta: d('pend'), invert: true, foot: `${fmtP(100 - T.p)}% yet to onboard`, ic: 'clock', tone: 'red', href: '#/s/schools/pending' })}
        ${statCard({ label: 'Onboarded %', value: fmtP(T.p) + '%', delta: d('p'), pctDelta: true, foot: `best block ${fmtP(ranked[0].p)}%`, ic: 'up', tone: 'indigo' })}
    </div>`));

    const best = ranked[0], worst = ranked[ranked.length - 1];
    const none = ranked.filter(b => b.o === 0).map(b => b.block);
    const mostPend = [...ranked].sort((a, b) => b.pend - a.pend)[0];
    const notes = [
        ['up', `Highest school onboarding: <b>${esc(best.block)}</b> at ${fmtP(best.p)}%; lowest: <b>${esc(worst.block)}</b> at ${fmtP(worst.p)}%.`],
        ['flag', `Most pending schools: <b>${esc(mostPend.block)}</b> (${fmtN(mostPend.pend)} of ${fmtN(mostPend.schools)} pending). <a href="#/s/schools/pending?block=${encodeURIComponent(mostPend.block)}">See list &rsaquo;</a>`],
        ['zero', none.length ? `<b>${none.length}</b> block(s) have not onboarded any school: ${none.map(esc).join(', ')}.` : 'Every block has onboarded at least one school.'],
    ];
    if (P) {
        const sr = schoolRows(S, P);
        const newly = sr.filter(x => x.dO > 0).length;
        const reverted = sr.filter(x => x.dO < 0).length;
        const still = ranked.filter(b => b.dO === 0).map(b => b.block);
        const gainer = [...ranked].sort((a, b) => (b.dO || 0) - (a.dO || 0))[0];
        notes.unshift(['check', `Since ${fmtD(P.date)}: <b>${fmtN(newly)}</b> schools newly onboarded. <a href="#/s/compare">Details &rsaquo;</a>`]);
        if (gainer && gainer.dO > 0) notes.push(['up', `Top gainer since ${fmtD(P.date)}: <b>${esc(gainer.block)}</b> (${signed(gainer.dO)} schools).`]);
        if (still.length) notes.push(['info', `No progress since ${fmtD(P.date)} in ${still.length} block(s): ${still.map(esc).join(', ')}.`]);
        if (reverted) notes.push(['down', `<span class="down">${fmtN(reverted)} school(s) were Completed on ${fmtD(P.date)} but are Pending now</span> &mdash; please verify.`]);
    }
    card(v, 'Key Highlights').append(h(`<ul class="highlights">${notes.map(([i, t]) => `<li><span class="hl-icon">${icon(i)}</span><span>${t}</span></li>`).join('')}</ul>`));

    const g = h('<div class="grid-2"></div>');
    v.append(g);
    chart(card(g, 'Schools by Status'), 260, {
        type: 'doughnut',
        data: { labels: ['Completed', 'Pending'], datasets: [{ data: [T.o, T.pend], backgroundColor: [COLORS.good, COLORS.bad], borderWidth: 2, borderColor: '#fff' }] },
        options: { plugins: { legend: { position: window.innerWidth < 600 ? 'bottom' : 'right', labels: { usePointStyle: true, padding: 14 } } }, cutout: '64%' },
    });
    const bands = [['Below 10%', 0, 10], ['10-25%', 10, 25], ['25-50%', 25, 50], ['50% and above', 50, 101]];
    chart(card(g, 'Blocks by School Onboarding Level'), 260, {
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
                    { label: 'Schools onboarded', data: all.map(s => s.total.o), borderColor: COLORS.good, tension: .3, yAxisID: 'y1', pointRadius: 3 },
                    { label: 'Pending schools', data: all.map(s => s.total.pend), borderColor: COLORS.bad, tension: .3, yAxisID: 'y1', pointRadius: 3 },
                ],
            },
            options: { interaction: { mode: 'index', intersect: false }, plugins: { legend: { labels: { usePointStyle: true } } }, scales: { y: { ticks: { callback: x => x + '%' } }, y1: { position: 'right', grid: { drawOnChartArea: false }, title: { display: true, text: 'Schools' } } } },
        });
    }

    const g2 = h('<div class="grid-2" style="margin-top:16px"></div>');
    v.append(g2);
    const mini = [C.rank, C.block, SC.schools, SC.o, C.p, col(SC.dO, { h: 'Change' })];
    tableCard(g2, sTitle('Top 5 Leading Blocks'), { cols: mini, rows: ranked.slice(0, 5), onRow: sGo }, plainSub(S, P));
    tableCard(g2, sTitle('Bottom 5 Lagging Blocks'), { cols: mini, rows: ranked.slice(-5).reverse(), onRow: sGo }, plainSub(S, P));
    groupSummary(v, S, P);
}

function sBlocks(v) {
    const S = cur(), P = prev();
    const rows = rankedBlocks(S, P);
    const cols = [C.rank, C.block, SC.schools, SC.o, SC.pend, C.p, SC.dO, C.dP, C.status];
    const foot = Object.assign(totalsRow(rows), { dP: P ? S.total.p - P.total.p : null });
    const tbl = makeTable({ cols, rows, foot, onRow: sGo });
    const title = sTitle('Block-wise Status');
    pageHead(v, 'Block-wise School Onboarding', subLine(S, P) + ' &middot; tap a block to see its schools', exportBtns(tableSpec(tbl, title, plainSub(S, P))));
    v.append(h(`<p class="help">Status: <span class="badge Leading">Leading</span> = top quarter of blocks by % of schools onboarded, <span class="badge Lagging">Lagging</span> = bottom quarter, <span class="badge Moderate">Moderate</span> = the rest.</p>`));
    card(v, 'All Blocks', null, true).append(tbl.el);

    const byName = [...rows].sort((a, b) => a.block.localeCompare(b.block));
    chart(card(v, 'Onboarded and pending schools by block'), Math.max(320, rows.length * 20 + 60), {
        type: 'bar',
        data: {
            labels: byName.map(b => b.block),
            datasets: [
                { label: 'Completed', data: byName.map(b => b.o), backgroundColor: COLORS.good },
                { label: 'Pending', data: byName.map(b => b.pend), backgroundColor: COLORS.bad },
            ],
        },
        options: { indexAxis: 'y', scales: { x: { stacked: true }, y: { stacked: true, ticks: { autoSkip: false, font: { size: 10 } } } }, interaction: { mode: 'index', intersect: false } },
    });
}

function sSchools(v, r) {
    const S = cur(), P = prev();
    const catKey = ST[r.arg] ? r.arg : 'all';
    const block = r.q.get('block') || '';
    let search = '';
    const all = schoolRows(S, P);
    const gf = groupParams(r);
    const filtered = () => all.filter(x => (!block || x.b === block) && (catKey === 'all' || x.c === catKey) && inGroups(x, gf) &&
        (!search || (x.n + ' ' + x.u).toLowerCase().includes(search)));
    const heading = ST[catKey] + (block ? ` - ${block} Block` : '') + groupTitle(gf);
    const title = sTitle(heading);
    const go = (c, b) => { location.hash = listHash('#/s/', c, b, gf); };

    const cols = [C.idx];
    if (!block) cols.push(C.sb);
    cols.push(C.school, C.udise, ...groupCols(), SC.st);
    if (P) cols.push({ k: 'cPrev', h: `Status on ${fmtD(P.date)}`, type: 'cat' },
        { k: 'move', h: 'Change', v: x => (x.cPrev == null ? 'New' : x.dO > 0 ? 'Newly onboarded' : x.dO < 0 ? 'Reverted' : '') });

    const tbl = makeTable({ cols, rows: filtered(), pageSize: 50 });
    pageHead(v, heading, subLine(S, P), exportBtns(tableSpec(tbl, () => title + (search ? ` (search: ${search})` : ''), plainSub(S, P))));

    const f = h(`<div class="filters">
        <label>Block<select id="fBlock"><option value="">All blocks</option>${blockNames().map(b => `<option${b === block ? ' selected' : ''}>${esc(b)}</option>`).join('')}</select></label>
        <label>Status<select id="fCat">${Object.keys(ST).map(k => `<option value="${k}"${k === catKey ? ' selected' : ''}>${k === 'all' ? 'All statuses' : CAT[k]}</option>`).join('')}</select></label>
        ${groupSelects(S, gf)}
        <label>Search school / UDISE<input type="search" id="fSearch" placeholder="Type to search..."></label>
    </div>`);
    v.append(f);
    $('#fBlock', f).onchange = e => go(catKey, e.target.value);
    $('#fCat', f).onchange = e => go(e.target.value, block);
    wireGroupSelects(f, '#/s/', catKey, block);

    const chips = h('<div class="chips"></div>');
    v.append(chips);
    const drawChips = rows => {
        const t = totalsRow(rows);
        chips.innerHTML = `<span class="chip">Schools <b>${fmtN(rows.length)}</b></span><span class="chip">Completed <b>${fmtN(t.o)}</b></span>
            <span class="chip">Pending <b>${fmtN(t.pend)}</b></span><span class="chip">Onboarded % <b>${fmtP(t.p)}</b></span>`;
    };
    drawChips(tbl.rows());
    let timer;
    $('#fSearch', f).oninput = e => {
        clearTimeout(timer);
        timer = setTimeout(() => { search = e.target.value.trim().toLowerCase(); const rows = filtered(); tbl.setRows(rows); drawChips(rows); }, 200);
    };
    card(v, 'Schools', null, true).append(tbl.el);
}

function sRank(v) {
    const S = cur(), P = prev();
    const ranked = rankedBlocks(S, P);
    const sub = plainSub(S, P);
    pageHead(v, 'Leading & Lagging (School Onboarding)', subLine(S, P));
    const bcols = [C.rank, C.block, SC.schools, SC.o, SC.pend, C.p, SC.dO];

    const g1 = h('<div class="grid-2"></div>');
    v.append(g1);
    tableCard(g1, sTitle('Top 10 Leading Blocks'), { cols: bcols, rows: ranked.slice(0, 10), onRow: sGo }, sub);
    tableCard(g1, sTitle('Bottom 10 Lagging Blocks'), { cols: bcols, rows: ranked.slice(-10).reverse(), onRow: sGo }, sub);

    const g2 = h('<div class="grid-2" style="margin-top:16px"></div>');
    v.append(g2);
    if (P) {
        tableCard(g2, sTitle(`Top Gainers since ${fmtD(P.date)}`), { cols: [C.idx, C.block, SC.dO, C.dP, C.p], rows: [...ranked].sort((a, b) => b.dO - a.dO).slice(0, 10), onRow: sGo }, sub);
        tableCard(g2, sTitle(`Least Progress since ${fmtD(P.date)}`), { cols: [C.idx, C.block, SC.dO, C.dP, C.p, SC.pend], rows: [...ranked].sort((a, b) => a.dO - b.dO || a.p - b.p).slice(0, 10), onRow: sGo }, sub);
    }
    tableCard(g2, sTitle('Blocks with Most Pending Schools'), { cols: [C.idx, C.block, SC.schools, SC.o, SC.pend, C.p], rows: [...ranked].sort((a, b) => b.pend - a.pend).slice(0, 10), onRow: sGo }, sub);
}

function sCompare(v, r) {
    if (state.snaps.length < 2) {
        pageHead(v, 'Day-wise Comparison (School Onboarding)');
        v.append(h('<div class="card"><div class="card-body">Upload school onboarding data for at least two dates to compare.</div></div>'));
        return;
    }
    const dates = state.snaps.map(s => s.date);
    const bDate = r.q.get('b') || state.cur;
    const aDate = r.q.get('a') || state.prev || dates[Math.max(0, dates.indexOf(bDate) - 1)];
    const A = snap(aDate), B = snap(bDate);
    const opts = sel => dates.slice().reverse().map(d => `<option value="${d}"${d === sel ? ' selected' : ''}>${fmtD(d)}</option>`).join('');
    const sub = `${fmtD(A.date)} vs ${fmtD(B.date)}`;
    pageHead(v, 'Day-wise Comparison (School Onboarding)', `Changes from <b>${fmtD(A.date)}</b> to <b>${fmtD(B.date)}</b>`);

    const f = h(`<div class="filters"><label>From date<select id="cA">${opts(aDate)}</select></label><label>To date<select id="cB">${opts(bDate)}</select></label></div>`);
    v.append(f);
    const go = () => { location.hash = `#/s/compare?a=${$('#cA').value}&b=${$('#cB').value}`; };
    $('#cA', f).onchange = go;
    $('#cB', f).onchange = go;

    const TA = A.total, TB = B.total;
    const k = (label, key, o = {}) => `<div class="kpi ${o.cls || ''}"><span class="kpi-label">${label}</span><span class="kpi-val">${o.pct ? fmtP(TB[key]) + '%' : fmtN(TB[key])}</span><span class="kpi-delta">${deltaHtml(TB[key] - TA[key], o.pct, o.invert)} <small>from ${o.pct ? fmtP(TA[key]) + '%' : fmtN(TA[key])}</small></span></div>`;
    v.append(h(`<div class="kpis">${k('Schools Onboarded', 'o', { cls: 'good' })}${k('Onboarded %', 'p', { pct: true })}${k('Pending Schools', 'pend', { invert: true, cls: 'bad' })}${k('Total Schools', 'schools')}</div>`));

    const names = [...new Set([...A.blocks, ...B.blocks].map(b => b.block))];
    const rows = names.map(n => {
        const a = A.byBlock.get(n) || finish(blank(n)), b = B.byBlock.get(n) || finish(blank(n));
        return { block: n, schools: b.schools, oA: a.o, o: b.o, dO: b.o - a.o, pA: a.p, p: b.p, dP: b.p - a.p, pend: b.pend };
    });
    const tot = { block: 'District Total', schools: TB.schools, oA: TA.o, o: TB.o, dO: TB.o - TA.o, pA: TA.p, p: TB.p, dP: TB.p - TA.p, pend: TB.pend };
    const dA = fmtD(A.date), dB = fmtD(B.date);
    tableCard(v, sTitle('Block-wise Comparison'), {
        cols: [C.idx, C.block, SC.schools, { k: 'oA', h: `Onb. ${dA}`, type: 'n' }, { k: 'o', h: `Onb. ${dB}`, type: 'n' }, SC.dO,
            { k: 'pA', h: `% ${dA}`, type: 'p' }, { k: 'p', h: `% ${dB}`, type: 'p' }, C.dP, SC.pend],
        rows, sort: 'dO', foot: tot,
    }, sub);

    const sr = schoolRows(B, A);
    const groups = [
        ['Newly onboarded', x => x.dO > 0],
        ['Still pending on both dates', x => x.c === 'pending' && x.cPrev === 'pending'],
        ['Completed on both dates', x => x.c === 'done' && x.cPrev === 'done'],
        ['Reverted to pending', x => x.dO < 0],
        ['New in later file', x => x.cPrev == null],
    ].map(([label, fn]) => ({ label, rows: sr.filter(fn) }));
    let active = groups[0];
    const cols = [C.idx, C.sb, C.school, C.udise, { k: 'cPrev', h: `Status ${dA}`, type: 'cat' }, { k: 'c', h: `Status ${dB}`, type: 'cat' }];
    const tbl = makeTable({ cols, rows: active.rows, pageSize: 50 });
    const body = card(v, 'School-level Movement', exportBtns(() => tableSpec(tbl, sTitle(`School Movement - ${active.label}`), sub)()), true);
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

function sTrend(v) {
    const all = state.snaps.map(s => enrich(s));
    const rows = all.map((s, i) => {
        const p = i ? all[i - 1].total : null, t = s.total;
        return { date: s.date, schools: t.schools, o: t.o, dO: p ? t.o - p.o : null, p: t.p, dP: p ? t.p - p.p : null, pend: t.pend };
    }).reverse();
    pageHead(v, 'Daily Trend (School Onboarding)', `${all.length} day(s) of data, from <b>${fmtD(all[0].date)}</b> to <b>${fmtD(all[all.length - 1].date)}</b>`);

    if (all.length > 1) {
        chart(card(v, 'District school onboarding by date'), 280, {
            type: 'bar',
            data: {
                labels: all.map(s => fmtD(s.date)),
                datasets: [
                    { type: 'line', label: 'Onboarded %', data: all.map(s => round2(s.total.p)), borderColor: COLORS.brand, backgroundColor: COLORS.brand, yAxisID: 'y1', tension: .25 },
                    { label: 'Schools onboarded', data: all.map(s => s.total.o), backgroundColor: 'rgba(46,125,50,.4)', yAxisID: 'y' },
                ],
            },
            options: { interaction: { mode: 'index', intersect: false }, scales: { y: { beginAtZero: true }, y1: { position: 'right', grid: { drawOnChartArea: false }, ticks: { callback: x => x + '%' } } } },
        });
    }

    tableCard(v, sTitle('District Date-wise Summary'), {
        cols: [{ k: 'date', h: 'Date', cls: 'nowrap', show: fmtD }, SC.schools, SC.o, col(SC.dO, { h: 'Daily change' }), C.p, C.dP, SC.pend],
        rows,
    });

    const metricBox = card(v, 'Block-wise Trend (by date)');
    const sel = h('<div class="filters"><label>Show<select><option value="p">Onboarded %</option><option value="o">Schools onboarded</option><option value="pend">Pending schools</option></select></label></div>');
    metricBox.append(sel);
    const holder = h('<div></div>');
    metricBox.append(holder);
    const draw = metric => {
        holder.innerHTML = '';
        const last = all[all.length - 1], first = all[0];
        const type = metric === 'p' ? 'p' : 'n';
        const cols = [C.idx, C.block, ...all.map(s => ({ k: s.date, h: fmtD(s.date), type })),
            { k: 'chg', h: 'Change (first to last)', type: metric === 'p' ? 'dp' : 'd', invert: metric === 'pend' }];
        const rows = last.blocks.map(b => {
            const r = { block: b.block };
            all.forEach(s => { const x = s.byBlock.get(b.block); r[s.date] = x ? x[metric] : null; });
            r.chg = r[last.date] != null && r[first.date] != null ? r[last.date] - r[first.date] : null;
            return r;
        });
        const label = sel.querySelector('option:checked').textContent;
        const tbl = makeTable({ cols, rows, sort: last.date });
        const bar = h('<div class="actions" style="margin-bottom:8px"></div>');
        bar.append(exportBtns(tableSpec(tbl, sTitle(`Block-wise Trend - ${label}`), `${fmtD(first.date)} to ${fmtD(last.date)}`)));
        holder.append(bar, tbl.el);
    };
    $('select', sel).onchange = e => draw(e.target.value);
    draw('p');
}

// ================================================================ Management / Category-wise report (both modules)
const groupColsFor = (school, label) => {
    const g = col(C.block, { h: label, cls: 'wrap' });
    return school
        ? [C.idx, g, SC.schools, SC.o, SC.pend, C.p, SC.dO, C.dP]
        : [C.idx, g, C.schools, C.t, C.o, col(C.pend, { invert: true }), C.p, C.dO, C.dP, C.full, C.part, C.zero];
};
const groupLink = (school, by, block) => r => {
    const q = new URLSearchParams({ [by]: r.block });
    if (block) q.set('block', block);
    location.hash = `${school ? '#/s/' : '#/'}schools?${q}`;
};

function pageGroups(v, r) {
    const S = cur(), P = prev(), school = state.mod === 'school';
    const by = GROUPS[r.q.get('by')] ? r.q.get('by') : 'mg';
    const block = r.q.get('block') || '';
    const label = GROUPS[by];
    const base = school ? '#/s/' : '#/';
    const title = `${MODS[state.mod]} - ${label}-wise Report${block ? ` - ${block} Block` : ''}`;
    pageHead(v, `${label}-wise Report (${MODS[state.mod]})`, subLine(S, P) + ' &middot; tap a row to see its schools');
    if (!state.master) {
        v.append(h('<div class="card"><div class="card-body">The school master is not loaded. Upload the School Search CSV (with <b>Management</b> and <b>Category</b> columns) on <a href="#/data">Upload / Manage Data</a>.</div></div>'));
        return;
    }

    const f = h(`<div class="filters">
        <label>Report by<select id="gBy">${Object.entries(GROUPS).map(([k, t]) => `<option value="${k}"${k === by ? ' selected' : ''}>${t}</option>`).join('')}</select></label>
        <label>Block<select id="gBlock"><option value="">All blocks</option>${blockNames().map(b => `<option${b === block ? ' selected' : ''}>${esc(b)}</option>`).join('')}</select></label>
    </div>`);
    v.append(f);
    const go = (b, bl) => { location.hash = `${base}groups?by=${b}${bl ? '&block=' + encodeURIComponent(bl) : ''}`; };
    $('#gBy', f).onchange = e => go(e.target.value, block);
    $('#gBlock', f).onchange = e => go(by, e.target.value);

    const rows = groupRows(S, P, by, block);
    const foot = totalsRow(rows);
    if (P) foot.dP = foot.p - totalsRow(groupRows(P, null, by, block)).p;
    tableCard(v, title, { cols: groupColsFor(school, label), rows, foot, sort: 'schools', onRow: groupLink(school, by, block) }, plainSub(S, P));

    chart(card(v, `Onboarded % by ${label}`), Math.max(220, rows.length * 34 + 60), {
        type: 'bar',
        data: {
            labels: rows.map(g => (by === 'ct' ? shortCat(g.block) : g.block)),
            datasets: [{ label: 'Onboarded %', data: rows.map(g => round2(g.p)), backgroundColor: rows.map(g => COLORS[pctCls(g.p)] || COLORS.brand2), borderRadius: 4, maxBarThickness: 26 }],
        },
        options: { indexAxis: 'y', plugins: { legend: { display: false } }, scales: { x: { beginAtZero: true, max: 100, ticks: { callback: x => x + '%' } }, y: { ticks: { autoSkip: false } } } },
    });

    // Block x group matrix
    const groups = rows.map(g => g.block);
    const metrics = school
        ? [['p', 'Onboarded %'], ['o', 'Schools onboarded'], ['pend', 'Pending schools'], ['schools', 'Total schools']]
        : [['p', 'Onboarded %'], ['o', 'Teachers onboarded'], ['pend', 'Pending teachers'], ['zero', '0% schools'], ['full', '100% schools'], ['schools', 'Total schools']];
    const box = card(v, `Block-wise by ${label}`);
    const sel = h(`<div class="filters"><label>Show<select>${metrics.map(([k, t]) => `<option value="${k}">${t}</option>`).join('')}</select></label></div>`);
    const holder = h('<div></div>');
    box.append(sel, holder);
    const cells = new Map(); // "block|group" -> aggregate
    S.schools.forEach(x => {
        const key = x.b + '|' + groupOf(x, by);
        if (!cells.has(key)) cells.set(key, blank(key));
        addTo(cells.get(key), x);
    });
    cells.forEach(finish);
    const draw = metric => {
        holder.innerHTML = '';
        const type = metric === 'p' ? 'p' : 'n';
        const head = g => (by === 'ct' ? shortCat(g) : g);
        const cols = [C.idx, C.block, ...groups.map((g, i) => ({ k: 'g' + i, h: head(g), type })), { k: 'all', h: 'All', type }];
        const mrows = S.blocks.map(b => {
            const row = { block: b.block, all: b[metric] };
            groups.forEach((g, i) => { const c = cells.get(b.block + '|' + g); row['g' + i] = c ? c[metric] : null; });
            return row;
        });
        const dist = groupRows(S, null, by, '');
        const mfoot = { block: 'District Total', all: S.total[metric] };
        groups.forEach((g, i) => { const d = dist.find(x => x.block === g); mfoot['g' + i] = d ? d[metric] : null; });
        const tbl = makeTable({ cols, rows: mrows, foot: mfoot });
        const bar = h('<div class="actions" style="margin-bottom:8px"></div>');
        const mLabel = sel.querySelector('option:checked').textContent;
        bar.append(exportBtns(tableSpec(tbl, `${MODS[state.mod]} - Block-wise by ${label} - ${mLabel}`, plainSub(S, P))));
        holder.append(bar, tbl.el);
    };
    $('select', sel).onchange = e => draw(e.target.value);
    draw('p');
    if (by === 'ct') v.append(h(`<p class="help">Category codes: ${groups.filter(g => g !== NO_MASTER).map(g => `<b>${esc(shortCat(g))}</b> = ${esc(g)}`).join(' &middot; ')}</p>`));
}

// Management and Category summary tables shown at the end of both dashboards
function groupSummary(v, S, P) {
    if (!state.master) return;
    const school = state.mod === 'school';
    const g = h('<div class="grid-2" style="margin-top:16px"></div>');
    v.append(g);
    ['mg', 'ct'].forEach(by => {
        const label = GROUPS[by];
        const cols = [col(C.block, { h: label, cls: 'wrap', v: by === 'ct' ? x => shortCat(x.block) : null }), school ? SC.schools : C.t, school ? SC.o : C.o, C.p, col(C.dO, { h: 'Change' })];
        const tbl = tableCard(g, `${MODS[state.mod]} by ${label}`, { cols, rows: groupRows(S, P, by, ''), sort: school ? 'schools' : 't', onRow: () => { location.hash = `${school ? '#/s/' : '#/'}groups?by=${by}`; } }, plainSub(S, P));
        tbl.el.append(h(`<p class="help" style="padding:0 16px 12px"><a href="${school ? '#/s/' : '#/'}groups?by=${by}">Full ${label.toLowerCase()}-wise report &rsaquo;</a></p>`));
    });
}

function emptyState() {
    const school = state.mod === 'school';
    return h(`<div class="card empty-state"><h2>No ${school ? 'school onboarding' : 'teacher onboarding'} data yet</h2>
        <p>Upload the combined ${school
        ? '<b>Schools_Onboarding_Status_All_Blocks</b> Excel file (Block Name, School Name, UDISE Code, Status)'
        : 'school-wise Excel file (the one with the <b>Block Name</b> column)'} to start monitoring.</p>
        <p><a class="btn primary" href="#/data">&#11014; Upload Excel</a></p>
        <p class="help">${CREDIT}</p></div>`);
}

// ---------------------------------------------------------------- upload handling
async function handleFiles(files) {
    if (!files.length) return;
    const override = $('#upDate') && $('#upDate').value;
    if (override && files.length > 1) { toast('Pick a date only when uploading a single file.', true); return; }
    const kinds = new Set();
    let masterChanged = false;
    for (const f of files) {
        try {
            const buf = await f.arrayBuffer();
            const master = parseMaster(buf);
            if (master) {
                await DB.put('meta', { key: 'master', fileName: f.name, uploadedAt: new Date().toISOString(), rows: master });
                masterChanged = true;
                toast(`School master ${f.name}: ${fmtN(master.length)} schools loaded`);
                continue;
            }
            const date = override || dateFromName(f.name);
            if (!date) throw new Error('no date in the file name; choose the "Data date" and upload again');
            const { kind, schools } = parseWorkbook(buf);
            const exists = state.mods[kind].snaps.find(s => s.date === date && s.source !== 'repo');
            if (exists && !confirm(`${MODS[kind]} data for ${fmtD(date)} already exists. Replace it with ${f.name}?`)) continue;
            await DB.put(kind, { date, fileName: f.name, uploadedAt: new Date().toISOString(), source: 'upload', schools });
            kinds.add(kind);
            toast(`${MODS[kind]} ${fmtD(date)}: ${fmtN(schools.length)} schools loaded`);
        } catch (e) {
            console.error(e);
            toast(`${f.name}: ${e.message}`, true);
        }
    }
    if (kinds.size || masterChanged) await reload(kinds);
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
                out.push({ date, fileName: file, source: 'repo', ...parseWorkbook(await r.arrayBuffer()) });
            } catch (err) { toast(`data/${file}: ${err.message}`, true); }
        }
        return out;
    } catch (e) { return []; }
}

// School master: data/school_master.csv in the repo, overridden by an uploaded master in this browser
async function loadRepoMaster() {
    try {
        const r = await fetch('data/school_master.csv', { cache: 'no-store' });
        if (!r.ok) return null;
        const rows = parseMaster(await r.arrayBuffer());
        return rows ? { fileName: 'school_master.csv', source: 'repo', rows } : null;
    } catch (e) { return null; }
}

let repoSnaps = null, repoMaster;
// jumpTo: a Set of modules whose date should move to their latest day (after an upload)
async function reload(jumpTo) {
    if (repoSnaps === null) [repoSnaps, repoMaster] = await Promise.all([loadRepo(), loadRepoMaster()]);
    const m = (await DB.all('meta')).find(x => x.key === 'master') || repoMaster;
    state.master = m ? new Map(m.rows.map(x => [x.u, x])) : null;
    state.masterInfo = m ? { fileName: m.fileName, source: m.source || 'upload', uploadedAt: m.uploadedAt, count: m.rows.length } : null;
    const active = state.mod;
    for (const kind of Object.keys(MODS)) {
        const local = await DB.all(kind);
        const map = new Map(repoSnaps.filter(s => s.kind === kind).map(s => [s.date, s]));
        local.forEach(s => map.set(s.date, { ...s, kind })); // uploaded data overrides repo data of the same date
        state.mod = kind;
        state.snaps = [...map.values()].sort(byDate);
        const dates = state.snaps.map(s => s.date);
        if ((jumpTo && jumpTo.has(kind)) || !dates.includes(state.cur)) setCur(dates[dates.length - 1] || '');
        else if (state.prev && !dates.includes(state.prev)) setCur(state.cur);
    }
    state.mod = active;
    state.cache = {};
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
// Teacher pages: #/<page>, school pages: #/s/<page>, the data page (#/data) is shared and keeps the active module
const routes = {
    teacher: { dashboard: pageDashboard, blocks: pageBlocks, schools: pageSchools, compare: pageCompare, trend: pageTrend, rank: pageRank, groups: pageGroups, data: pageData },
    school: { dashboard: sDashboard, blocks: sBlocks, schools: sSchools, compare: sCompare, trend: sTrend, rank: sRank, groups: pageGroups, data: pageData },
};

function parseHash() {
    const raw = location.hash.replace(/^#\/?/, '');
    const [path, q = ''] = raw.split('?');
    const parts = path.split('/').filter(Boolean);
    const mod = parts[0] === 's' ? 'school' : parts[0] === 'data' ? state.mod : 'teacher';
    if (parts[0] === 's') parts.shift();
    return { mod, name: routes[mod][parts[0]] ? parts[0] : 'dashboard', arg: parts[1] || '', q: new URLSearchParams(q) };
}

function setMod(mod) {
    if (state.mod === mod) return;
    state.mod = mod;
    fillDateSelects();
}

function render() {
    state.charts.forEach(c => c.destroy());
    state.charts = [];
    const r = parseHash();
    setMod(r.mod);
    document.body.classList.toggle('mod-school', r.mod === 'school');
    document.querySelectorAll('.mod-switch [data-mod]').forEach(b => b.classList.toggle('active', b.dataset.mod === r.mod));
    const key = (r.mod === 'school' && r.name !== 'data' ? 's/' : '') + r.name + (r.arg ? '/' + r.arg : '');
    document.querySelectorAll('#sidebar nav a').forEach(a => a.classList.toggle('active', a.dataset.route === key));
    const v = $('#view');
    v.innerHTML = '';
    document.body.classList.remove('menu-open');
    try {
        if (!state.snaps.length && r.name !== 'data') v.append(emptyState());
        else routes[r.mod][r.name](v, r);
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
    // Module switch keeps the same page (dashboard, blocks, ...) where possible
    document.querySelectorAll('.mod-switch [data-mod]').forEach(b => { b.onclick = () => {
        const r = parseHash();
        const page = r.name === 'data' ? 'dashboard' : r.name + (r.arg && b.dataset.mod === r.mod ? '/' + r.arg : '');
        location.hash = (b.dataset.mod === 'school' ? '#/s/' : '#/') + page;
    }; });
    window.addEventListener('hashchange', () => { render(); window.scrollTo(0, 0); });
    await DB.open();
    await reload();
}

init();
