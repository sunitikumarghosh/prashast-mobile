// ==UserScript==
// @name         PRASHAST - Teacher Onboarding: All Blocks -> One Excel
// @namespace    prashast-tools
// @version      1.0
// @description  Visits every block in "Teacher Onboarding | Block Wise Report", collects the school-wise table of each, and saves one combined Excel with a Block Name column.
// @match        https://prashast.education.gov.in/prashast-admin/*
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @grant        none
// @run-at       document-idle
// ==/UserScript==

(function () {
    'use strict';

    // ---------------- Settings ----------------
    // true = also click "Download > Download Excel" inside every block (44 separate files, like doing it by hand)
    const DOWNLOAD_EACH_BLOCK_EXCEL = false;
    // Extra wait (ms) after each page/table change. Increase if the site is slow.
    const SETTLE_MS = 700;
    const TIMEOUT_MS = 30000;
    // ------------------------------------------

    let stopRequested = false;

    // ---------- small helpers ----------
    const sleep = ms => new Promise(r => setTimeout(r, ms));
    const norm = s => (s || '').replace(/\s+/g, ' ').trim();

    function isVisible(el) {
        if (!el) return false;
        const r = el.getBoundingClientRect();
        const st = getComputedStyle(el);
        return r.width > 0 && r.height > 0 && st.visibility !== 'hidden' && st.display !== 'none';
    }

    async function waitFor(fn, what, timeout = TIMEOUT_MS) {
        const t0 = Date.now();
        while (Date.now() - t0 < timeout) {
            if (stopRequested) throw new Error('Stopped by user');
            let v;
            try { v = fn(); } catch (e) { v = null; }
            if (v) return v;
            await sleep(200);
        }
        throw new Error('Timed out waiting for: ' + what);
    }

    function realClick(el) {
        el.scrollIntoView({ block: 'center' });
        ['pointerdown', 'mousedown', 'pointerup', 'mouseup', 'click'].forEach(type =>
            el.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, view: window })));
    }

    function isLoading() {
        return [...document.querySelectorAll('.ant-spin-spinning, .ant-spin-dot-spin, .spinner, .loader, .loading')]
            .some(isVisible);
    }

    // Deepest visible element whose text contains `text`
    function findTextEl(text, root = document.body) {
        let best = null;
        root.querySelectorAll('*').forEach(el => {
            if (!isVisible(el)) return;
            if (!norm(el.textContent).includes(text)) return;
            if (!best || best.contains(el)) best = el;
        });
        return best;
    }

    function currentView() {
        const t = norm(document.body.innerText);
        if (t.includes('School Wise Report')) return 'school';
        if (t.includes('Block Wise Report')) return 'block';
        return null;
    }

    // ---------- table helpers ----------
    function getTable(firstHeaderMatch) {
        const th = [...document.querySelectorAll('th')]
            .find(h => isVisible(h) && norm(h.textContent).includes(firstHeaderMatch));
        if (!th) return null;
        const root = th.closest('.ant-table') || th.closest('table');
        const headers = [...root.querySelectorAll('thead th')].map(h => norm(h.textContent)).filter(Boolean);
        const rows = [...root.querySelectorAll('tbody tr')].filter(tr =>
            !tr.classList.contains('ant-table-placeholder') &&
            !tr.classList.contains('ant-table-measure-row') &&
            tr.querySelectorAll('td').length >= 2 && isVisible(tr));
        return { root, headers, rows };
    }

    const rowsSignature = tbl => tbl ? tbl.rows.map(r => norm(r.innerText)).join('\n') : '';

    async function waitTableStable(header) {
        let last = null, stableSince = Date.now();
        await waitFor(() => {
            if (isLoading()) { stableSince = Date.now(); return false; }
            const tbl = getTable(header);
            if (!tbl) return false;
            const sig = rowsSignature(tbl);
            if (sig !== last) { last = sig; stableSince = Date.now(); return false; }
            return Date.now() - stableSince >= SETTLE_MS;
        }, 'table "' + header + '" to settle');
        return getTable(header);
    }

    // ---------- pagination helpers ----------
    function getPagination() {
        return [...document.querySelectorAll('.ant-pagination, nz-pagination, [class*="pagination"]')]
            .find(isVisible) || null;
    }

    function totalRecords() {
        const m = norm(document.body.innerText).match(/(\d+)\s+records/i);
        return m ? parseInt(m[1], 10) : null;
    }

    function nextButton() {
        const p = getPagination();
        if (!p) return null;
        const btn = p.querySelector('.ant-pagination-next, [title="Next Page"], [aria-label="Next Page"]');
        if (!btn) return null;
        const disabled = btn.classList.contains('ant-pagination-disabled') ||
            btn.getAttribute('aria-disabled') === 'true' ||
            (btn.querySelector('button') && btn.querySelector('button').disabled);
        return disabled ? null : btn;
    }

    async function goToFirstPage(header) {
        const p = getPagination();
        if (!p) return;
        const first = [...p.querySelectorAll('.ant-pagination-item, li')].find(li => norm(li.textContent) === '1');
        if (first && !first.classList.contains('ant-pagination-item-active')) {
            const before = rowsSignature(getTable(header));
            realClick(first.querySelector('a') || first);
            await waitFor(() => rowsSignature(getTable(header)) !== before, 'page 1');
            await waitTableStable(header);
        }
    }

    // Try to pick the largest "xx / page" option so fewer pages are needed.
    async function setMaxPageSize(header) {
        const p = getPagination();
        if (!p) return;
        const sel = p.querySelector('.ant-pagination-options .ant-select, .ant-pagination-options nz-select, .ant-select');
        if (!sel) return;
        const trigger = sel.querySelector('.ant-select-selector, nz-select-top-control') || sel;
        const optSel = '.ant-select-item-option, nz-option-item, .ant-select-dropdown-menu-item';
        const options = () => [...document.querySelectorAll(optSel)].filter(isVisible);

        trigger.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
        await sleep(300);
        if (!options().length) { trigger.click(); await sleep(300); }
        const opts = options();
        if (!opts.length) return;

        const sizeOf = o => parseInt((norm(o.textContent).match(/\d+/) || ['0'])[0], 10);
        const best = opts.reduce((a, b) => (sizeOf(b) > sizeOf(a) ? b : a));
        const currentText = norm(sel.textContent);
        if (currentText.includes(String(sizeOf(best)))) { document.body.click(); return; }

        const before = rowsSignature(getTable(header));
        realClick(best);
        await waitFor(() => rowsSignature(getTable(header)) !== before, 'page size change', 10000).catch(() => {});
        await waitTableStable(header);
    }

    // Read all pages of the current table, calling onRow(cellsTexts, rowEl) for every row.
    async function readAllPages(header) {
        const out = [];
        let tbl = await waitTableStable(header);
        for (let guard = 0; guard < 200; guard++) {
            tbl.rows.forEach(tr => out.push([...tr.querySelectorAll('td')].map(td => norm(td.innerText))));
            const nb = nextButton();
            if (!nb) break;
            const before = rowsSignature(tbl);
            realClick(nb.querySelector('button, a') || nb);
            await waitFor(() => rowsSignature(getTable(header)) !== before, 'next page');
            tbl = await waitTableStable(header);
        }
        return { headers: tbl.headers, rows: out };
    }

    // ---------- navigation ----------
    async function openBlock(blockName) {
        await waitFor(() => currentView() === 'block', 'Block Wise Report view');
        await goToFirstPage('Block Name');
        for (let guard = 0; guard < 50; guard++) {
            const tbl = await waitTableStable('Block Name');
            const nameIdx = tbl.headers.findIndex(h => h.includes('Block Name'));
            const row = tbl.rows.find(tr => norm(tr.querySelectorAll('td')[nameIdx].innerText) === blockName);
            if (row) {
                const cell = row.querySelectorAll('td')[nameIdx];
                const btn = cell.querySelector('button, a, [role="button"], i, svg, .anticon') ||
                    row.querySelector('button, a, [role="button"]');
                realClick(btn || cell);
                await waitFor(() => currentView() === 'school' &&
                    norm(document.body.innerText).includes(blockName) &&
                    getTable('School Name'), 'school table of ' + blockName);
                return;
            }
            const nb = nextButton();
            if (!nb) break;
            const before = rowsSignature(tbl);
            realClick(nb.querySelector('button, a') || nb);
            await waitFor(() => rowsSignature(getTable('Block Name')) !== before, 'next block page');
        }
        throw new Error('Block not found in list: ' + blockName);
    }

    async function backToBlocks() {
        const title = findTextEl('School Wise Report');
        let target = null;
        // 1) the arrow just left of the title
        if (title) {
            let node = title;
            for (let i = 0; i < 4 && node && !target; i++, node = node.parentElement) {
                let prev = node.previousElementSibling;
                while (prev && !target) {
                    if (isVisible(prev)) target = prev.querySelector('button, a, i, svg, .anticon') || prev;
                    prev = prev.previousElementSibling;
                }
            }
        }
        // 2) any visible arrow-left icon
        if (!target) {
            target = [...document.querySelectorAll('.anticon-arrow-left, [nztype="arrow-left"], [class*="back"]')]
                .find(isVisible);
        }
        if (target) {
            realClick(target.closest('button, a') || target);
            try { await waitFor(() => currentView() === 'block', 'Block view', 8000); return; } catch (e) { /* try breadcrumb */ }
        }
        // 3) breadcrumb "Block"
        const crumb = [...document.querySelectorAll('.ant-breadcrumb a, .ant-breadcrumb-link, nz-breadcrumb-item a, nz-breadcrumb-item span')]
            .find(el => isVisible(el) && norm(el.textContent) === 'Block');
        if (crumb) realClick(crumb);
        await waitFor(() => currentView() === 'block', 'Block view after back');
    }

    async function clickDownloadExcel() {
        const dl = [...document.querySelectorAll('button')].find(b => isVisible(b) && norm(b.textContent).startsWith('Download'));
        if (!dl) return;
        realClick(dl);
        const item = await waitFor(() => [...document.querySelectorAll('li, a, button, span, div')]
            .find(el => isVisible(el) && norm(el.textContent) === 'Download Excel'), 'Download Excel menu', 5000).catch(() => null);
        if (item) { realClick(item); await sleep(2500); }
        document.body.click();
    }

    // ---------- parsing ----------
    const toNum = v => (v !== '' && !isNaN(Number(v.replace(/,/g, '')))) ? Number(v.replace(/,/g, '')) : v;

    function schoolRowToObj(blockName, headers, cells) {
        const o = { 'Block Name': blockName };
        headers.forEach((h, i) => {
            const v = cells[i] ?? '';
            if (/School Name/i.test(h)) {
                const k = v.lastIndexOf('|');
                o['School Name'] = k >= 0 ? v.slice(0, k).trim() : v;
                o['UDISE Code'] = k >= 0 ? v.slice(k + 1).trim() : '';
            } else if (/S\.?\s*No/i.test(h)) {
                o['S. No. (in block)'] = toNum(v);
            } else {
                o[h] = toNum(v);
            }
        });
        return o;
    }

    // ---------- main ----------
    async function run() {
        stopRequested = false;
        if (typeof XLSX === 'undefined') throw new Error('SheetJS (XLSX) not loaded');
        if (currentView() === 'school') await backToBlocks();
        if (currentView() !== 'block') throw new Error('Open "Teacher Onboarding | Block Wise Report" first.');

        status('Reading block list…');
        await setMaxPageSize('Block Name');
        await goToFirstPage('Block Name');
        const blockData = await readAllPages('Block Name');
        const nameIdx = blockData.headers.findIndex(h => h.includes('Block Name'));
        const summary = blockData.rows.map(c => {
            const o = {};
            blockData.headers.forEach((h, i) => { o[h] = toNum(c[i] ?? ''); });
            return o;
        });
        const blocks = blockData.rows.map(c => c[nameIdx]).filter(Boolean);
        status(`Found ${blocks.length} blocks`);

        const all = [];
        const problems = [];
        for (let i = 0; i < blocks.length; i++) {
            const b = blocks[i];
            status(`(${i + 1}/${blocks.length}) ${b}: opening…`);
            await openBlock(b);
            await setMaxPageSize('School Name');
            await goToFirstPage('School Name');
            const expected = totalRecords();
            const data = await readAllPages('School Name');
            data.rows.forEach(c => all.push(schoolRowToObj(b, data.headers, c)));
            if (expected != null && expected !== data.rows.length) {
                problems.push(`${b}: page says ${expected} records, collected ${data.rows.length}`);
            }
            status(`(${i + 1}/${blocks.length}) ${b}: ${data.rows.length} schools (total ${all.length})`);
            if (DOWNLOAD_EACH_BLOCK_EXCEL) await clickDownloadExcel();
            await backToBlocks();
        }

        // Global serial number first
        const rows = all.map((r, i) => Object.assign({ 'S. No.': i + 1 }, r));
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(rows);
        ws['!cols'] = Object.keys(rows[0] || {}).map(k => ({ wch: k === 'School Name' ? 50 : Math.max(12, k.length + 2) }));
        ws['!autofilter'] = { ref: ws['!ref'] };
        XLSX.utils.book_append_sheet(wb, ws, 'All Schools');
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summary), 'Block Summary');
        if (problems.length) {
            XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['Count mismatches'], ...problems.map(p => [p])]), 'Check');
        }
        const d = new Date();
        const stamp = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
        XLSX.writeFile(wb, `Teacher_Onboarding_School_Wise_All_Blocks_${stamp}.xlsx`);

        status(`Done: ${all.length} schools from ${blocks.length} blocks.` +
            (problems.length ? ` ⚠ ${problems.length} count mismatch(es), see "Check" sheet.` : ''));
    }

    // ---------- UI ----------
    let statusEl;
    function status(msg) {
        if (statusEl) statusEl.textContent = msg;
        console.log('[PRASHAST export]', msg);
    }

    function addPanel() {
        if (document.getElementById('prashast-export-panel')) return;
        const box = document.createElement('div');
        box.id = 'prashast-export-panel';
        box.style.cssText = 'position:fixed;left:12px;bottom:12px;z-index:99999;background:#fff;border:1px solid #1a237e;' +
            'border-radius:8px;padding:8px 10px;font:12px sans-serif;box-shadow:0 2px 10px rgba(0,0,0,.2);max-width:320px';
        const start = document.createElement('button');
        start.textContent = '⬇ Export all blocks (school wise)';
        start.style.cssText = 'background:#1a237e;color:#fff;border:0;border-radius:4px;padding:6px 10px;cursor:pointer;margin-right:6px';
        const stop = document.createElement('button');
        stop.textContent = 'Stop';
        stop.style.cssText = 'border:1px solid #999;background:#fff;border-radius:4px;padding:6px 10px;cursor:pointer';
        statusEl = document.createElement('div');
        statusEl.style.cssText = 'margin-top:6px;color:#333';
        statusEl.textContent = 'Open Teacher Onboarding → Block Wise Report, then click Export.';

        start.onclick = async () => {
            start.disabled = true;
            try { await run(); } catch (e) { status('Error: ' + e.message); console.error(e); }
            start.disabled = false;
        };
        stop.onclick = () => { stopRequested = true; status('Stopping…'); };
        box.append(start, stop, statusEl);
        document.body.appendChild(box);
    }

    addPanel();
    setInterval(addPanel, 3000); // re-add if the SPA re-renders the body
})();
