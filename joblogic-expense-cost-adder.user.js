// ==UserScript==
// @name         Joblogic - Add expense costs to jobs (check + add, Excel upload)
// @namespace    http://tampermonkey.net/
// @version      1.2
// @description  Upload (or paste) an expense export (Payhawk .xlsx/.csv, Pleo also works). For each row it finds the job (Job ref / Salesforce ref / Quote UP-number from Job Ref/Details + notes), reads the job's Costs page and checks whether the receipt NET is already on it (same matching as the Cost Reconciler). Rows NOT in the job are ticked to ADD as a Material line (cost = Net, qty 1, 20% VAT, job uplift; project/quoted jobs added NON-chargeable). Rows already in the job (or entered incorrectly, or on an undelivered PO) are left alone. Check first, untick anything you don't want, then Add ticked to jobs.
// @match        https://go.joblogic.com/*
// @require      https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js
// @grant        none
// @run-at       document-idle
// @updateURL    https://raw.githubusercontent.com/loisstarling-coder/joblogic-expense-cost-adder/main/joblogic-expense-cost-adder.user.js
// @downloadURL  https://raw.githubusercontent.com/loisstarling-coder/joblogic-expense-cost-adder/main/joblogic-expense-cost-adder.user.js
// ==/UserScript==

(function () {
    'use strict';

    // ===================================================================
    // CONFIG
    // ===================================================================
    const CFG = {
        taxCodeId: 'c1d73a68-7887-4f91-9124-26100ef712b0', taxCodeValue: '20.00',
        taxCodeDesc: '20% (VAT on Income) (20.00%)',
        payBandId: 'f76846b0-674b-4410-8473-c2f22508f51c', payBandDesc: 'Basic',
        libraryId: 54838, libraryName: 'Standard Parts Library'
    };
    const DELAY_BETWEEN_ROWS = 250;     // ms politeness delay between sheet rows (check)
    const DELAY_BETWEEN_WRITES = 400;   // ms politeness delay between writes
    const MONEY_TOL_ABS = 0.02;         // absolute £ tolerance for a "match"
    const MONEY_TOL_PCT = 0.005;        // + 0.5% relative tolerance (rounding)
    const VAT_FALLBACK = 0.20;

    // ===================================================================
    // STATE
    // ===================================================================
    let panel, fileInput, fileLabel, pasteArea, logArea, checkBtn, addBtn, stopBtn, copyBtn, progressText, resultsBox;
    let running = false;
    let results = [];                   // per-row outcome
    let loadedSheet = null;             // {headers, rows} from an uploaded file
    const jobCache = new Map();         // jobId -> { lines, parseOk }
    const searchCache = new Map();      // term -> {id, jobNumber,...} | null
    const upliftCache = new Map();      // jobId -> {uplift, pct}

    const SCRIPT_ID = 'expense-cost-adder';
    const SCRIPT_LABEL = '➕ Add expense costs to jobs';
    const SCRIPT_VERSION = ((typeof GM_info !== 'undefined' && GM_info.script && GM_info.script.version) || '1.2');
    const SCRIPT_COLOR = '#4c9f01';
    const SCRIPT_DESC = 'Upload the Payhawk expense export (.xlsx) or paste it with the header row, then click Check. Each row is checked against its job\'s costs: already in job → left alone; not in job → ticked to add. Untick anything you don\'t want, then Add ticked to jobs.';

    const sleep = (ms) => new Promise(r => setTimeout(r, ms));
    const csrf = () => document.querySelector('input[name="__RequestVerificationToken"]')?.value || '';

    // ===================================================================
    // NUMBER / TEXT HELPERS (from the reconciler)
    // ===================================================================
    function money(v) {
        if (v == null) return null;
        const s = String(v).replace(/[£$,\s]/g, '').replace(/[()]/g, '');
        const n = parseFloat(s);
        return isNaN(n) ? null : Math.abs(n);
    }
    function isNegative(v) { return /^\s*(-|\()/.test(String(v == null ? '' : v)); }
    function vatRate(v) {
        if (v == null || v === '') return null;
        const s = String(v);
        const pct = s.match(/(\d+(?:\.\d+)?)\s*%/);
        if (pct) return parseFloat(pct[1]) / 100;
        const num = parseFloat(s);
        if (isNaN(num)) return null;
        return num > 1 ? num / 100 : num;
    }
    function qtyNum(v) {
        if (v == null || v === '') return null;
        const s = String(v).trim();
        const hm = s.match(/(\d+)\s*h\s*(\d+)?\s*m?/i);
        if (hm) return parseInt(hm[1], 10) + (hm[2] ? parseInt(hm[2], 10) / 60 : 0);
        const n = parseFloat(s.replace(/,/g, ''));
        return isNaN(n) ? null : n;
    }
    const norm = (s) => String(s == null ? '' : s).toLowerCase().replace(/\s+/g, ' ').trim();
    function eqMoney(a, b) {
        if (a == null || b == null) return false;
        const tol = Math.max(MONEY_TOL_ABS, Math.abs(b) * MONEY_TOL_PCT, Math.abs(a) * MONEY_TOL_PCT);
        return Math.abs(a - b) <= tol;
    }
    function nearInteger(ratio) {
        if (!isFinite(ratio) || ratio < 1.5) return null;
        const k = Math.round(ratio);
        return Math.abs(ratio - k) <= 0.02 ? k : null;
    }
    const cap = (s) => s.charAt(0).toUpperCase() + s.slice(1);
    const esc = (s) => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

    // ===================================================================
    // DATES
    // Excel uploads give serial numbers (46233 = 29/07/2026); pasted Payhawk rows are
    // D/M/Y; Pleo exports are US M/D/Y. Always emit JobLogic's DD/MM/YYYY, never a
    // month > 12 (that silently breaks AddMaterialCosts).
    // ===================================================================
    let inputDateMDY = false;
    function normalizeDate(v) {
        if (typeof v === 'number' || /^\d{5}(\.\d+)?$/.test(String(v || '').trim())) {
            const n = Number(v);
            if (n > 20000 && n < 80000) {
                const d = new Date(Math.round((n - 25569) * 864e5));
                return `${String(d.getUTCDate()).padStart(2, '0')}/${String(d.getUTCMonth() + 1).padStart(2, '0')}/${d.getUTCFullYear()}`;
            }
            return '';
        }
        const iso = String(v || '').match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
        if (iso) return `${iso[3].padStart(2, '0')}/${iso[2].padStart(2, '0')}/${iso[1]}`;
        const m = String(v || '').match(/(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})/);
        if (!m) return '';
        const a = parseInt(m[1], 10), b = parseInt(m[2], 10);
        const y = m[3].length === 2 ? '20' + m[3] : m[3];
        let day, mo;
        if (a > 12) { day = a; mo = b; }
        else if (b > 12) { mo = a; day = b; }
        else if (inputDateMDY) { mo = a; day = b; }
        else { day = a; mo = b; }
        if (mo < 1 || mo > 12) return '';
        return `${String(day).padStart(2, '0')}/${String(mo).padStart(2, '0')}/${y}`;
    }
    function detectDateOrder(rows, headers) {
        let mdy = 0, dmy = 0;
        for (const r of rows) {
            const m = String(dateCell(r) || '').match(/^(\d{1,2})[-/.](\d{1,2})[-/.]/);
            if (!m) continue;
            const a = parseInt(m[1], 10), b = parseInt(m[2], 10);
            if (a > 12 && b <= 12) dmy++;
            else if (b > 12 && a <= 12) mdy++;
        }
        if (mdy !== dmy) return mdy > dmy;
        // undecided: Pleo is US M/D/Y, everything else (Payhawk, UK Excel) is D/M/Y
        return headers.some(h => /orig\.? amount/i.test(h)) && !headers.some(h => /expense id/i.test(h));
    }
    function todayDMY() {
        const d = new Date();
        return `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`;
    }

    // ===================================================================
    // REFERENCE EXTRACTION (from the reconciler)
    // ===================================================================
    const RE_JOBREF = /\b(?:SRE|RE|PROJ|PM|M|R)\s?0*\d{2,}\b/gi;
    const RE_SF      = /\b1\d{7}\b/g;
    const RE_QUOTE   = /\bUP\s?0*\d{2,}\b/gi;
    const normJobRef = (raw) => raw.replace(/\s+/g, '').toUpperCase();
    function extractRefs(text) {
        const t = String(text || '');
        const out = { job: [], sf: [], quote: [], internal: [] };
        (t.match(RE_QUOTE) || []).forEach(m => {
            const code = m.replace(/\s+/g, '').toUpperCase();
            // Real quotes are UP0xxxx. UP1000 etc. are internal materials codes, not quotes.
            if (/^UP0\d+$/.test(code)) out.quote.push(code);
            else out.internal.push(code);
        });
        const t2 = t.replace(RE_QUOTE, ' ');
        (t2.match(RE_JOBREF) || []).forEach(m => out.job.push(normJobRef(m)));
        (t.match(RE_SF) || []).forEach(m => out.sf.push(m));
        out.job = [...new Set(out.job)];
        out.sf = [...new Set(out.sf)];
        out.quote = [...new Set(out.quote)];
        out.internal = [...new Set(out.internal)];
        return out;
    }
    function jobRefCandidates(ref) {
        const m = String(ref).toUpperCase().match(/^(SRE|RE|PROJ|PM|M|R)0*(\d+)$/);
        if (!m) return [ref];
        const prefix = m[1], num = m[2];
        return [...new Set([prefix + num.padStart(7, '0'), ref, prefix + ' ' + num, prefix + num])];
    }

    // ===================================================================
    // API: JOB / QUOTE SEARCH (from the reconciler)
    // ===================================================================
    async function searchJob(term) {
        if (searchCache.has(term)) return searchCache.get(term);
        const r = await fetch('/api/Job/SearchJsonData', {
            method: 'POST', credentials: 'same-origin',
            headers: { 'Content-Type': 'application/json', 'X-Requested-With': 'XMLHttpRequest', '__RequestVerificationToken': csrf() },
            body: JSON.stringify({
                SearchTerm: term, PageSize: 10, PageIndex: 1, EngineerType: 0,
                IncludePPMJobs: true, IncludeReactiveJobs: true,
                StartLoggedDate: '', EndLoggedDate: '', StartDate: '', EndDate: '',
                StartCompleteDate: '', EndCompleteDate: '', StartNextContactDate: '', EndNextContactDate: ''
            })
        });
        if (!r.ok) throw new Error('Job search HTTP ' + r.status);
        const d = await r.json();
        const jobs = d.AdditionalData?.Jobs || d.Data || [];
        let res = null;
        if (jobs.length) {
            const exact = jobs.find(j => norm(j.JobNumber) === norm(term) || norm(j.ReferenceNumber) === norm(term));
            const m = exact || jobs[0];
            res = { id: m.Id || m.JobId, jobNumber: m.JobNumber || m.ReferenceNumber || term, exact: !!exact, jobStatus: m.StatusDescription || '' };
        }
        searchCache.set(term, res);
        return res;
    }
    const QUOTE_FIELDS = ['CustomerId', 'SiteId', 'AssetId', 'searchTerm', 'statusIds', 'startDate', 'endDate',
        'ownerIds', 'QuoteDepositStatus', 'chanceOfSale', 'expiryStartDate', 'expiryEndDate',
        'nextContactStartDate', 'nextContactEndDate', 'saleStartDate', 'saleEndDate', 'tagIds',
        'excludeTagIds', 'tradeIds', 'priorityIds', 'assetClassIds', 'areaIds', 'siteTypeIds',
        'OrderBy', 'SelectedTab'];
    async function searchQuote(quoteNumber) {
        const key = 'Q:' + quoteNumber;
        if (searchCache.has(key)) return searchCache.get(key);
        const fd = new FormData();
        QUOTE_FIELDS.forEach(n => fd.append(n, n === 'searchTerm' ? quoteNumber : (n === 'SelectedTab' ? 'All' : '')));
        fd.append('PageIndex', '1');
        fd.append('PageSize', '25');
        const r = await fetch('/api/Quote/QuoteSearchJson', {
            method: 'POST', credentials: 'same-origin',
            headers: { 'X-Requested-With': 'XMLHttpRequest', '__RequestVerificationToken': csrf() },
            body: fd
        });
        if (!r.ok) throw new Error('Quote search HTTP ' + r.status);
        const d = await r.json();
        const quotes = d.AdditionalData?.Quotes || [];
        let res = null;
        if (quotes.length) {
            const q = quotes.find(q => norm(q.QuoteNumber) === norm(quoteNumber)) || quotes[0];
            res = { id: q.Id, quoteNumber: q.QuoteNumber, parentJobId: q.ParentJobAutoId || q.JobId || null, parentJobNumber: q.ParentJobStringId || q.JobNumber || null };
        }
        searchCache.set(key, res);
        return res;
    }
    async function resolveJob(refs) {
        for (const ref of refs.job) {
            let fallback = null;
            for (const cand of jobRefCandidates(ref)) {
                const j = await searchJob(cand);
                if (j && j.exact) return { ...j, via: 'job ref ' + j.jobNumber };
                if (j && !fallback) fallback = { ...j, via: 'job ref ' + ref + ' → ' + j.jobNumber };
            }
            if (fallback) return fallback;
        }
        for (const ref of refs.sf) {
            const j = await searchJob(ref); if (j) return { ...j, via: 'Salesforce ' + ref };
        }
        for (const q of refs.quote) {
            const qr = await searchQuote(q);
            if (qr && qr.parentJobId) {
                const jn = qr.parentJobNumber || ('job ' + qr.parentJobId);
                return { id: qr.parentJobId, jobNumber: jn, via: `quote ${qr.quoteNumber} → ${jn}` };
            }
            if (qr && qr.id) return { id: null, jobNumber: null, via: `quote ${qr.quoteNumber} (no upgraded job)` };
        }
        return null;
    }

    // ===================================================================
    // API: READ JOB COSTS / RELATED QUOTE / UNDELIVERED PO (from the reconciler)
    // ===================================================================
    function extractJsLiteral(html, varName) {
        const i = html.indexOf(varName);
        if (i < 0) return null;
        const eq = html.indexOf('=', i);
        if (eq < 0) return null;
        let s = eq + 1;
        while (s < html.length && /\s/.test(html[s])) s++;
        if (html[s] !== '{' && html[s] !== '[') return null;
        let depth = 0, inStr = false, escp = false, end = -1;
        for (let j = s; j < html.length; j++) {
            const c = html[j];
            if (escp) { escp = false; continue; }
            if (c === '\\') { escp = true; continue; }
            if (c === '"') { inStr = !inStr; continue; }
            if (inStr) continue;
            if (c === '{' || c === '[') depth++;
            else if (c === '}' || c === ']') { depth--; if (depth === 0) { end = j + 1; break; } }
        }
        if (end < 0) return null;
        try { return JSON.parse(html.slice(s, end)); } catch (e) { return null; }
    }
    async function getJobCosts(jobId, fresh) {
        if (!fresh && jobCache.has(jobId)) return jobCache.get(jobId);
        const r = await fetch(`/Job/GetCosts?jobId=${jobId}&isReadOnly=False`, {
            credentials: 'same-origin', headers: { 'X-Requested-With': 'XMLHttpRequest' }
        });
        if (!r.ok) throw new Error('GetCosts HTTP ' + r.status);
        const html = await r.text();
        const model = extractJsLiteral(html, 'JobLinesPM');
        const lines = [];
        if (model && Array.isArray(model.GroupedLines)) {
            model.GroupedLines.forEach(g => {
                (g.Lines || []).forEach(l => {
                    lines.push({
                        source: 'job', category: g.Heading || l.LineType || '', lineType: l.LineType || g.Heading || '',
                        id: l.Id, invoiced: !!l.HasBeenInvoiced, desc: l.Description || '', date: l.DateIncurred || '',
                        engineer: l.EngineerName || '', qtyRaw: l.Quantity, qty: qtyNum(l.Quantity),
                        unit: money(l.CostPerUnit), vat: vatRate(l.VatRate), subtotal: money(l.SubTotalCost)
                    });
                });
            });
        }
        const out = { lines, parseOk: !!model };
        jobCache.set(jobId, out);
        return out;
    }
    async function getUndeliveredPOLines(jobId) {
        const out = [];
        try {
            const r = await fetch(`/api/JobCost/GetRemainingUndeliveredPOLineItems?jobId=${jobId}`, {
                credentials: 'same-origin', headers: { 'X-Requested-With': 'XMLHttpRequest' }
            });
            if (!r.ok) return out;
            const j = await r.json();
            (j.AdditionalData?.Items || []).forEach(it => {
                const unit = money(it.ListPrice), q = qtyNum(it.Quantity);
                out.push({
                    source: 'undelivered PO ' + (it.PurchaseOrderNo || ''), poNo: it.PurchaseOrderNo || '',
                    category: 'Material', desc: it.Description || '', date: it.Date || '', engineer: '',
                    qtyRaw: it.Quantity, qty: q, unit, vat: null, subtotal: (unit != null && q != null) ? unit * q : unit
                });
            });
        } catch (e) { /* optional */ }
        return out;
    }
    // Only the quote info is needed here (to decide chargeable vs non-chargeable).
    async function getRelatedQuotes(jobId) {
        const r = await fetch(`/Job/GetRelatedWorks?jobId=${jobId}`, {
            credentials: 'same-origin', headers: { 'X-Requested-With': 'XMLHttpRequest' }
        });
        if (!r.ok) throw new Error('GetRelatedWorks HTTP ' + r.status);
        const html = await r.text();
        const doc = new DOMParser().parseFromString(html, 'text/html');
        const quotes = [], seen = new Set();
        const addQuote = (num, id, status) => {
            const key = String(num || id || '').toUpperCase();
            if (!key || seen.has(key)) return;
            seen.add(key);
            quotes.push({ num: num || ('quote ' + id), rejected: /reject|cancel|declin|lost/i.test(status || '') });
        };
        const model = extractJsLiteral(html, 'relatedJobsModel');
        ((model && model.RelatedQuotes && model.RelatedQuotes.Quotes) || [])
            .forEach(q => addQuote(q.QuoteNumber || q.QuoteStringId, q.Id || q.QuoteId, q.StatusDescription));
        [...doc.querySelectorAll('a[href*="/Quote/Detail/"]')].forEach(a => {
            const idm = (a.getAttribute('href') || '').match(/\/Quote\/Detail\/(\d+)/);
            addQuote(a.textContent.replace(/\s+/g, ' ').trim(), idm ? idm[1] : null, '');
        });
        return {
            quoteLabel: quotes.length ? quotes.map(q => q.num).join(', ') : null,
            hasActiveQuote: quotes.some(q => !q.rejected)
        };
    }

    // ===================================================================
    // MATCHING ENGINE (from the reconciler)
    // ===================================================================
    const NON_MATERIAL = /labour|travel|mileage|overtime|call-?out/i;
    function classifyLine(net, gross, vat, line) {
        const r = (vat != null ? vat : VAT_FALLBACK);
        const u = line.unit, sub = line.subtotal;
        if (u == null && sub == null) return null;
        if (eqMoney(u, net)) return { verdict: 'OK', why: `Unit cost £${u.toFixed(2)} = receipt net £${net.toFixed(2)}`, score: 0 };
        if (eqMoney(sub, net)) return { verdict: 'OK', why: `Line subtotal £${sub.toFixed(2)} = receipt net £${net.toFixed(2)}`, score: 0 };
        for (const t of [{ val: u, label: 'unit cost' }, { val: sub, label: 'line subtotal' }].filter(t => t.val != null)) {
            if (eqMoney(t.val, gross) || eqMoney(t.val, net * (1 + r))) {
                return { verdict: 'INCORRECT', factor: '×(1+VAT)', score: 1,
                    why: `${cap(t.label)} £${t.val.toFixed(2)} is the VAT-inclusive figure; it should be the net £${net.toFixed(2)}.` };
            }
        }
        const kNet = nearInteger(u / net);
        if (kNet) return { verdict: 'INCORRECT', factor: `×${kNet}`, score: 2,
            why: `Unit cost £${u.toFixed(2)} ≈ ${kNet} × the net £${net.toFixed(2)} (total entered as the unit cost?).` };
        const kGross = nearInteger(u / (net * (1 + r)));
        if (kGross) return { verdict: 'INCORRECT', factor: `×${kGross}×(1+VAT)`, score: 3,
            why: `Unit cost £${u.toFixed(2)} ≈ ${kGross} × the VAT-inclusive £${(net * (1 + r)).toFixed(2)}.` };
        if (eqMoney(u, net / (1 + r))) return { verdict: 'INCORRECT', factor: '÷(1+VAT)', score: 2,
            why: `Unit cost £${u.toFixed(2)} ≈ net with VAT removed again (£${(net / (1 + r)).toFixed(2)}).` };
        return null;
    }
    function matchAgainstLines(net, gross, sheetVat, lines) {
        let best = null;
        for (const line of lines) {
            const c = classifyLine(net, gross, line.vat != null ? line.vat : sheetVat, line);
            if (!c) continue;
            const cand = { ...c, line };
            const rank = (x) => (x.verdict === 'OK' ? 0 : 10) + (x.score || 0);
            if (!best || rank(cand) < rank(best)) best = cand;
        }
        return best;
    }
    function matchCombination(net, gross, vat, lines) {
        const r = (vat != null ? vat : VAT_FALLBACK);
        const mats = lines.filter(l => l.unit && l.unit > 0 && !NON_MATERIAL.test(l.category || ''));
        if (mats.length < 2) return null;
        const mk = (idxs, verdict, factor) => {
            const ls = idxs.map(i => mats[i]);
            const list = ls.map(l => `"${(l.desc || '').slice(0, 24)}" £${l.unit.toFixed(2)}`).join(' + ');
            return { verdict, factor: factor || '', combo: ls, line: { source: ls[0].source, desc: ls.map(l => l.desc).join(' + ') },
                why: verdict === 'OK' ? `${ls.length} cost lines sum to the net £${net.toFixed(2)} — ${list}`
                                      : `${ls.length} cost lines sum to the VAT-inclusive total (should be net £${net.toFixed(2)}) — ${list}` };
        };
        const test = (sum, idxs) => {
            if (eqMoney(sum, net)) return mk(idxs, 'OK');
            if (eqMoney(sum, gross) || eqMoney(sum, net * (1 + r))) return mk(idxs, 'INCORRECT', '×(1+VAT)');
            return null;
        };
        const n = mats.length;
        for (let i = 0; i < n; i++) for (let j = i + 1; j < n; j++) { const res = test(mats[i].unit + mats[j].unit, [i, j]); if (res) return res; }
        for (let i = 0; i < n; i++) for (let j = i + 1; j < n; j++) for (let k = j + 1; k < n; k++) { const res = test(mats[i].unit + mats[j].unit + mats[k].unit, [i, j, k]); if (res) return res; }
        return null;
    }
    function candidateLines(net, gross, lines, merchant) {
        const tokens = (String(merchant || '').toLowerCase().match(/[a-z]{4,}/g) || []);
        const scored = [];
        for (const l of lines) {
            if (NON_MATERIAL.test(l.category || '')) continue;
            const vals = [l.unit, l.subtotal].filter(v => v != null && v > 0);
            if (!vals.length) continue;
            let dist = Infinity;
            vals.forEach(v => [net, gross].forEach(b => { if (b > 0) dist = Math.min(dist, Math.abs(v - b) / b); }));
            const desc = String(l.desc || '').toLowerCase();
            scored.push({ line: l, dist, descMatch: tokens.some(tok => desc.includes(tok)) });
        }
        scored.sort((a, b) => a.dist - b.dist);
        return scored;
    }

    // ===================================================================
    // SHEET INPUT: .xlsx/.csv upload (SheetJS) or pasted TSV/CSV
    // ===================================================================
    function parseSheet(text) {
        const raw = text.replace(/\r\n?/g, '\n').replace(/\n+$/, '');
        if (!raw.trim()) return { headers: [], rows: [] };
        const delim = raw.split('\n')[0].includes('\t') ? '\t' : ',';
        return recordsToRows(parseDelimited(raw, delim));
    }
    function parseDelimited(text, delim) {
        const out = []; let row = [], field = '', inQ = false;
        for (let i = 0; i < text.length; i++) {
            const c = text[i];
            if (inQ) {
                if (c === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else inQ = false; }
                else field += c;
            } else {
                if (c === '"' && field === '') inQ = true;
                else if (c === delim) { row.push(field); field = ''; }
                else if (c === '\n') { row.push(field); out.push(row); row = []; field = ''; }
                else field += c;
            }
        }
        row.push(field); out.push(row);
        return out;
    }
    function recordsToRows(recs) {
        if (!recs.length) return { headers: [], rows: [] };
        const headers = recs[0].map(h => String(h == null ? '' : h).trim());
        const rows = recs.slice(1)
            .filter(r => r.some(c => String(c == null ? '' : c).trim() !== ''))
            .map(r => { const o = {}; headers.forEach((h, i) => { const v = r[i]; o[h] = typeof v === 'number' ? v : String(v == null ? '' : v).trim(); }); return o; });
        return { headers, rows };
    }
    async function readFile(file) {
        if (typeof XLSX === 'undefined') throw new Error('Excel reader (SheetJS) did not load — paste the rows instead.');
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        // raw:true keeps numbers as numbers (dates arrive as Excel serials, converted in normalizeDate)
        return recordsToRows(XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' }));
    }
    function col(row, names) {
        const keys = Object.keys(row);
        const has = (k) => k != null && row[k] !== '' && row[k] != null;
        for (const n of names) { const k = keys.find(k => norm(k) === norm(n)); if (has(k)) return row[k]; }
        for (const n of names) { const k = keys.find(k => norm(k).includes(norm(n))); if (has(k)) return row[k]; }
        return '';
    }
    const str = (v) => String(v == null ? '' : v).trim();
    const dateCell = (row) => col(row, ['Document Date', 'Date', 'Created Date', 'Settlement Date']);

    // Payhawk columns first, Pleo / reconciler names as fallbacks.
    function mapRow(row) {
        const jobRef = str(col(row, ['Job Ref(s)', 'Job Ref/Details', 'Job Ref', 'Job reference']));
        const note = [str(col(row, ['Expense Note', 'Note'])), str(col(row, ['Expense Line Note']))].filter(Boolean).join(' / ');
        const supplier = str(col(row, ['Supplier Name', 'Source description', 'Source desc', 'Merchant', 'Description']));
        const netRaw = col(row, ['Net Amount (GBP)', 'Net Amount', 'Net']);
        const grossRaw = col(row, ['Total Amount (GBP)', 'Total Amount', 'Orig. amount', 'Orig amount', 'Amount']);
        const vr = vatRate(col(row, ['Tax Rate %', 'Tax Rate', 'VAT Rate']));
        let net = money(netRaw), gross = money(grossRaw);
        const r = (vr != null ? vr : VAT_FALLBACK);
        if (net == null && gross != null) net = gross / (1 + r);
        if (gross == null && net != null) gross = net * (1 + r);
        const expenseId = str(col(row, ['Expense ID']));
        const receipt = expenseId || str(col(row, ['Receipt', 'Document Number']));
        const docNo = str(col(row, ['Document Number']));
        return {
            jobRef, note, supplier, net, gross, vat: vr, expenseId, receipt, docNo,
            refund: isNegative(netRaw) || isNegative(grossRaw),
            refs: extractRefs([jobRef, note, supplier].join('  ||  ')),
            owner: str(col(row, ['Expense Owner', 'Owner'])),
            date: normalizeDate(dateCell(row)),
            // Cost line description: supplier + note, tagged with the Payhawk expense id
            // so it can be traced (and recognised on a re-run).
            costDesc: buildDescription(supplier, note, expenseId, str(col(row, ['Xero description', 'Cost description'])))
        };
    }
    function buildDescription(supplier, note, expenseId, explicit) {
        let d = explicit || [supplier, note.replace(/^job ref\/details:\s*/i, '')].filter(Boolean).join(' - ') || 'Materials';
        const tag = expenseId ? ` [PH#${expenseId}]` : '';
        return (d.slice(0, 250 - tag.length) + tag).trim();
    }

    // ===================================================================
    // CHECK ONE ROW  ->  { status, action: 'add' | 'possible' | 'skip', ... }
    // ===================================================================
    async function checkRow(row, idx) {
        const m = mapRow(row);
        const base = { idx, receipt: m.receipt, merchant: m.supplier, note: m.note, owner: m.owner, date: m.date,
            net: m.net, gross: m.gross, description: m.costDesc, jobRef: m.jobRef, action: 'skip' };

        if (m.net == null || m.net === 0) return { ...base, status: 'NO VALUE', detail: 'No Net/Amount on this row.' };
        const blob = `${m.note} ${m.supplier} ${m.jobRef}`.toLowerCase();
        if (m.refund || /\brefund|credit note|returned/.test(blob)) return { ...base, status: 'IGNORE', detail: 'Looks like a refund / return — not a job cost.' };
        if (/personal|accidental|by mistake|wrong card/.test(blob)) return { ...base, status: 'IGNORE', detail: `Looks like a personal expense — "${m.note.slice(0, 60)}"` };

        if (!m.refs.job.length && !m.refs.sf.length && !m.refs.quote.length) {
            const internal = m.refs.internal.length ? ` (${m.refs.internal.join(', ')} is an internal code, not a job/quote)` : '';
            return { ...base, status: 'NO REFERENCE', detail: `No job ref / SF number / quote found${internal}. Job Ref: "${m.jobRef || '(blank)'}"` };
        }
        const job = await resolveJob(m.refs);
        if (!job) return { ...base, status: 'JOB NOT FOUND', detail: `Could not find a job from: ${[...m.refs.job, ...m.refs.sf, ...m.refs.quote].join(', ')}` };
        if (!job.id) return { ...base, status: 'NO JOB', detail: `${job.via} — quote not yet upgraded to a job.` };

        const { lines, parseOk } = await getJobCosts(job.id);
        if (!parseOk) return { ...base, jobId: job.id, jobNumber: job.jobNumber, status: 'ERROR', detail: 'Could not read the job\'s cost lines — not added (safety).' };
        let quoteLabel = null, hasActiveQuote = false;
        try { const q = await getRelatedQuotes(job.id); quoteLabel = q.quoteLabel; hasActiveQuote = q.hasActiveQuote; } catch (e) { /* optional */ }
        const poLines = await getUndeliveredPOLines(job.id);

        const isProject = /^PROJ/i.test(job.jobNumber || '');
        const chargeable = !(isProject || hasActiveQuote);
        const cancelled = /cancel|suspend/i.test(job.jobStatus || '');
        const jobBits = { jobId: job.id, jobNumber: job.jobNumber, jobUrl: `${location.origin}/Job/Detail/${job.id}`, via: job.via, chargeable,
            warn: cancelled ? `Job status: ${job.jobStatus}` : '' };

        // Our own tag from an earlier run of this script = definitely already there.
        if (m.expenseId) {
            const tagged = lines.find(l => (l.desc || '').includes(`[PH#${m.expenseId}]`));
            if (tagged) return { ...base, ...jobBits, status: 'ALREADY IN JOB', detail: `Line "${tagged.desc.slice(0, 50)}" already carries expense #${m.expenseId}.` };
        }

        let best = matchAgainstLines(m.net, m.gross, m.vat, lines) || matchCombination(m.net, m.gross, m.vat, lines);
        if (best && best.verdict === 'OK') {
            return { ...base, ...jobBits, status: 'ALREADY IN JOB', detail: best.why + (best.line.pending ? ' (added earlier in this sheet)' : '') };
        }
        if (best) {
            return { ...base, ...jobBits, status: 'INCORRECT', detail: `Already on the job but entered wrong: ${best.why} Not added — fix the existing line.` };
        }
        if (poLines.length) {
            const po = matchAgainstLines(m.net, m.gross, m.vat, poLines) || matchCombination(m.net, m.gross, m.vat, poLines);
            if (po) return { ...base, ...jobBits, status: 'ON UNDELIVERED PO', detail: `On ${po.line.source}: ${po.why}. Deliver the PO instead of adding a line.` };
        }

        const top = candidateLines(m.net, m.gross, lines.concat(poLines), m.supplier).slice(0, 2);
        const plausible = top.find(c => c.dist <= 0.30 || c.descMatch);
        const fmt = l => `"${(l.desc || l.category || 'line').slice(0, 40)}" £${((l.unit && l.unit > 0) ? l.unit : (l.subtotal || 0)).toFixed(2)}`;
        const chargeNote = chargeable ? 'chargeable' : `NON-chargeable (${isProject ? 'project job' : 'quoted job ' + (quoteLabel || '')})`;
        if (plausible) {
            return { ...base, ...jobBits, status: 'POSSIBLE MATCH', action: 'possible',
                detail: `No exact match for £${m.net.toFixed(2)} but a similar line exists: ${fmt(plausible.line)}. Unticked — tick to add anyway (${chargeNote}).` };
        }
        const hasMats = lines.some(l => ((l.unit && l.unit > 0) || (l.subtotal && l.subtotal > 0)) && !NON_MATERIAL.test(l.category || ''));
        return { ...base, ...jobBits, status: (hasMats || quoteLabel || poLines.length) ? 'NOT IN JOB' : 'NO COSTS', action: 'add',
            detail: `Not on ${job.jobNumber} (${job.via}). Will add £${m.net.toFixed(2)} as a ${chargeNote} material line.` };
    }

    // Record a to-be-added row in the job cache so a duplicate row later in the
    // same sheet is recognised as already in the job.
    function reservePending(res) {
        const c = jobCache.get(res.jobId);
        if (c) c.lines.push({ source: 'job', category: 'Material', pending: true, idx: res.idx, desc: res.description, unit: res.net, subtotal: res.net, qty: 1, vat: null });
    }
    function releasePending(res) {
        const c = jobCache.get(res.jobId);
        if (c) c.lines = c.lines.filter(l => !(l.pending && l.idx === res.idx));
    }

    // ===================================================================
    // WRITE (verified: POST /api/JobLine/AddMaterialCosts — from the writer)
    // ===================================================================
    async function jobMaterialDefaults(jobId) {
        if (upliftCache.has(jobId)) return upliftCache.get(jobId);
        let d = { uplift: 0, pct: 1 };
        try {
            const r = await fetch('/api/JobCost/GetAddMaterialCostMetadata?jobId=' + jobId, { credentials: 'same-origin', headers: { 'X-Requested-With': 'XMLHttpRequest' } });
            const j = await r.json(); const ad = j.AdditionalData || j;
            d = { uplift: Number(ad.Uplift) || 0, pct: ad.PriceCalculationType != null ? ad.PriceCalculationType : 1 };
        } catch (e) { /* default 0 */ }
        upliftCache.set(jobId, d);
        return d;
    }
    function buildMaterialBody(jobId, description, dateStr, net, uplift, pct, chargeable) {
        const c = CFG, cost = Number(net);
        const up = chargeable ? (Number(uplift) || 0) : 0;
        const sell = chargeable ? (cost * (1 + up / 100)).toFixed(2) : '0.00';
        const dPart = dateStr || todayDMY();
        const dt = /\d{1,2}:\d{2}/.test(dPart) ? dPart : (dPart + ' 09:00');
        return {
            JobId: jobId, TimeId: null, VirtualTimeId: null,
            CostLines: [{
                Id: null, PartNumber: null, Quantity: '1', ReturnQuantity: 0, IsReturnItemToStock: false,
                CostPerUnit: cost.toFixed(2), CostPerHour: '0.00', Uplift: up.toFixed(2), SellPerUnit: sell, SellPerHour: '0.00',
                CreateLibraryAllowed: true, CategoryId: null, CategoryDescription: null, ForEquipmentUse: false,
                Make: null, Model: null, HasFixedSell: false, SetupSell: 0,
                TaxCodeId: c.taxCodeId, TaxCodeValue: c.taxCodeValue, TaxCodeDescription: c.taxCodeDesc,
                IsChargeable: !!chargeable, PriceCalculationType: pct != null ? pct : 1,
                PayBandId: c.payBandId, PayBandDescription: c.payBandDesc, SellPayBandId: c.payBandId,
                Description: description, CreatePayBandAllowed: true, LibraryId: c.libraryId, LibraryName: c.libraryName,
                DateIncurred: dt, HasQuote: false, ItemId: 0, JobLineOption: 5,
                QuotedValueTaxCodeId: c.taxCodeId, QuotedValueTaxCodeDescription: c.taxCodeDesc,
                forEquipmentUse: false, IsIssueFromStock: false, CurrencySymbol: '£', AssignType: 0,
                RackShelfId: null, LocationId: null, Discount: '0.00', TagIds: null, Status: 'Required',
                LimitedSORAccess: false, PartSerial: null, SellingRateId: null
            }]
        };
    }
    async function addMaterialLine(it) {
        // Last-second re-check against the live job, in case someone added it since the check.
        const live = await getJobCosts(it.jobId, true);
        const dupe = live.lines.find(l => (it.receipt && (l.desc || '').includes(`[PH#${it.receipt}]`)) || eqMoney(l.unit, it.net) || eqMoney(l.subtotal, it.net));
        if (dupe && it.action === 'add') throw new Error(`SKIPPED — now already on the job as "${(dupe.desc || '').slice(0, 40)}"`);
        const d = await jobMaterialDefaults(it.jobId);
        const r = await fetch('/api/JobLine/AddMaterialCosts', {
            method: 'POST', credentials: 'same-origin',
            headers: { 'Content-Type': 'application/json', 'X-Requested-With': 'XMLHttpRequest', '__RequestVerificationToken': csrf() },
            body: JSON.stringify(buildMaterialBody(it.jobId, it.description, it.date, it.net, d.uplift, d.pct, it.chargeable))
        });
        const txt = await r.text(); let j = null; try { j = JSON.parse(txt); } catch (e) {}
        if (!r.ok) throw new Error('HTTP ' + r.status + ' ' + txt.slice(0, 120));
        if (j && j.success === false) throw new Error(j.Message || j.message || 'AddMaterialCosts returned success=false');
        // Confirm it landed
        const after = await getJobCosts(it.jobId, true);
        return after.lines.some(l => l.desc === it.description && (eqMoney(l.unit, it.net) || eqMoney(l.subtotal, it.net)));
    }

    // ===================================================================
    // MAIN: CHECK
    // ===================================================================
    async function check() {
        if (running) return;
        const parsed = loadedSheet || parseSheet(pasteArea.value);
        if (!parsed.rows.length) { alert('Upload the Excel export, or paste the rows including the header row.'); return; }
        running = true; results = []; jobCache.clear();
        inputDateMDY = detectDateOrder(parsed.rows, parsed.headers);
        checkBtn.style.display = 'none'; addBtn.style.display = 'none'; copyBtn.style.display = 'none'; stopBtn.style.display = 'inline-block';
        logArea.innerHTML = ''; resultsBox.innerHTML = '';
        log(`Read ${parsed.rows.length} rows, ${parsed.headers.length} columns. Typed dates read as ${inputDateMDY ? 'M/D/Y' : 'D/M/Y'}.`, '#0af');

        const stats = {};
        for (let i = 0; i < parsed.rows.length; i++) {
            if (!running) { log('Stopped by user.', '#f55'); break; }
            setProgress(`Checking row ${i + 1}/${parsed.rows.length}`);
            let res;
            try { res = await checkRow(parsed.rows[i], i); }
            catch (e) { res = { idx: i, status: 'ERROR', action: 'skip', detail: e.message }; }
            res.selected = res.action === 'add';
            if (res.selected) reservePending(res);
            results.push(res);
            stats[res.status] = (stats[res.status] || 0) + 1;
            renderResult(res);
            await sleep(DELAY_BETWEEN_ROWS);
        }
        log('');
        log('===== SUMMARY =====', '#0af');
        Object.keys(STATUS_COLOR).filter(s => stats[s]).forEach(s => log(`${s.padEnd(18)} ${stats[s]}`, STATUS_COLOR[s]));
        running = false;
        stopBtn.style.display = 'none'; checkBtn.style.display = 'inline-block'; copyBtn.style.display = 'inline-block';
        refreshAddButton();
        setProgress(`Checked ${results.length} rows. Review the ticks, then Add ticked to jobs.`);
    }

    // ===================================================================
    // MAIN: ADD TICKED
    // ===================================================================
    async function addTicked() {
        if (running) return;
        const items = results.filter(r => r.selected && !r.written && r.jobId && r.net != null);
        if (!items.length) return;
        const nonCharge = items.filter(i => !i.chargeable).length;
        if (!confirm(`Add ${items.length} material line(s) to live jobs now?\n\n${items.length - nonCharge} chargeable, ${nonCharge} non-chargeable (project/quoted jobs).\nEach is re-checked against the job just before adding.\n\nThis cannot be undone from here.`)) return;
        running = true;
        checkBtn.style.display = 'none'; addBtn.style.display = 'none'; stopBtn.style.display = 'inline-block';
        log(''); log('ADDING…', '#f55');
        let ok = 0, skipped = 0, fail = 0;
        for (let i = 0; i < items.length; i++) {
            if (!running) { log('Stopped by user.', '#f55'); break; }
            const it = items[i];
            setProgress(`Adding ${i + 1}/${items.length}: ${it.jobNumber}`);
            releasePending(it);
            try {
                const confirmed = await addMaterialLine(it);
                it.written = true; ok++;
                it.writeNote = confirmed ? 'Added ✓' : 'Added (could not re-read to confirm — check the job)';
                log(`  ✓ ${it.jobNumber}: £${it.net.toFixed(2)} ${it.chargeable ? 'chargeable' : 'NON-chargeable'} "${it.description.slice(0, 40)}"${confirmed ? '' : ' (unconfirmed)'}`, '#0fa');
            } catch (e) {
                it.writeNote = e.message;
                if (/^SKIPPED/.test(e.message)) { skipped++; it.selected = false; log(`  – ${it.jobNumber}: ${e.message}`, '#fb0'); }
                else { fail++; log(`  ✗ ${it.jobNumber}: ${e.message}`, '#f55'); }
            }
            renderResult(it);
            await sleep(DELAY_BETWEEN_WRITES);
        }
        log(''); log(`DONE — ${ok} added, ${skipped} skipped (already there), ${fail} failed.`, '#0af');
        setProgress(`Done: ${ok} added, ${skipped} skipped, ${fail} failed.`);
        running = false;
        stopBtn.style.display = 'none'; checkBtn.style.display = 'inline-block';
        refreshAddButton();
    }

    function copyResults() {
        const cell = v => String(v == null ? '' : v).replace(/[\t\r\n]+/g, ' ').replace(/"/g, "'").trim();
        const headers = ['Expense ID', 'Supplier', 'Note', 'Owner', 'Date', 'Net', 'Gross', 'Job Ref (sheet)', 'Status', 'Job', 'Job ID', 'Chargeable', 'Detail', 'Result'];
        const out = [headers.join('\t')];
        results.forEach(r => out.push([
            cell(r.receipt), cell(r.merchant), cell(r.note), cell(r.owner), cell(r.date),
            r.net != null ? r.net.toFixed(2) : '', r.gross != null ? r.gross.toFixed(2) : '', cell(r.jobRef),
            cell(r.status), r.jobUrl ? `=HYPERLINK("${r.jobUrl}","${cell(r.jobNumber)}")` : '', cell(r.jobId),
            r.jobId ? (r.chargeable ? 'Yes' : 'No') : '', cell(r.detail), cell(r.writeNote || '')
        ].join('\t')));
        navigator.clipboard.writeText(out.join('\n')).then(
            () => setProgress('Results copied — paste into your sheet.'),
            () => setProgress('Copy failed.')
        );
    }

    // ===================================================================
    // UI RENDER
    // ===================================================================
    const STATUS_COLOR = {
        'NOT IN JOB': '#f90', 'NO COSTS': '#b388ff', 'POSSIBLE MATCH': '#3cc', 'ALREADY IN JOB': '#0fa', 'INCORRECT': '#fb0',
        'ON UNDELIVERED PO': '#4af', 'JOB NOT FOUND': '#f55', 'NO JOB': '#f77', 'NO REFERENCE': '#999', 'IGNORE': '#888',
        'NO VALUE': '#999', 'ERROR': '#f55'
    };
    function renderResult(res) {
        const id = 'jl-eca-item-' + res.idx;
        let card = document.getElementById(id);
        if (!card) {
            card = document.createElement('div');
            card.id = id;
            card.style.cssText = 'background:#11111f;margin:6px 0;padding:8px 10px;border-radius:4px;border-left:4px solid #888;';
            resultsBox.appendChild(card);
            resultsBox.scrollTop = resultsBox.scrollHeight;
        }
        const color = res.written ? '#0f8' : (STATUS_COLOR[res.status] || '#888');
        card.style.borderLeftColor = color;
        const tickable = (res.action === 'add' || res.action === 'possible') && !res.written;
        const chg = res.jobId ? (res.chargeable ? '<span style="color:#0fa">chargeable</span>' : '<span style="color:#fb0">NON-chargeable</span>') : '';
        card.innerHTML =
            `<div style="display:flex;justify-content:space-between;gap:8px;align-items:baseline;">
               <label style="cursor:${tickable ? 'pointer' : 'default'};">
                 ${tickable ? `<input type="checkbox" ${res.selected ? 'checked' : ''} style="vertical-align:middle;margin-right:4px;">` : ''}
                 <strong style="color:${color}">${res.written ? 'ADDED ✓' : esc(res.status)}</strong>
                 <span style="color:#888"> ${esc(res.merchant || '')} · £${res.net != null ? res.net.toFixed(2) : '?'} net${res.date ? ' · ' + esc(res.date) : ''}</span>
               </label>
               <span>${res.jobUrl ? `<a href="${esc(res.jobUrl)}" target="_blank" style="color:#6cf">${esc(res.jobNumber)}</a>` : `<span style="color:#67a">${res.receipt ? '#' + esc(res.receipt) : ''}</span>`}</span>
             </div>
             <div style="color:#bbb;font-size:11px;margin-top:3px;white-space:pre-wrap;">${esc(res.detail || '')}${chg && tickable ? ' · ' + chg : ''}</div>` +
            (res.warn ? `<div style="color:#fb0;font-size:11px;margin-top:3px;">⚠ ${esc(res.warn)}</div>` : '') +
            (res.writeNote ? `<div style="color:${res.written ? '#0fa' : '#f99'};font-size:11px;margin-top:3px;">➜ ${esc(res.writeNote)}</div>` : '');
        const cb = card.querySelector('input[type=checkbox]');
        if (cb) cb.addEventListener('change', () => {
            res.selected = cb.checked;
            if (res.selected) reservePending(res); else releasePending(res);
            refreshAddButton();
        });
    }
    function refreshAddButton() {
        const n = results.filter(r => r.selected && !r.written).length;
        addBtn.textContent = `Add ${n} ticked to jobs`;
        addBtn.style.display = (n && !running) ? 'inline-block' : 'none';
    }
    function log(msg, color) {
        const line = document.createElement('div');
        line.style.cssText = `color:${color || '#ccc'};white-space:pre-wrap;word-break:break-word;`;
        line.textContent = msg;
        logArea.appendChild(line);
        logArea.scrollTop = logArea.scrollHeight;
    }
    const setProgress = (m) => { progressText.textContent = m; };

    // ===== Shared JL userscript launcher dock (identical in every script) =====
    const JL_DOCK_ID = 'jl-userscript-dock', JL_ORDER_KEY = 'jl-userscript-dock-order', JL_MIN_KEY = 'jl-userscript-dock-min', JL_TOP_KEY = 'jl-userscript-dock-top';
    const JL_BTN_CSS = 'color:#fff;padding:7px 13px;border-radius:4px;border:1px solid transparent;cursor:grab;font-family:"Open Sans",sans-serif;font-size:14px;box-shadow:0 1px 3px rgba(0,0,0,.25);white-space:nowrap;';
    const jlDockList = () => document.getElementById('jl-userscript-dock-list');
    function jlReadOrder() { try { return JSON.parse(localStorage.getItem(JL_ORDER_KEY)) || []; } catch (e) { return []; } }
    function jlSaveOrder() { const l = jlDockList(); if (!l) return; localStorage.setItem(JL_ORDER_KEY, JSON.stringify([...l.children].map(b => b.dataset.scriptId).filter(Boolean))); }
    function jlApplyOrder() { const l = jlDockList(); if (!l) return; [...l.children].sort((a, b) => { const o = jlReadOrder(); let ia = o.indexOf(a.dataset.scriptId), ib = o.indexOf(b.dataset.scriptId); if (ia < 0) ia = 1e9; if (ib < 0) ib = 1e9; return ia - ib; }).forEach(b => l.appendChild(b)); }
    function jlAfter(l, y) { let c = { o: -Infinity, el: null }; for (const el of l.querySelectorAll('button:not(.jl-dragging)')) { const r = el.getBoundingClientRect(); const off = y - (r.top + r.height / 2); if (off < 0 && off > c.o) c = { o: off, el }; } return c.el; }
    function jlSetDockMin(min) { const l = jlDockList(), t = document.getElementById('jl-userscript-dock-toggle'); if (l) l.style.display = min ? 'none' : 'flex'; if (t) t.textContent = (min ? '▸' : '▾') + ' Advanced Controls'; try { localStorage.setItem(JL_MIN_KEY, min ? '1' : '0'); } catch (e) {} }
    function jlGetDock() {
        if (!document.getElementById('jl-dock-style')) { const st = document.createElement('style'); st.id = 'jl-dock-style'; st.textContent = '#jl-userscript-dock button:hover{filter:brightness(1.18);}'; (document.head || document.documentElement).appendChild(st); }
        let d = document.getElementById(JL_DOCK_ID);
        if (!d) { d = document.createElement('div'); d.id = JL_DOCK_ID; document.body.appendChild(d); }
        d.style.cssText = 'position:fixed;top:80px;right:8px;z-index:100000;display:flex;flex-direction:column;gap:8px;align-items:flex-end;';
        const savedTop = localStorage.getItem(JL_TOP_KEY); if (savedTop !== null) d.style.top = savedTop + 'px';
        let t = document.getElementById('jl-userscript-dock-toggle');
        if (!t) {
            t = document.createElement('button');
            t.id = 'jl-userscript-dock-toggle';
            t.title = 'Drag to move up/down • click to expand/collapse';
            t.style.cssText = JL_BTN_CSS + 'background:#072d3d;border-color:#072d3d;touch-action:none;';
            let drag = null;
            t.addEventListener('pointerdown', e => { drag = { y: e.clientY, top: d.getBoundingClientRect().top, moved: false }; try { t.setPointerCapture(e.pointerId); } catch (x) {} t.style.cursor = 'grabbing'; e.preventDefault(); });
            t.addEventListener('pointermove', e => { if (!drag) return; const dy = e.clientY - drag.y; if (Math.abs(dy) > 4) drag.moved = true; if (drag.moved) { const top = Math.max(4, Math.min(window.innerHeight - 40, drag.top + dy)); d.style.top = top + 'px'; } });
            const endDrag = e => { if (!drag) return; const moved = drag.moved; drag = null; t.style.cursor = 'grab'; try { t.releasePointerCapture(e.pointerId); } catch (x) {} if (moved) { try { localStorage.setItem(JL_TOP_KEY, parseInt(d.style.top, 10)); } catch (x) {} } else { jlSetDockMin(jlDockList().style.display !== 'none'); } };
            t.addEventListener('pointerup', endDrag);
            t.addEventListener('pointercancel', endDrag);
            d.appendChild(t);
        }
        let l = document.getElementById('jl-userscript-dock-list');
        if (!l) {
            l = document.createElement('div');
            l.id = 'jl-userscript-dock-list';
            l.style.cssText = 'display:flex;flex-direction:column;gap:8px;align-items:flex-end;';
            l.addEventListener('dragover', e => { e.preventDefault(); const dr = l.querySelector('.jl-dragging'); if (!dr) return; const a = jlAfter(l, e.clientY); if (a == null) l.appendChild(dr); else l.insertBefore(dr, a); });
            l.addEventListener('drop', e => { e.preventDefault(); jlSaveOrder(); });
            d.appendChild(l);
        }
        [...d.children].forEach(c => { if (c.id && c.id.indexOf('jl-launch-') === 0) l.appendChild(c); });
        jlApplyOrder();
        jlSetDockMin(localStorage.getItem(JL_MIN_KEY) !== '0');
        return d;
    }
    function jlDockButton(id, label, color, onClick, desc) {
        jlGetDock();
        const l = jlDockList();
        let b = document.getElementById('jl-launch-' + id);
        if (b) return b;
        const bg = color || '#072d3d';
        b = document.createElement('button');
        b.id = 'jl-launch-' + id;
        b.dataset.scriptId = id;
        b.textContent = label;
        b.title = (desc ? desc + '\n\n' : '') + '(click to open • drag to reorder)';
        b.draggable = true;
        b.style.cssText = JL_BTN_CSS + 'background:' + bg + ';border-color:' + bg + ';';
        b.addEventListener('click', () => { if (b.dataset.justDragged) { delete b.dataset.justDragged; return; } onClick(); });
        b.addEventListener('dragstart', () => { b.classList.add('jl-dragging'); b.style.opacity = '0.4'; });
        b.addEventListener('dragend', () => { b.classList.remove('jl-dragging'); b.style.opacity = '1'; b.dataset.justDragged = '1'; setTimeout(() => { delete b.dataset.justDragged; }, 60); jlSaveOrder(); });
        l.appendChild(b);
        jlApplyOrder();
        return b;
    }
    function jlHelpBanner(text) {
        const b = document.createElement('div');
        b.className = 'jl-help-banner';
        b.style.cssText = 'background:#0e3a4f;color:#e3edf2;font-family:"Open Sans",sans-serif;font-size:11px;line-height:1.45;padding:8px 10px;border-radius:4px;margin:0 0 8px 0;border-left:3px solid #ff7919;';
        b.textContent = text;
        return b;
    }
    function jlRegisterPanel(panelEl, id, label, color, desc) {
        const shown = (panelEl.style.display && panelEl.style.display !== 'none') ? panelEl.style.display : 'block';
        panelEl.style.display = 'none';
        const btn = jlDockButton(id, label, color, () => {
            const opening = panelEl.style.display === 'none';
            panelEl.style.display = opening ? shown : 'none';
            if (opening && desc) {
                const box = getComputedStyle(panelEl).position === 'fixed' ? panelEl : (panelEl.firstElementChild || panelEl);
                if (box && !box.querySelector(':scope > .jl-help-banner')) box.insertBefore(jlHelpBanner(desc), box.firstChild);
            }
            btn.style.boxShadow = opening ? '0 0 0 2px #fff, 0 1px 3px rgba(0,0,0,.25)' : '0 1px 3px rgba(0,0,0,.25)';
        }, desc);
        return btn;
    }
    // ===== end shared dock =====

    function createUI() {
        if (document.getElementById('jl-expense-adder-panel')) return;

        panel = document.createElement('div');
        panel.id = 'jl-expense-adder-panel';
        const c = document.createElement('div');
        c.style.cssText = 'position:fixed;top:10px;right:10px;z-index:99999;background:#1a1a2e;color:#eee;border-radius:8px;padding:14px;width:580px;max-height:92vh;display:flex;flex-direction:column;font-family:monospace;font-size:12px;box-shadow:0 4px 20px rgba(0,0,0,.5);';

        const header = document.createElement('div');
        header.style.cssText = 'display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;';
        const title = document.createElement('strong'); title.style.fontSize = '14px'; title.textContent = 'Add expense costs to jobs  (v' + SCRIPT_VERSION + ')';
        const x = document.createElement('button'); x.textContent = '–'; x.title = 'Collapse';
        x.style.cssText = 'background:none;border:none;color:#eee;font-size:20px;cursor:pointer;line-height:1;';
        x.addEventListener('click', () => { panel.style.display = 'none'; });
        header.appendChild(title); header.appendChild(x);

        const fileRow = document.createElement('div');
        fileRow.style.cssText = 'display:flex;gap:8px;align-items:center;margin-bottom:6px;';
        fileInput = document.createElement('input');
        fileInput.type = 'file'; fileInput.accept = '.xlsx,.xls,.csv'; fileInput.style.display = 'none';
        const pick = mkBtn('📂 Choose Excel file', '#08a', () => fileInput.click());
        fileLabel = document.createElement('span'); fileLabel.style.color = '#9ab'; fileLabel.textContent = 'or paste rows (with header) below';
        fileInput.addEventListener('change', async () => {
            const f = fileInput.files[0]; if (!f) return;
            try {
                loadedSheet = await readFile(f);
                fileLabel.textContent = `${f.name} — ${loadedSheet.rows.length} rows`;
                fileLabel.style.color = '#0fa';
                pasteArea.value = ''; pasteArea.disabled = true; pasteArea.placeholder = 'Using the uploaded file. Clear it to paste instead.';
                clearFile.style.display = 'inline-block';
            } catch (e) { loadedSheet = null; fileLabel.textContent = 'Could not read file: ' + e.message; fileLabel.style.color = '#f55'; }
        });
        const clearFile = mkBtn('✕', '#555', () => {
            loadedSheet = null; fileInput.value = ''; pasteArea.disabled = false;
            pasteArea.placeholder = 'Expense ID\tCreated Date\t…\tJob Ref/Details\t…  (paste header + rows here)';
            fileLabel.textContent = 'or paste rows (with header) below'; fileLabel.style.color = '#9ab'; clearFile.style.display = 'none';
        });
        clearFile.style.display = 'none';
        fileRow.appendChild(pick); fileRow.appendChild(clearFile); fileRow.appendChild(fileLabel); fileRow.appendChild(fileInput);

        pasteArea = document.createElement('textarea');
        pasteArea.placeholder = 'Expense ID\tCreated Date\t…\tJob Ref/Details\t…  (paste header + rows here)';
        pasteArea.style.cssText = 'width:100%;height:70px;background:#0a0a1a;color:#eee;border:1px solid #555;border-radius:4px;font-family:monospace;font-size:11px;padding:6px;box-sizing:border-box;resize:vertical;';

        const controls = document.createElement('div');
        controls.style.cssText = 'margin:8px 0;display:flex;gap:8px;align-items:center;flex-wrap:wrap;';
        checkBtn = mkBtn('Check against jobs', '#0a8', check);
        addBtn = mkBtn('Add ticked to jobs', '#a22', addTicked); addBtn.style.display = 'none';
        stopBtn = mkBtn('Stop', '#a22', () => { running = false; }); stopBtn.style.display = 'none';
        copyBtn = mkBtn('Copy results', '#08a', copyResults); copyBtn.style.display = 'none';
        progressText = document.createElement('span'); progressText.style.color = '#0fa'; progressText.textContent = 'Ready.';
        [checkBtn, addBtn, stopBtn, copyBtn, progressText].forEach(el => controls.appendChild(el));

        resultsBox = document.createElement('div');
        resultsBox.style.cssText = 'overflow-y:auto;max-height:46vh;margin-bottom:6px;';
        const logLabel = document.createElement('div'); logLabel.style.cssText = 'color:#888;font-size:10px;margin-top:4px;'; logLabel.textContent = 'log';
        logArea = document.createElement('div');
        logArea.style.cssText = 'overflow-y:auto;background:#0a0a1a;padding:6px;border-radius:4px;max-height:18vh;font-size:11px;';

        [header, fileRow, pasteArea, controls, resultsBox, logLabel, logArea].forEach(el => c.appendChild(el));
        panel.appendChild(c); document.body.appendChild(panel);
        jlRegisterPanel(panel, SCRIPT_ID, SCRIPT_LABEL, SCRIPT_COLOR, SCRIPT_DESC);
    }
    function mkBtn(text, bg, fn) {
        const b = document.createElement('button');
        b.textContent = text;
        b.style.cssText = `background:${bg};color:#fff;border:none;padding:7px 13px;border-radius:4px;cursor:pointer;`;
        b.addEventListener('click', fn);
        return b;
    }

    // --- BOOT ---
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', createUI);
    else createUI();
})();
