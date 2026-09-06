// ============================================================
//  V4.4 DEPOT IMPORT ENGINE (CSV, XLSX, PDF, JSON) & EMPFEHLUNGSLISTE
// ============================================================

const LIBRARY_KEY = 'depotBibliothek_V50';
const EMPFEHLUNG_KEY_PREFIX = 'empfehlungsliste';
function empfehlungsKey(blockId) { return `${blockId}::${EMPFEHLUNG_KEY_PREFIX}`; }

function anlageschwerpunktToBlock(schwerpunkt) {
    if (!schwerpunkt) return null;
    const s = schwerpunkt.toLowerCase().trim();
    if (s.includes('geldmarkt')) return 'block-kasse';
    if (s.includes('anleihen euro kurz') || s.includes('anleihen euro kurz laufzeit') ||
        s.includes('anleihen hochzins laufzeit')) return 'block-kasse';
    if (s.includes('anleihen') || s.includes('renten') || s.includes('bond')) return 'block-defensiv';
    if (s.includes('vermögensverwalter - defensiv') || (s.includes('verm') && s.includes('defensiv'))) return 'block-defensiv';
    if (s.includes('vermögensverwalter - ausgewogen') || (s.includes('verm') && s.includes('ausgewogen'))) return 'block-ausgewogen';
    if (s.includes('vermögensverwalter - dynamisch') || (s.includes('verm') && s.includes('dynamisch'))) return 'block-dynamisch';
    if (s.includes('alternative volatilitätsstrategien') || s.includes('alternative') || s.includes('spezial')) return 'block-spezial';
    if (s.includes('aktien weit') || s.includes('weites benchmarking')) return 'block-maerkte-weit';
    if (s.includes('aktien eng') || s.includes('enges benchmarking')) return 'block-maerkte-eng';
    if (s.includes('aktien')) return 'block-maerkte-weit';
    return null;
}

function parseGermanNumber(s) {
    if (!s) return 0;
    const cleaned = String(s).replace(/\./g, '').replace(',', '.').replace(/[^0-9.]/g, '');
    const v = parseFloat(cleaned);
    return isNaN(v) ? 0 : v;
}

function normalizeWKN(wkn) {
    return String(wkn).toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function buildWknLookup() {
    const map = {};
    if (typeof managementBlocks === 'undefined') return map;
    managementBlocks.forEach(block => {
        block.funds.forEach(fund => {
            if (fund._isDelistet || fund._isEmpfehlungsliste) return;
            const matches = fund.info.matchAll(/WKN:\s*([A-Z0-9]{6})/gi);
            for (const m of matches) {
                const wkn = normalizeWKN(m[1]);
                if (!map[wkn]) map[wkn] = [];
                map[wkn].push({ block, fund });
            }
        });
    });
    const blockOrder = managementBlocks.map(b => b.id);
    Object.keys(map).forEach(wkn => {
        if (map[wkn].length <= 1) return;
        const nonTagesgeld = map[wkn].filter(e => e.block.id !== 'block-tagesgeld');
        const candidates = nonTagesgeld.length > 0 ? nonTagesgeld : map[wkn];
        candidates.sort((a, b) => blockOrder.indexOf(a.block.id) - blockOrder.indexOf(b.block.id));
        map[wkn] = [candidates[0]];
    });
    return map;
}

function parseCsvDepot(text, wknLookup) {
    const matched = [], unmatched = [];
    function cleanAmount(s) { return s.replace(/[^\d.,]/g, '').trim(); }
    const clean = text.split(/\r?\n/);

    for (let i = 0; i < clean.length; i++) {
        const cols = clean[i].split(';');
        const wknRaw = (cols[0] || '').trim();
        const wkn = normalizeWKN(wknRaw);
        if (wkn.length !== 6) continue;

        let schwerpunkt = (cols[2] || '').trim();
        let betragRaw = cleanAmount(cols[4] || cols[3] || '');
        let sparrateRaw = cleanAmount(cols[6] || cols[5] || '');
        let einmal = parseGermanNumber(betragRaw);
        let sparrate = parseGermanNumber(sparrateRaw);

        if (einmal <= 0 && sparrate <= 0 && i > 0) {
            const prevCols = clean[i - 1].split(';');
            const prevWknCheck = normalizeWKN((prevCols[0] || '').trim());
            if (prevWknCheck.length !== 6) {
                const prevBetragRaw = cleanAmount(prevCols[4] || prevCols[3] || '');
                const prevSparRaw = cleanAmount(prevCols[6] || prevCols[5] || '');
                einmal = parseGermanNumber(prevBetragRaw);
                sparrate = parseGermanNumber(prevSparRaw);
                if (!schwerpunkt) schwerpunkt = (prevCols[2] || '').trim();
            }
        }
        if (einmal <= 0 && sparrate <= 0) continue;
        const nameRaw = i + 1 < clean.length ? (clean[i + 1].split(';')[0] || '').trim() : '';

        if (wknLookup[wkn]) {
            wknLookup[wkn].forEach(entry => {
                if (!matched.some(m => m.block.id === entry.block.id && m.fund.name === entry.fund.name))
                    matched.push({ block: entry.block, fund: entry.fund, einmal, sparrate, wkn, schwerpunkt });
            });
        } else {
            if (!unmatched.some(u => u.wkn === wkn))
                unmatched.push({ name: nameRaw || `Unbekannter Fonds (${wkn})`, wkn, einmal, sparrate, schwerpunkt });
        }
    }
    return { matched, unmatched };
}

function parseXlsxDepot(workbook, wknLookup) {
    const matched = [], unmatched = [];
    if (typeof XLSX === 'undefined') return { matched, unmatched };
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

    function getAmt(row) {
        for (const idx of [4, 3]) {
            const raw = String(row[idx] === null || row[idx] === undefined ? '' : row[idx]).replace(/\u00a0/g, '').replace(/€/g, '').trim();
            const v = parseGermanNumber(raw);
            if (v > 0) return v;
        }
        return 0;
    }

    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const rawCell = row[0];
        const wkn = normalizeWKN(String(rawCell === null || rawCell === undefined ? '' : rawCell).trim());
        if (wkn.length !== 6) continue;

        let schwerpunkt = String(row[2] || '').trim();
        let einmal = getAmt(row);
        let sparrate = 0;

        if (einmal <= 0 && i > 0) {
            const prev = rows[i - 1];
            const prevWkn = normalizeWKN(String(prev[0] === null || prev[0] === undefined ? '' : prev[0]).trim());
            if (prevWkn.length !== 6) {
                einmal = getAmt(prev);
                if (!schwerpunkt) schwerpunkt = String(prev[2] || '').trim();
            }
        }
        if (einmal <= 0 && sparrate <= 0) continue;
        const nameRaw = i + 1 < rows.length ? String(rows[i + 1][0] === null || rows[i + 1][0] === undefined ? '' : rows[i + 1][0]).trim() : '';

        if (wknLookup[wkn]) {
            wknLookup[wkn].forEach(entry => {
                if (!matched.some(m => m.block.id === entry.block.id && m.fund.name === entry.fund.name))
                    matched.push({ block: entry.block, fund: entry.fund, einmal, sparrate, wkn, schwerpunkt });
            });
        } else {
            if (!unmatched.some(u => u.wkn === wkn))
                unmatched.push({ name: nameRaw || `Unbekannter Fonds (${wkn})`, wkn, einmal, sparrate, schwerpunkt });
        }
    }
    return { matched, unmatched };
}

if (typeof pdfjsLib !== 'undefined')
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

async function extractPdfText(file) {
    if (typeof pdfjsLib === 'undefined') throw new Error('PDF.js nicht geladen.');
    const pdf = await pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;
    let fullText = '';
    for (let i = 1; i <= pdf.numPages; i++) {
        const content = await (await pdf.getPage(i)).getTextContent();
        const items = content.items.sort((a,b) => {
            const ay = Math.round(a.transform[5]*10), by = Math.round(b.transform[5]*10);
            return ay !== by ? by - ay : a.transform[4] - b.transform[4];
        });
        fullText += items.map(it => it.str).join(' ') + '\n';
    }
    return fullText;
}

function parsePdfData(text, wknLookup) {
    const matched = [], unmatched = [];
    const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
    const wknPat = /\b([A-Z0-9]{6})\b/g;
    const amtPat = /(\d{1,3}(?:\.\d{3})*,\d{2})/g;
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        const wkns = []; let wm;
        wknPat.lastIndex = 0;
        while ((wm = wknPat.exec(line)) !== null) wkns.push(wm[1]);
        const ctx = [line, lines[i+1]||'', lines[i+2]||''].join(' ');
        const amts = []; let am;
        amtPat.lastIndex = 0;
        while ((am = amtPat.exec(ctx)) !== null) amts.push(parseGermanNumber(am[1]));
        for (const wknRaw of wkns) {
            const wkn = normalizeWKN(wknRaw);
            if (wkn.length !== 6) continue;
            if (wknLookup[wkn]) {
                if ((amts[0]||0) <= 0) continue;
                wknLookup[wkn].forEach(entry => {
                    if (!matched.some(m => m.block.id === entry.block.id && m.fund.name === entry.fund.name))
                        matched.push({ block: entry.block, fund: entry.fund, einmal: amts[0], sparrate: amts[1]||0, wkn, schwerpunkt: '' });
                });
            } else if ((amts[0]||0) > 0) {
                const prev = lines.slice(Math.max(0,i-3),i).filter(l => !/^\d/.test(l) && l.length > 5);
                if (!unmatched.some(u => u.wkn === wkn))
                    unmatched.push({ name: prev.at(-1) || `(${wkn})`, wkn, einmal: amts[0], sparrate: amts[1]||0, schwerpunkt: '' });
            }
        }
    }
    return { matched, unmatched };
}

function categorizeUnmatched(unmatched) {
    const mapped = [];
    const ambiguous = [];
    unmatched.forEach(u => {
        const blockId = anlageschwerpunktToBlock(u.schwerpunkt);
        if (blockId) mapped.push({ ...u, blockId });
        else ambiguous.push(u);
    });
    return { mapped, ambiguous };
}

function applyImportData(matched, mappedUnmatched) {
    portfolioGlobals.fundInvestments = {};
    portfolioGlobals.fundSparrates = {};
    portfolio = [];
    localStorage.removeItem('empfehlungslisteFonds_V44');
    localStorage.removeItem('delistedFunds_V44');

    matched.forEach(({ block, fund, einmal, sparrate }) => {
        const key = `${block.id}::${fund.name}`;
        if (einmal > 0) { portfolioGlobals.fundInvestments[key] = einmal; addFund(block, fund); }
        if (sparrate > 0) { portfolioGlobals.fundSparrates[key] = sparrate; addFund(block, fund); }
    });

    if (mappedUnmatched && mappedUnmatched.length > 0) {
        const empFondsMap = {};
        try { Object.assign(empFondsMap, JSON.parse(localStorage.getItem('empfehlungslisteFonds_V44') || '{}')); } catch {}
        mappedUnmatched.forEach(u => {
            const { blockId, einmal, sparrate } = u;
            const eKey = empfehlungsKey(blockId);
            portfolioGlobals.fundInvestments[eKey] = (portfolioGlobals.fundInvestments[eKey] || 0) + einmal;
            if (sparrate > 0) portfolioGlobals.fundSparrates[eKey] = (portfolioGlobals.fundSparrates[eKey] || 0) + sparrate;
            if (!empFondsMap[blockId]) empFondsMap[blockId] = [];
            if (!empFondsMap[blockId].some(f => f.wkn === u.wkn))
                empFondsMap[blockId].push({ name: u.name, wkn: u.wkn, schwerpunkt: u.schwerpunkt, einmal: u.einmal, sparrate: u.sparrate || 0 });
            
            const block = managementBlocks.find(b => b.id === blockId);
            if (block) {
                const layer = getOrCreateLayer(block.id, block.title);
                if (!layer.funds.some(f => f._isEmpfehlungsliste))
                    layer.funds.push({ name: EMPFEHLUNG_KEY_PREFIX, info: '', type: '', ertrag: '', _isEmpfehlungsliste: true });
            }
        });
        localStorage.setItem('empfehlungslisteFonds_V44', JSON.stringify(empFondsMap));
    }

    const totalEinmalGesamt = Object.values(portfolioGlobals.fundInvestments).reduce((s, v) => s + v, 0);
    portfolioGlobals.totalInvestment = totalEinmalGesamt;
    const totInp = document.getElementById('total-investment-input');
    if (totInp) totInp.value = formatNumberInput(totalEinmalGesamt);
    saveGlobals(); savePortfolio();
}

// ============================================================
//  V4.4 DEPOT IMPORT ENGINE (CSV, XLSX, PDF, JSON) & EMPFEHLUNGSLISTE
// ============================================================


function anlageschwerpunktToBlock(schwerpunkt) {
    if (!schwerpunkt) return null;
    const s = schwerpunkt.toLowerCase().trim();
    if (s.includes('geldmarkt')) return 'block-kasse';
    if (s.includes('anleihen euro kurz') || s.includes('anleihen euro kurz laufzeit') ||
        s.includes('anleihen hochzins laufzeit')) return 'block-kasse';
    if (s.includes('anleihen') || s.includes('renten') || s.includes('bond')) return 'block-defensiv';
    if (s.includes('vermögensverwalter - defensiv') || (s.includes('verm') && s.includes('defensiv'))) return 'block-defensiv';
    if (s.includes('vermögensverwalter - ausgewogen') || (s.includes('verm') && s.includes('ausgewogen'))) return 'block-ausgewogen';
    if (s.includes('vermögensverwalter - dynamisch') || (s.includes('verm') && s.includes('dynamisch'))) return 'block-dynamisch';
    if (s.includes('alternative volatilitätsstrategien') || s.includes('alternative') || s.includes('spezial')) return 'block-spezial';
    if (s.includes('aktien weit') || s.includes('weites benchmarking')) return 'block-maerkte-weit';
    if (s.includes('aktien eng') || s.includes('enges benchmarking')) return 'block-maerkte-eng';
    if (s.includes('aktien')) return 'block-maerkte-weit';
    return null;
}

function parseGermanNumber(s) {
    if (!s) return 0;
    const cleaned = String(s).replace(/\./g, '').replace(',', '.').replace(/[^0-9.]/g, '');
    const v = parseFloat(cleaned);
    return isNaN(v) ? 0 : v;
}

function normalizeWKN(wkn) {
    return String(wkn).toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function buildWknLookup() {
    const map = {};
    if (typeof managementBlocks === 'undefined') return map;
    managementBlocks.forEach(block => {
        block.funds.forEach(fund => {
            if (fund._isDelistet || fund._isEmpfehlungsliste) return;
            const matches = fund.info.matchAll(/WKN:\s*([A-Z0-9]{6})/gi);
            for (const m of matches) {
                const wkn = normalizeWKN(m[1]);
                if (!map[wkn]) map[wkn] = [];
                map[wkn].push({ block, fund });
            }
        });
    });
    const blockOrder = managementBlocks.map(b => b.id);
    Object.keys(map).forEach(wkn => {
        if (map[wkn].length <= 1) return;
        const nonTagesgeld = map[wkn].filter(e => e.block.id !== 'block-tagesgeld');
        const candidates = nonTagesgeld.length > 0 ? nonTagesgeld : map[wkn];
        candidates.sort((a, b) => blockOrder.indexOf(a.block.id) - blockOrder.indexOf(b.block.id));
        map[wkn] = [candidates[0]];
    });
    return map;
}

function parseCsvDepot(text, wknLookup) {
    const matched = [], unmatched = [];
    function cleanAmount(s) { return s.replace(/[^\d.,]/g, '').trim(); }
    const clean = text.split(/\r?\n/);

    for (let i = 0; i < clean.length; i++) {
        const cols = clean[i].split(';');
        const wknRaw = (cols[0] || '').trim();
        const wkn = normalizeWKN(wknRaw);
        if (wkn.length !== 6) continue;

        let schwerpunkt = (cols[2] || '').trim();
        let betragRaw = cleanAmount(cols[4] || cols[3] || '');
        let sparrateRaw = cleanAmount(cols[6] || cols[5] || '');
        let einmal = parseGermanNumber(betragRaw);
        let sparrate = parseGermanNumber(sparrateRaw);

        if (einmal <= 0 && sparrate <= 0 && i > 0) {
            const prevCols = clean[i - 1].split(';');
            const prevWknCheck = normalizeWKN((prevCols[0] || '').trim());
            if (prevWknCheck.length !== 6) {
                const prevBetragRaw = cleanAmount(prevCols[4] || prevCols[3] || '');
                const prevSparRaw = cleanAmount(prevCols[6] || prevCols[5] || '');
                einmal = parseGermanNumber(prevBetragRaw);
                sparrate = parseGermanNumber(prevSparRaw);
                if (!schwerpunkt) schwerpunkt = (prevCols[2] || '').trim();
            }
        }
        if (einmal <= 0 && sparrate <= 0) continue;
        const nameRaw = i + 1 < clean.length ? (clean[i + 1].split(';')[0] || '').trim() : '';

        if (wknLookup[wkn]) {
            wknLookup[wkn].forEach(entry => {
                if (!matched.some(m => m.block.id === entry.block.id && m.fund.name === entry.fund.name))
                    matched.push({ block: entry.block, fund: entry.fund, einmal, sparrate, wkn, schwerpunkt });
            });
        } else {
            if (!unmatched.some(u => u.wkn === wkn))
                unmatched.push({ name: nameRaw || `Unbekannter Fonds (${wkn})`, wkn, einmal, sparrate, schwerpunkt });
        }
    }
    return { matched, unmatched };
}

function parseXlsxDepot(workbook, wknLookup) {
    const matched = [], unmatched = [];
    if (typeof XLSX === 'undefined') return { matched, unmatched };
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

    function getAmt(row) {
        for (const idx of [4, 3]) {
            const raw = String(row[idx] === null || row[idx] === undefined ? '' : row[idx]).replace(/ /g, '').replace(/€/g, '').trim();
            const v = parseGermanNumber(raw);
            if (v > 0) return v;
        }
        return 0;
    }

    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const rawCell = row[0];
        const wkn = normalizeWKN(String(rawCell === null || rawCell === undefined ? '' : rawCell).trim());
        if (wkn.length !== 6) continue;

        let schwerpunkt = String(row[2] || '').trim();
        let einmal = getAmt(row);
        let sparrate = 0;

        if (einmal <= 0 && i > 0) {
            const prev = rows[i - 1];
            const prevWkn = normalizeWKN(String(prev[0] === null || prev[0] === undefined ? '' : prev[0]).trim());
            if (prevWkn.length !== 6) {
                einmal = getAmt(prev);
                if (!schwerpunkt) schwerpunkt = String(prev[2] || '').trim();
            }
        }
        if (einmal <= 0 && sparrate <= 0) continue;
        const nameRaw = i + 1 < rows.length ? String(rows[i + 1][0] === null || rows[i + 1][0] === undefined ? '' : rows[i + 1][0]).trim() : '';

        if (wknLookup[wkn]) {
            wknLookup[wkn].forEach(entry => {
                if (!matched.some(m => m.block.id === entry.block.id && m.fund.name === entry.fund.name))
                    matched.push({ block: entry.block, fund: entry.fund, einmal, sparrate, wkn, schwerpunkt });
            });
        } else {
            if (!unmatched.some(u => u.wkn === wkn))
                unmatched.push({ name: nameRaw || `Unbekannter Fonds (${wkn})`, wkn, einmal, sparrate, schwerpunkt });
        }
    }
    return { matched, unmatched };
}

if (typeof pdfjsLib !== 'undefined')
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

async function extractPdfText(file) {
    if (typeof pdfjsLib === 'undefined') throw new Error('PDF.js nicht geladen.');
    const pdf = await pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;
    let fullText = '';
    for (let i = 1; i <= pdf.numPages; i++) {
        const content = await (await pdf.getPage(i)).getTextContent();
        const items = content.items.sort((a,b) => {
            const ay = Math.round(a.transform[5]*10), by = Math.round(b.transform[5]*10);
            return ay !== by ? by - ay : a.transform[4] - b.transform[4];
        });
        fullText += items.map(it => it.str).join(' ') + '\n';
    }
    return fullText;
}

// ============================================================
//  V5.0 – Robustes Portfolio
//  - Schichten 1-6 stehen direkt unter den Zeitphasen 1-6
//  - Tagesgeld & Unternehmerisches Risiko nebeneinander losgelöst
//  - Fondsauswahl mit "+ Auswählen" / "✓ Ausgewählt" Toggle
//  - Hover-Tooltip für Top 5 Ländergewichtungen je Fonds
//  - Vollständige Ländergewichtungen für Einmalbeitrag & Sparrate
// ============================================================

const layerColors = {
    "block-kasse": "#FFB300",
    "block-defensiv": "#6E9E2E",
    "block-ausgewogen": "#8CC63F",
    "block-dynamisch": "#A8D84E",
    "block-maerkte-weit": "#5D9CEC",
    "block-maerkte-eng": "#2B4C7E",
    "block-spezial": "#00B1EB",
    "block-tagesgeld": "#FFCA28"
};

const zeitphasen = [
    { id: "phase1", name: "Zeitphase 1", duration: "< 1 Jahr", assetClass: "Kapitalreserve", mappedBlocks: ["block-kasse"], exactMatchBlock: "block-kasse" },
    { id: "phase2", name: "Zeitphase 2", duration: "> 2 Jahre", assetClass: "Defensive Vermögensverwalter", mappedBlocks: ["block-defensiv"], exactMatchBlock: "block-defensiv" },
    { id: "phase3", name: "Zeitphase 3", duration: "> 4 Jahre", assetClass: "Ausgewogene Vermögensverwalter", mappedBlocks: ["block-ausgewogen"], exactMatchBlock: "block-ausgewogen" },
    { id: "phase4", name: "Zeitphase 4", duration: "> 6 Jahre", assetClass: "Dynamische Vermögensverwalter", mappedBlocks: ["block-dynamisch"], exactMatchBlock: "block-dynamisch" },
    { id: "phase5", name: "Zeitphase 5", duration: "> 10 Jahre", assetClass: "Märkte Weites Benchmarking auf Aktien", mappedBlocks: ["block-maerkte-weit"], exactMatchBlock: "block-maerkte-weit" },
    { id: "phase6", name: "Zeitphase 6", duration: "> 10 Jahre", assetClass: "Märkte Enges Benchmarking auf Aktien", mappedBlocks: ["block-maerkte-eng"], exactMatchBlock: "block-maerkte-eng" }
];

const CLUSTER_DEFS = {
    'Nordamerika': { color: '#3b82f6', countries: new Set(['USA', 'Vereinigte Staaten', 'Kanada', 'Mexiko']) },
    'Europa': { color: '#10b981', countries: new Set(['Deutschland', 'Frankreich', 'Großbritannien', 'Niederlande', 'Schweiz', 'Österreich', 'Spanien', 'Italien', 'Schweden', 'Dänemark', 'Finnland', 'Norwegen', 'Belgien', 'Luxemburg', 'Europa', 'Sonstige Länder']) },
    'Asien': { color: '#f59e0b', countries: new Set(['Japan', 'China', 'Indien', 'Taiwan', 'Südkorea', 'Hongkong', 'Singapur', 'Asien', 'Mauritius', 'Indonesien', 'Vietnam', 'Pakistan']) },
    'Schwellenländer': { color: '#8b5cf6', countries: new Set(['Brasilien', 'Schwellenländer', 'Chile', 'Peru', 'Kolumbien', 'Südafrika', 'Ägypten', 'Rumänien', 'Kuwait', 'Kasachstan', 'Namibia']) },
    'Sonstige': { color: '#64748b', countries: new Set(['global', 'Global', 'sonstige']) }
};
const CLUSTER_ORDER = ['Nordamerika', 'Europa', 'Asien', 'Schwellenländer', 'Sonstige'];

function classifyFund(type) {
    if (!type) return 'aktien';
    const t = type.toLowerCase();
    if (t.includes('anleihen') || t.includes('fixed income') || t.includes('bond') || t.includes('geldmarkt') || t.includes('kasse') || t.includes('renten')) {
        return 'anleihen';
    }
    return 'aktien';
}

const STORAGE_KEY  = 'portfolioV50_setup';
const GLOBALS_KEY  = 'portfolioGlobalsV50_setup';

const delistetFundName = (blockTitle) => `${blockTitle} – Fonds delistet`;

function ensureDelistetFunds() {
    if (typeof managementBlocks === 'undefined') return;
    managementBlocks.forEach(block => {
        const name = delistetFundName(block.title);
        if (!block.funds.some(f => f.name === name)) {
            block.funds.push({
                name,
                info: 'Fonds aus der VEM-Liste entfernt, aber noch im Depot',
                type: 'Delistet',
                ertrag: '',
                _isDelistet: true
            });
        }
    });
}
ensureDelistetFunds();

let portfolio = loadPortfolio();

function loadPortfolio() {
    try { const r = localStorage.getItem(STORAGE_KEY); if (r) return JSON.parse(r); } catch (e) { }
    return [];
}
function savePortfolio() { localStorage.setItem(STORAGE_KEY, JSON.stringify(portfolio)); }

function loadGlobals() {
    try { const r = localStorage.getItem(GLOBALS_KEY); if (r) return JSON.parse(r); } catch (e) { }
    return {
        totalInvestment: 0,
        totalSparrate: 0,
        fundInvestments: {},
        fundSparrates: {}
    };
}
let portfolioGlobals = loadGlobals();
function saveGlobals() { localStorage.setItem(GLOBALS_KEY, JSON.stringify(portfolioGlobals)); }

window.__hardResetApp = function() {
    const keysToDelete = [];
    for (let i = 0; i < localStorage.length; i++) {
        const k = localStorage.key(i);
        if (k && (k.includes('V50') || k.includes('V44') || k.includes('portfolioGlobals') ||
                  k.includes('depotBibliothek') || k.includes('delistedFunds') ||
                  k.includes('empfehlungslisteFonds'))) {
            keysToDelete.push(k);
        }
    }
    keysToDelete.forEach(k => localStorage.removeItem(k));
    location.reload();
};

function getOrCreateLayer(blockId, blockTitle) {
    let layer = portfolio.find(l => l.blockId === blockId);
    if (!layer) {
        layer = { blockId, blockTitle, allocation: 0, funds: [] };
        portfolio.push(layer);
    }
    return layer;
}

function isFundSelected(blockId, fundName) {
    const l = portfolio.find(l => l.blockId === blockId);
    return l ? l.funds.some(f => f.name === fundName) : false;
}

function addFund(block, fund) {
    const layer = getOrCreateLayer(block.id, block.title);
    if (!isFundSelected(block.id, fund.name)) {
        const entry = { name: fund.name, info: fund.info, type: fund.type, ertrag: fund.ertrag };
        if (fund._isDelistet)         entry._isDelistet       = true;
        if (fund._isEmpfehlungsliste) entry._isEmpfehlungsliste = true;
        layer.funds.push(entry);
    }
    savePortfolio();
}

function removeFund(blockId, fundName) {
    const layer = portfolio.find(l => l.blockId === blockId);
    if (layer) {
        const idx = layer.funds.findIndex(f => f.name === fundName);
        if (idx > -1) {
            layer.funds.splice(idx, 1);
            if (layer.funds.length === 0) {
                portfolio = portfolio.filter(l => l.blockId !== blockId);
            }
            delete portfolioGlobals.fundInvestments[`${blockId}::${fundName}`];
            delete portfolioGlobals.fundSparrates[`${blockId}::${fundName}`];
            saveGlobals();
        }
    }
    savePortfolio();
}

function formatCurrency(val) {
    return new Intl.NumberFormat('de-DE', { style: 'currency', currency: 'EUR' }).format(val || 0);
}
function parseCurrencyInput(str) {
    if (!str) return 0;
    const cleanStr = String(str).replace(/\./g, '').replace(',', '.').replace(/[^0-9.]/g, '');
    return parseFloat(cleanStr) || 0;
}
function formatNumberInput(num) {
    if (!num && num !== 0) return '';
    return new Intl.NumberFormat('de-DE', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(num);
}

document.addEventListener('DOMContentLoaded', () => {

    const hardResetBtn = document.getElementById('hard-reset-btn');
    if (hardResetBtn) {
        hardResetBtn.addEventListener('click', () => {
            const overlay = document.createElement('div');
            overlay.style.cssText = `
                position:fixed; inset:0; background:rgba(0,0,0,0.7);
                display:flex; align-items:center; justify-content:center; z-index:999999;`;
            overlay.innerHTML = `
                <div style="background:#1a1a2e; border:2px solid #dc2626; border-radius:14px;
                     padding:36px 40px; min-width:340px; max-width:420px; text-align:center;
                     box-shadow:0 8px 40px rgba(220,38,38,0.4);">
                    <div style="font-size:2.5em; margin-bottom:12px;">⚠️</div>
                    <div style="color:#fff; font-size:1.15em; font-weight:700; margin-bottom:8px;">Hard Reset V5.0</div>
                    <div style="color:#fca5a5; font-size:0.9em; margin-bottom:16px;">Alle gespeicherten V5.0 Daten werden gelöscht.</div>
                    <div style="display:flex; gap:12px; justify-content:center;">
                        <button id="hr-cancel" style="padding:10px 24px; border-radius:8px; border:1px solid #4b5563; background:transparent; color:#d1d5db; cursor:pointer;">Abbrechen</button>
                        <button id="hr-confirm" style="padding:10px 28px; border-radius:8px; border:none; background:#dc2626; color:#fff; cursor:pointer; font-weight:700;">Ja, alles löschen</button>
                    </div>
                </div>`;
            document.body.appendChild(overlay);
            overlay.querySelector('#hr-cancel').addEventListener('click', () => overlay.remove());
            overlay.querySelector('#hr-confirm').addEventListener('click', () => {
                overlay.remove();
                window.__hardResetApp();
            });
            overlay.addEventListener('click', e => { if (e.target === overlay) overlay.remove(); });
        });
    }

    const gridContainer = document.getElementById('v5-grid-container');
    const tagesgeldContainer = document.getElementById('block-tagesgeld-container');
    const spezialContainer = document.getElementById('block-spezial-container');

    const modal = document.getElementById('funds-modal');
    const modalTitle = document.getElementById('modal-title');
    const modalList = document.getElementById('modal-funds-list');
    const modalFilter = document.getElementById('modal-filter');
    const closeBtn = document.querySelector('#funds-modal .close-modal');

    const totalInvestmentInput = document.getElementById('total-investment-input');
    const totalDistributedDisplay = document.getElementById('total-distributed-display');
    const totalRemainingDisplay = document.getElementById('total-remaining-display');
    const totalInvestmentSparrate = document.getElementById('total-investment-sparrate');
    const sparrateDistributedDisplay = document.getElementById('sparrate-distributed-display');
    const sparrateRemainingDisplay = document.getElementById('sparrate-remaining-display');

    const panelEmpty = document.getElementById('panel-empty');
    const panelLayers = document.getElementById('panel-layers');
    const resetBtn = document.getElementById('reset-portfolio-btn');
    const savePortfolioBtn = document.getElementById('save-portfolio-btn');

    let currentBlockData = null;
    let _searchTargetFundName = null;
    let activePhaseId = null;

    // ── PROMINENT SEARCH FUNCTIONALITY ────────────────────────
    const searchInput = document.getElementById('fund-search-input');
    const searchResults = document.getElementById('fund-search-results');

    if (searchInput && searchResults) {
        searchInput.addEventListener('input', (e) => {
            const query = e.target.value.toLowerCase().trim();
            searchResults.innerHTML = '';
            if (query.length < 2) { searchResults.style.display = 'none'; return; }

            let matches = [];
            managementBlocks.forEach(block => {
                block.funds.forEach(fund => {
                    const searchStr = `${fund.name} ${fund.info} ${fund.type}`.toLowerCase();
                    if (searchStr.includes(query)) matches.push({ fund, block });
                });
            });

            if (matches.length === 0) {
                searchResults.innerHTML = '<li><span class="fund-search-meta">Keine Fonds gefunden.</span></li>';
                searchResults.style.display = 'block';
                return;
            }

            matches.forEach(match => {
                const li = document.createElement('li');
                li.innerHTML = `
                    <span class="fund-search-name">${match.fund.name}</span>
                    <span class="fund-search-meta">${match.fund.info} | Schicht: ${match.block.title}</span>`;
                li.addEventListener('click', () => {
                    searchInput.value = ''; searchResults.style.display = 'none';
                    _searchTargetFundName = match.fund.name;
                    openFundModalForBlock(match.block);
                });
                searchResults.appendChild(li);
            });
            searchResults.style.display = 'block';
        });

        document.addEventListener('click', e => {
            if (!searchInput.contains(e.target) && !searchResults.contains(e.target)) {
                searchResults.style.display = 'none';
            }
        });
    }

    // ── RENDER 6 PHASE COLUMNS ───────────────────────────────
    function renderGridColumns() {
        if (!gridContainer) return;
        gridContainer.innerHTML = '';

        zeitphasen.forEach(phase => {
            const col = document.createElement('div');
            col.className = 'v5-phase-col';

            const header = document.createElement('div');
            header.className = `v5-phase-header ${activePhaseId === phase.id ? 'active' : ''}`;
            header.dataset.phaseId = phase.id;
            header.innerHTML = `
                <div>${phase.name}</div>
                <span class="duration-badge">${phase.duration}</span>`;
            header.addEventListener('click', () => handlePhaseClick(phase));

            const blockId = phase.exactMatchBlock;
            const block = managementBlocks.find(b => b.id === blockId);
            const cylinder = createCylinderElement(block);

            col.appendChild(header);
            col.appendChild(cylinder);
            gridContainer.appendChild(col);
        });
    }

    // ── RENDER STANDALONE CYLINDERS (NEBENEINANDER OHNE ÜBERSCHRIFT) ──
    function renderStandaloneCylinders() {
        if (tagesgeldContainer) {
            tagesgeldContainer.innerHTML = '';
            const bTagesgeld = managementBlocks.find(b => b.id === 'block-tagesgeld');
            if (bTagesgeld) tagesgeldContainer.appendChild(createCylinderElement(bTagesgeld));
        }

        if (spezialContainer) {
            spezialContainer.innerHTML = '';
            const bSpezial = managementBlocks.find(b => b.id === 'block-spezial');
            if (bSpezial) spezialContainer.appendChild(createCylinderElement(bSpezial));
        }
    }

    function createCylinderElement(block) {
        const blockEl = document.createElement('div');
        blockEl.className = 'tower-layer highlighted';
        blockEl.id = block.id;
        blockEl.dataset.blockId = block.id;

        const bg = layerColors[block.id] || '#2563eb';
        blockEl.style.backgroundColor = bg;

        const layerData = portfolio.find(l => l.blockId === block.id);
        const count = layerData ? layerData.funds.length : 0;

        let sumEinmal = 0;
        let sumSpar = 0;
        if (layerData) {
            layerData.funds.forEach(f => {
                const k = `${block.id}::${f.name}`;
                sumEinmal += (portfolioGlobals.fundInvestments[k] || 0);
                sumSpar += (portfolioGlobals.fundSparrates[k] || 0);
            });
        }

        blockEl.innerHTML = `
            <div class="layer-content">
                <div class="block-title">${block.title}</div>
            </div>
            <div class="layer-selection-count ${count > 0 ? 'visible' : ''}">
                <span class="count-val">${count}</span> Fonds
            </div>
            <div class="layer-assigned-amount ${(sumEinmal > 0 || sumSpar > 0) ? 'visible' : ''}">
                ${sumEinmal > 0 ? `<span>${formatCurrency(sumEinmal)}</span>` : ''}
                ${sumSpar > 0 ? `<span style="font-size:0.75rem; color:#059669;">${formatCurrency(sumSpar)}/mtl.</span>` : ''}
            </div>`;

        blockEl.addEventListener('click', () => openFundModalForBlock(block));
        return blockEl;
    }

    function handlePhaseClick(phase) {
        if (activePhaseId === phase.id) {
            activePhaseId = null;
            resetHighlights();
            return;
        }
        activePhaseId = phase.id;

        document.querySelectorAll('.v5-phase-header').forEach(h => h.classList.remove('active'));
        const activeHeader = document.querySelector(`.v5-phase-header[data-phase-id="${phase.id}"]`);
        if (activeHeader) activeHeader.classList.add('active');

        document.querySelectorAll('.tower-layer').forEach(card => card.classList.remove('highlighted', 'exact-match'));
        
        phase.mappedBlocks.forEach(bId => {
            const card = document.getElementById(bId);
            if (card) {
                card.classList.add('highlighted', 'exact-match');
                const bgColor = window.getComputedStyle(card).backgroundColor;
                card.style.setProperty('--badge-color', bgColor);
            }
        });
    }

    function resetHighlights() {
        document.querySelectorAll('.v5-phase-header').forEach(h => h.classList.remove('active'));
        document.querySelectorAll('.tower-layer').forEach(card => {
            card.classList.add('highlighted');
            card.classList.remove('exact-match');
        });
    }

    // ── MODAL OPEN & FUND LIST WITH "+ AUSWÄHLEN" / "✓ AUSGEWÄHLT" & TOP 5 COUNTRY HOVER TOOLTIP ──
    function openFundModalForBlock(block) {
        currentBlockData = block;
        modalTitle.textContent = block.title;
        modalFilter.value = 'all';
        populateFilterDropdown(block.funds, block.id);
        renderFundList(block.funds);
        modal.classList.add('open');

        if (_searchTargetFundName) {
            const targetName = _searchTargetFundName;
            _searchTargetFundName = null;
            setTimeout(() => {
                const allItems = modalList.querySelectorAll('.fund-list-item');
                allItems.forEach(item => {
                    if (item.dataset.fundName === targetName) {
                        item.scrollIntoView({ behavior: 'smooth', block: 'center' });
                        item.classList.add('search-highlight-flash');
                        setTimeout(() => item.classList.remove('search-highlight-flash'), 2500);
                    }
                });
            }, 200);
        }
    }

    function populateFilterDropdown(funds, blockId) {
        modalFilter.innerHTML = '<option value="all">Alle Arten anzeigen (Filter)</option>';
        const types = [...new Set(funds.map(f => f.type).filter(Boolean))];
        types.forEach(t => {
            const opt = document.createElement('option');
            opt.value = t; opt.textContent = t;
            modalFilter.appendChild(opt);
        });
    }

    const refreshModalButtons = () => {
        if (!currentBlockData) return;
        modalList.querySelectorAll('.fund-list-item').forEach(li => {
            const sel = isFundSelected(currentBlockData.id, li.dataset.fundName);
            const btn = li.querySelector('.btn-fund-select');
            if (!btn) return;
            li.classList.toggle('is-selected', sel);
            btn.classList.toggle('selected', sel);
            btn.textContent = sel ? '✓ Ausgewählt' : '+ Auswählen';
        });
    };

    function renderFundList(funds) {
        modalList.innerHTML = '';
        const filterVal = modalFilter ? modalFilter.value : 'all';
        const filtered = filterVal === 'all' ? funds : funds.filter(f => f.type === filterVal);

        if (!filtered || filtered.length === 0) {
            modalList.innerHTML = '<li class="fund-list-item"><span class="fund-name">Keine Fonds für diesen Filter.</span></li>';
            return;
        }

        filtered.forEach(fund => {
            const sel = isFundSelected(currentBlockData.id, fund.name);
            const key = `${currentBlockData.id}::${fund.name}`;
            const valEinmal = portfolioGlobals.fundInvestments[key] || 0;
            const valSpar = portfolioGlobals.fundSparrates[key] || 0;

            // ── TOP 5 COUNTRY HOVER TOOLTIP GENERATOR (EXACT V4.4) ──
            let countryTooltipHtml = '';
            if (fund.countryWeightings && fund.countryWeightings.length > 0) {
                const fClass = classifyFund(fund.type);
                const sortedCW = [...fund.countryWeightings].sort((a, b) => b.weight - a.weight).slice(0, 5);

                const listHtml = sortedCW.map(c => {
                    let subLine = '';
                    const formattedWeight = typeof c.weight === 'number' ? c.weight.toFixed(1).replace('.', ',') : c.weight;
                    if (fClass === 'aktien') {
                        subLine = `<div style="font-size:0.8em;color:#93c5fd;padding-left:4px;">&#x2514; Aktien ${formattedWeight}%</div>`;
                    } else if (fClass === 'anleihen') {
                        subLine = `<div style="font-size:0.8em;color:#86efac;padding-left:4px;">&#x2514; Anleihen ${formattedWeight}%</div>`;
                    }
                    return `<li style="margin-bottom:5px;"><div style="display:flex;justify-content:space-between;gap:10px;"><span>${c.country}</span><strong>${c.weight}%</strong></div>${subLine}</li>`;
                }).join('');

                const classLabel = fClass === 'aktien' ? '<span style="font-size:0.75em;background:#1e3a8a;color:#93c5fd;border-radius:4px;padding:1px 6px;margin-left:6px;">Aktien</span>'
                                 : fClass === 'anleihen' ? '<span style="font-size:0.75em;background:#14532d;color:#86efac;border-radius:4px;padding:1px 6px;margin-left:6px;">Anleihen</span>'
                                 : '<span style="font-size:0.75em;background:#334155;color:#cbd5e1;border-radius:4px;padding:1px 6px;margin-left:6px;">Gemischt</span>';

                countryTooltipHtml = `
                    <div class="fund-country-tooltip">
                        <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" class="country-icon">
                            <circle cx="12" cy="12" r="10"></circle>
                            <line x1="2" y1="12" x2="22" y2="12"></line>
                            <path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10z"></path>
                        </svg>
                        <div class="tooltip-content">
                            <strong style="font-size:1.05em; border-bottom:1px solid #555; display:flex; align-items:center; padding-bottom:5px; margin-bottom:5px;">Top 5 Länder ${classLabel}</strong>
                            <ul style="margin:0;padding:0;list-style:none;">${listHtml}</ul>
                        </div>
                    </div>`;
            }

            const ertragCls = fund.ertrag === 'thesaurierend' ? 'thesaurierend' : fund.ertrag === 'ausschüttend' ? 'ausschuettend' : '';
            const ertragHtml = fund.ertrag ? `<span class="fund-badge-ertrag ${ertragCls}">${fund.ertrag}</span>` : '';
            const isDelistet = !!fund._isDelistet;

            const item = document.createElement('li');
            item.className = `fund-list-item${sel ? ' is-selected' : ''}${isDelistet ? ' fund-delistet' : ''}`;
            item.dataset.fundName = fund.name;

            const delistetBadgeHtml = isDelistet
                ? `<span class="delisted-badge">Fonds delistet</span>`
                : '';

            const fundInput = `
                <div class="fund-input-wrapper">
                    <div class="fund-input-col">
                        <label class="fund-input-label">Einmalbeitrag</label>
                        <div class="currency-input-wrapper">
                            <input type="text" class="fund-investment-input" data-fund-key="${key}"
                                placeholder="0,00" inputmode="decimal"
                                value="${valEinmal > 0 ? formatNumberInput(valEinmal) : ''}">
                            <span class="currency-symbol">€</span>
                        </div>
                    </div>
                    <div class="fund-input-col">
                        <label class="fund-input-label">Sparrate mtl.</label>
                        <div class="currency-input-wrapper">
                            <input type="text" class="fund-sparrate-input" data-fund-key="${key}"
                                placeholder="0,00" inputmode="decimal"
                                value="${valSpar > 0 ? formatNumberInput(valSpar) : ''}">
                            <span class="currency-symbol">€</span>
                        </div>
                    </div>
                </div>`;

            item.innerHTML = `
                <div class="fund-info-wrapper">
                    <div class="fund-header" style="display:flex; align-items:center; gap:8px;">
                        <span class="fund-name">${fund.name}</span>
                        ${delistetBadgeHtml}
                        ${countryTooltipHtml}
                    </div>
                    <div style="margin-top:4px;display:flex;gap:8px;flex-wrap:wrap;align-items:center;">
                        ${!isDelistet && fund.type ? `<span class="fund-badge">${fund.type}</span>` : ''}
                        ${ertragHtml}
                    </div>
                    <span class="fund-info" style="margin-top:6px;display:block;">${fund.info}</span>
                </div>
                ${fundInput}
                <button class="btn-fund-select${sel ? ' selected' : ''}"
                    title="${sel ? 'Klicken zum Entfernen' : 'Zum Portfolio hinzufügen'}">
                    ${sel ? '✓ Ausgewählt' : '+ Auswählen'}
                </button>`;

            // Toggle select button
            const selBtn = item.querySelector('.btn-fund-select');
            selBtn.addEventListener('click', (e) => {
                e.stopPropagation();
                if (isFundSelected(currentBlockData.id, fund.name)) {
                    removeFund(currentBlockData.id, fund.name);
                } else {
                    addFund(currentBlockData, fund);
                }
                refreshModalButtons();
                updatePortfolioUI();
            });

            // EUR Input listeners for Einmalbeitrag
            const inpEinmal = item.querySelector('.fund-investment-input');
            if (inpEinmal) {
                inpEinmal.addEventListener('input', e => {
                    const v = parseCurrencyInput(e.target.value);
                    if (v > 0) {
                        portfolioGlobals.fundInvestments[key] = v;
                        if (!isFundSelected(currentBlockData.id, fund.name)) {
                            addFund(currentBlockData, fund);
                            refreshModalButtons();
                        }
                    } else {
                        delete portfolioGlobals.fundInvestments[key];
                    }
                    saveGlobals();
                    updatePortfolioUI();
                });
                inpEinmal.addEventListener('blur', e => {
                    const v = parseCurrencyInput(e.target.value);
                    e.target.value = v > 0 ? formatNumberInput(v) : '';
                });
                inpEinmal.addEventListener('focus', e => {
                    const v = portfolioGlobals.fundInvestments[key] || 0;
                    e.target.value = v > 0 ? v.toString().replace('.', ',') : '';
                });
                inpEinmal.addEventListener('keydown', e => { if (e.key === 'Enter') e.target.blur(); });
            }

            // EUR Input listeners for Sparrate
            const inpSpar = item.querySelector('.fund-sparrate-input');
            if (inpSpar) {
                inpSpar.addEventListener('input', e => {
                    const v = parseCurrencyInput(e.target.value);
                    if (v > 0) {
                        portfolioGlobals.fundSparrates[key] = v;
                        if (!isFundSelected(currentBlockData.id, fund.name)) {
                            addFund(currentBlockData, fund);
                            refreshModalButtons();
                        }
                    } else {
                        delete portfolioGlobals.fundSparrates[key];
                    }
                    saveGlobals();
                    updatePortfolioUI();
                });
                inpSpar.addEventListener('blur', e => {
                    const v = parseCurrencyInput(e.target.value);
                    e.target.value = v > 0 ? formatNumberInput(v) : '';
                });
                inpSpar.addEventListener('focus', e => {
                    const v = portfolioGlobals.fundSparrates[key] || 0;
                    e.target.value = v > 0 ? v.toString().replace('.', ',') : '';
                });
                inpSpar.addEventListener('keydown', e => { if (e.key === 'Enter') e.target.blur(); });
            }

            // Hover Tooltip Positioner JS
            const tooltipWrapper = item.querySelector('.fund-country-tooltip');
            if (tooltipWrapper) {
                const tooltipContent = tooltipWrapper.querySelector('.tooltip-content');
                tooltipWrapper.addEventListener('mouseenter', () => {
                    const iconRect = tooltipWrapper.getBoundingClientRect();
                    const modalHeaderEl = document.querySelector('#funds-modal .modal-header');
                    const headerBottom = modalHeaderEl ? modalHeaderEl.getBoundingClientRect().bottom : 80;
                    const tooltipH = 175;
                    const spaceAbove = iconRect.top - headerBottom;

                    if (spaceAbove >= tooltipH + 10) {
                        tooltipContent.classList.remove('tip-below');
                        tooltipContent.classList.add('tip-above');
                        tooltipContent.style.top = (iconRect.top - tooltipH - 8) + 'px';
                    } else {
                        tooltipContent.classList.remove('tip-above');
                        tooltipContent.classList.add('tip-below');
                        tooltipContent.style.top = (iconRect.bottom + 8) + 'px';
                    }
                    const left = Math.max(4, iconRect.right - 210);
                    tooltipContent.style.left = left + 'px';
                });
            }

            modalList.appendChild(item);
        });
    }

    if (modalFilter) {
        modalFilter.addEventListener('change', () => {
            if (currentBlockData) renderFundList(currentBlockData.funds);
        });
    }

    const closeModal = () => {
        modal.classList.remove('open');
        resetHighlights();
    };

    if (closeBtn) closeBtn.addEventListener('click', closeModal);
    window.addEventListener('click', e => { if (e.target === modal) closeModal(); });

    // ── FINANCIAL INPUT LISTENERS ──────────────────────────────
    if (totalInvestmentInput) {
        totalInvestmentInput.value = portfolioGlobals.totalInvestment > 0 ? formatNumberInput(portfolioGlobals.totalInvestment) : '';
        totalInvestmentInput.addEventListener('change', (e) => {
            portfolioGlobals.totalInvestment = parseCurrencyInput(e.target.value);
            saveGlobals();
            updatePortfolioUI();
        });
    }

    if (totalInvestmentSparrate) {
        totalInvestmentSparrate.value = portfolioGlobals.totalSparrate > 0 ? formatNumberInput(portfolioGlobals.totalSparrate) : '';
        totalInvestmentSparrate.addEventListener('change', (e) => {
            portfolioGlobals.totalSparrate = parseCurrencyInput(e.target.value);
            saveGlobals();
            updatePortfolioUI();
        });
    }

    // ── LÄNDERGEWICHTUNGEN & SCHICHT-BREAKDOWN ENGINE ──────
    function updateCountryBreakdown() {
        const infoBar = document.getElementById('portfolio-info-bar');
        const summaryEinmal = document.getElementById('portfolio-country-summary');
        const summarySpar = document.getElementById('portfolio-country-summary-sparrate');
        const clusterListEinmal = document.getElementById('portfolio-cluster-list');
        const clusterListSpar = document.getElementById('portfolio-cluster-list-sparrate');
        const layerBreakdownList = document.getElementById('layer-breakdown-list');

        if (!infoBar) return;

        let totalE = 0;
        let totalS = 0;
        portfolio.forEach(l => {
            l.funds.forEach(f => {
                const k = `${l.blockId}::${f.name}`;
                totalE += (portfolioGlobals.fundInvestments[k] || 0);
                totalS += (portfolioGlobals.fundSparrates[k] || 0);
            });
        });

        if (totalE === 0 && totalS === 0) {
            infoBar.style.display = 'none';
            return;
        }

        infoBar.style.display = 'grid';

        // 1. Layer Breakdown
        if (layerBreakdownList) {
            layerBreakdownList.innerHTML = '';
            portfolio.forEach(l => {
                const b = managementBlocks.find(b => b.id === l.blockId);
                if (!b) return;
                let lE = 0; let lS = 0;
                l.funds.forEach(f => {
                    const k = `${l.blockId}::${f.name}`;
                    lE += (portfolioGlobals.fundInvestments[k] || 0);
                    lS += (portfolioGlobals.fundSparrates[k] || 0);
                });
                if (lE > 0 || lS > 0) {
                    const color = layerColors[l.blockId] || '#2563eb';
                    const li = document.createElement('li');
                    li.style.cssText = `display:flex; justify-content:space-between; margin-bottom:8px; padding:6px 10px; background:#f8fafc; border-left:4px solid ${color}; border-radius:4px; font-size:13px; font-weight:600;`;
                    li.innerHTML = `<span>${b.title}</span> <span>${lE > 0 ? formatCurrency(lE) : ''} ${lS > 0 ? `(${formatCurrency(lS)} mtl.)` : ''}</span>`;
                    layerBreakdownList.appendChild(li);
                }
            });
        }

        // 2. Compute Clusters for Einmal & Sparrate
        const computeData = (isSpar) => {
            let countryVals = {};
            let countryValsAktien = {};
            let countryValsAnleihen = {};
            let sum = 0;

            portfolio.forEach(layer => {
                layer.funds.forEach(fund => {
                    const k = `${layer.blockId}::${fund.name}`;
                    const val = isSpar ? (portfolioGlobals.fundSparrates[k] || 0) : (portfolioGlobals.fundInvestments[k] || 0);
                    if (val > 0) {
                        sum += val;
                        const blockData = managementBlocks.find(b => b.id === layer.blockId);
                        if (blockData) {
                            const fundData = blockData.funds.find(f => f.name === fund.name);
                            if (fundData && fundData.countryWeightings) {
                                const fClass = classifyFund(fundData.type);
                                fundData.countryWeightings.forEach(cw => {
                                    const v = val * (cw.weight / 100);
                                    const c = cw.country;
                                    countryVals[c] = (countryVals[c] || 0) + v;
                                    if (fClass === 'aktien') countryValsAktien[c] = (countryValsAktien[c] || 0) + v;
                                    else countryValsAnleihen[c] = (countryValsAnleihen[c] || 0) + v;
                                });
                            }
                        }
                    }
                });
            });
            return { countryVals, countryValsAktien, countryValsAnleihen, sum };
        };

        const renderClusterPills = (containerEl, listEl, data, label) => {
            const { countryVals, sum } = data;
            if (sum <= 0 || Object.keys(countryVals).length === 0) {
                listEl.innerHTML = `<div style="color:#888; font-size:13px; text-align:center; padding:10px 0;">Keine ${label} verplant</div>`;
                return;
            }

            const clusterVals = {};
            const clusterDetails = {};
            CLUSTER_ORDER.forEach(k => { clusterVals[k] = 0; clusterDetails[k] = {}; });

            Object.keys(countryVals).forEach(country => {
                const val = countryVals[country];
                let assigned = false;
                for (const key of CLUSTER_ORDER) {
                    if (CLUSTER_DEFS[key].countries.has(country.trim())) {
                        clusterVals[key] += val;
                        clusterDetails[key][country] = (clusterDetails[key][country] || 0) + val;
                        assigned = true;
                        break;
                    }
                }
                if (!assigned) {
                    clusterVals['Sonstige'] += val;
                    clusterDetails['Sonstige'][country] = (clusterDetails['Sonstige'][country] || 0) + val;
                }
            });

            listEl.innerHTML = '';
            const fmt = n => n.toFixed(1).replace('.', ',') + '%';

            CLUSTER_ORDER.forEach(cKey => {
                const pVal = clusterVals[cKey];
                if (pVal <= 0) return;
                const pct = (pVal / sum) * 100;
                const color = CLUSTER_DEFS[cKey].color;

                const topCountries = Object.entries(clusterDetails[cKey])
                    .map(([c, v]) => ({ c, p: (v / sum) * 100 }))
                    .sort((a, b) => b.p - a.p).slice(0, 5);

                const pill = document.createElement('div');
                pill.className = 'cluster-pill-item';
                pill.style.borderLeft = `4px solid ${color}`;
                pill.innerHTML = `
                    <div class="cluster-pill-header">
                        <span>${cKey}</span>
                        <span>${fmt(pct)}</span>
                    </div>
                    <ul class="cluster-top-countries">
                        ${topCountries.map(d => `<li style="display:flex; justify-content:space-between; margin-top:2px;"><span>${d.c}</span> <span>${fmt(d.p)}</span></li>`).join('')}
                    </ul>`;
                listEl.appendChild(pill);
            });
        };

        if (summaryEinmal && clusterListEinmal) {
            renderClusterPills(summaryEinmal, clusterListEinmal, computeData(false), 'Einmalbeiträge');
        }
        if (summarySpar && clusterListSpar) {
            renderClusterPills(summarySpar, clusterListSpar, computeData(true), 'Sparraten');
        }
    }

    function updatePortfolioUI() {
        renderGridColumns();
        renderStandaloneCylinders();
        updateCountryBreakdown();

        let sumEinmal = 0;
        let sumSpar = 0;

        Object.values(portfolioGlobals.fundInvestments).forEach(v => sumEinmal += (Number(v) || 0));
        Object.values(portfolioGlobals.fundSparrates).forEach(v => sumSpar += (Number(v) || 0));

        if (totalDistributedDisplay) totalDistributedDisplay.textContent = formatCurrency(sumEinmal);
        if (totalRemainingDisplay) {
            const remE = (portfolioGlobals.totalInvestment || 0) - sumEinmal;
            totalRemainingDisplay.textContent = formatCurrency(remE);
            totalRemainingDisplay.style.color = remE < 0 ? '#ef4444' : '#10b981';
        }

        if (sparrateDistributedDisplay) sparrateDistributedDisplay.textContent = formatCurrency(sumSpar) + ' mtl.';
        if (sparrateRemainingDisplay) {
            const remS = (portfolioGlobals.totalSparrate || 0) - sumSpar;
            sparrateRemainingDisplay.textContent = formatCurrency(remS) + ' mtl.';
            sparrateRemainingDisplay.style.color = remS < 0 ? '#ef4444' : '#10b981';
        }

        // Update Allocation Summary Progress Bar (V4.4)
        const allocFill = document.getElementById('alloc-bar-fill');
        const allocTotal = document.getElementById('alloc-total');
        if (allocFill && allocTotal) {
            const totInv = portfolioGlobals.totalInvestment || 0;
            const pct = totInv > 0 ? (sumEinmal / totInv) * 100 : 0;
            allocTotal.textContent = pct.toFixed(1).replace('.', ',');
            allocFill.style.width = Math.min(pct, 100) + '%';
            if (Math.abs(pct - 100) < 0.1) {
                allocFill.style.backgroundColor = '#10b981';
            } else if (pct > 100) {
                allocFill.style.backgroundColor = '#ef4444';
            } else {
                allocFill.style.backgroundColor = '#3b82f6';
            }
        }

        if (panelLayers) {
            if (portfolio.length === 0) {
                if (panelEmpty) panelEmpty.style.display = 'block';
                panelLayers.style.display = 'none';
            } else {
                if (panelEmpty) panelEmpty.style.display = 'none';
                panelLayers.style.display = 'block';
                renderRightPanelLayers();
            }
        }
    }

    function renderRightPanelLayers() {
        if (!panelLayers) return;
        panelLayers.innerHTML = '';

        portfolio.forEach(layer => {
            const block = managementBlocks.find(b => b.id === layer.blockId);
            if (!block) return;

            const div = document.createElement('div');
            div.className = 'panel-layer-card';
            const color = layerColors[block.id] || '#2563eb';
            div.style.borderLeft = `4px solid ${color}`;

            let layerEinmal = 0;
            let layerSpar = 0;

            const layerEl = document.createElement('div');

            layer.funds.forEach(f => {
                if (f._isEmpfehlungsliste) return;
                const k = `${layer.blockId}::${f.name}`;
                const eVal = portfolioGlobals.fundInvestments[k] || 0;
                const sVal = portfolioGlobals.fundSparrates[k] || 0;
                layerEinmal += eVal;
                layerSpar += sVal;

                const ertragTxt = f.ertrag ? ` · ${f.ertrag}` : '';

                const item = document.createElement('div');
                item.className = 'panel-fund-item';
                item.style.cssText = 'display:flex; justify-content:space-between; align-items:center; padding:6px 0; border-top:1px dashed #e2e8f0;';
                item.innerHTML = `
                    <div class="panel-fund-info" style="flex:1; padding-right:8px;">
                        <div class="panel-fund-name" style="font-weight:600; font-size:13px; color:#1e293b;" title="${f.name}">${f.name}</div>
                        <div class="panel-fund-meta" style="font-size:11px; color:#64748b;">
                            ${f.info || ''}${ertragTxt}
                            ${(eVal > 0 || sVal > 0) ? ` · <strong style="color:var(--color-mlp-blau);">${eVal > 0 ? formatCurrency(eVal) : ''} ${sVal > 0 ? `(${formatCurrency(sVal)} mtl.)` : ''}</strong>` : ''}
                        </div>
                    </div>
                    <button class="panel-fund-remove" title="Entfernen" data-block="${layer.blockId}" data-fund="${f.name}"
                        style="background:none; border:none; color:#ef4444; font-size:14px; font-weight:bold; cursor:pointer; padding:2px 6px; border-radius:4px; transition:background 0.15s;">
                        ✕
                    </button>`;
                layerEl.appendChild(item);
            });

            // ── "Von Empfehlungsliste genommen" Sonderfeld (V4.4) ──────
            (() => {
                const eKey = `${layer.blockId}::empfehlungsliste`;
                const eInv = portfolioGlobals.fundInvestments[eKey] || 0;
                const eSpar = portfolioGlobals.fundSparrates[eKey] || 0;

                let empFondsForBlock = [];
                try {
                    const allEmpFonds = JSON.parse(localStorage.getItem('empfehlungslisteFonds_V44') || '{}');
                    empFondsForBlock = allEmpFonds[layer.blockId] || [];
                } catch { empFondsForBlock = []; }

                if (eInv === 0 && eSpar === 0 && empFondsForBlock.length === 0) return;

                layerEinmal += eInv;
                layerSpar += eSpar;

                const eRow = document.createElement('div');
                eRow.className = 'panel-fund-item panel-empfehlung-row';

                const eInvFmt = eInv > 0 ? formatNumberInput(eInv) : '';
                const eSparFmt = eSpar > 0 ? formatNumberInput(eSpar) : '';

                let popupHtml = '';
                if (empFondsForBlock.length > 0) {
                    const popupItems = empFondsForBlock.map(u =>
                        `<span><b>${u.name}</b> · WKN ${u.wkn} · ${u.schwerpunkt || '?'} · ${formatCurrency(u.einmal || 0)}${u.sparrate > 0 ? ' + ' + formatCurrency(u.sparrate) + ' mtl.' : ''}</span>`
                    ).join('');
                    popupHtml = `<div class="empfehlung-popup">${popupItems}</div>`;
                }

                eRow.innerHTML = `
                    <div class="panel-fund-info" style="width:100%;">
                        <div class="panel-fund-name empfehlung-label empfehlung-label-hoverable" title="Fonds, die nicht mehr auf der VEM-Empfehlungsliste stehen" style="position:relative; cursor:${empFondsForBlock.length > 0 ? 'help' : 'default'}; font-weight:700;">
                            <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#d97706" stroke-width="2.5" style="vertical-align:-2px; margin-right:4px;"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>
                            Von Empfehlungsliste genommen
                            ${popupHtml}
                        </div>
                        <div class="panel-fund-meta" style="display:flex; gap:8px; align-items:center; flex-wrap:wrap; margin-top:6px;">
                            <input class="fund-investment-input empfehlung-input" type="text" inputmode="decimal"
                                placeholder="Einmal €" title="Einmalbetrag"
                                data-ekey="${eKey}" data-spar="0"
                                value="${eInvFmt}" style="width:100px;">
                            <input class="fund-investment-input empfehlung-input" type="text" inputmode="decimal"
                                placeholder="Sparrate €" title="Sparrate monatlich"
                                data-ekey="${eKey}" data-spar="1"
                                value="${eSparFmt}" style="width:100px;">
                        </div>
                    </div>`;

                eRow.querySelectorAll('.empfehlung-input').forEach(inp => {
                    inp.addEventListener('change', () => {
                        const v = parseCurrencyInput(inp.value);
                        const k = inp.dataset.ekey;
                        const isSpar = inp.dataset.spar === '1';
                        if (isSpar) {
                            if (v > 0) portfolioGlobals.fundSparrates[k] = v;
                            else delete portfolioGlobals.fundSparrates[k];
                        } else {
                            if (v > 0) portfolioGlobals.fundInvestments[k] = v;
                            else delete portfolioGlobals.fundInvestments[k];
                        }
                        saveGlobals();
                        updatePortfolioUI();
                    });
                });

                layerEl.appendChild(eRow);
            });

            div.innerHTML = `
                <div class="panel-layer-header" style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
                    <strong style="color:${color}; font-size:14px;">${block.title}</strong>
                    <span style="font-weight:bold; font-size:12px; color:#334155;">
                        ${layerEinmal > 0 ? formatCurrency(layerEinmal) : ''} ${layerSpar > 0 ? `(${formatCurrency(layerSpar)} mtl.)` : ''}
                    </span>
                </div>`;

            const listWrapper = document.createElement('div');
            listWrapper.className = 'panel-funds-list';
            listWrapper.appendChild(layerEl);
            div.appendChild(listWrapper);

            div.querySelectorAll('.panel-fund-remove').forEach(btn => {
                btn.addEventListener('click', (e) => {
                    e.stopPropagation();
                    const bId = btn.dataset.block;
                    const fName = btn.dataset.fund;
                    removeFund(bId, fName);
                    if (currentBlockData && modal && modal.classList.contains('open')) {
                        renderFundList(currentBlockData.funds);
                    }
                    updatePortfolioUI();
                });
            });

            panelLayers.appendChild(div);
        });
    }

    if (savePortfolioBtn) {
        savePortfolioBtn.addEventListener('click', () => {
            if (!window.jspdf) return;
            const { jsPDF } = window.jspdf;
            const doc = new jsPDF();
            const now = new Date();
            const dateStr = now.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' });
            const fileName = `Anlagevorschlag_V5_${dateStr.replace(/\./g, '-')}.pdf`;

            doc.setFontSize(18); doc.setTextColor(3, 61, 93);
            doc.text('Anlagevorschlag - Robustes Portfolio V5.0', 14, 20);
            doc.setFontSize(11); doc.setTextColor(50, 50, 50);
            doc.text(`Datum: ${dateStr}`, 14, 30);
            doc.text(`Anzulegendes Gesamtvermögen: ${formatCurrency(portfolioGlobals.totalInvestment)}`, 14, 38);

            const tableData = [];
            managementBlocks.forEach(block => {
                block.funds.forEach(fund => {
                    const key = `${block.id}::${fund.name}`;
                    const amountEinmal = portfolioGlobals.fundInvestments[key] || 0;
                    const amountSpar = portfolioGlobals.fundSparrates[key] || 0;
                    if (amountEinmal > 0 || amountSpar > 0) {
                        const ertragStr = fund.ertrag ? ` (${fund.ertrag})` : '';
                        let amountStr = '';
                        if (amountEinmal > 0) amountStr += formatCurrency(amountEinmal);
                        if (amountSpar > 0) amountStr += (amountStr ? ' + ' : '') + formatCurrency(amountSpar) + ' mtl.';
                        tableData.push([fund.name, fund.info, `${fund.type}${ertragStr}`, amountStr]);
                    }
                });
            });
            if (tableData.length === 0) return;

            doc.autoTable({
                startY: 45,
                head: [['Fondsname', 'WKN / ISIN', 'Typ / Ertrag', 'Anlagebetrag']],
                body: tableData,
                theme: 'striped',
                headStyles: { fillColor: [3, 61, 93], textColor: [255, 255, 255] }
            });

            doc.save(fileName);
        });
    }

    if (resetBtn) {
        resetBtn.addEventListener('click', () => {
            resetPortfolio();
            portfolioGlobals.fundInvestments = {};
            portfolioGlobals.fundSparrates = {};
            saveGlobals();
            updatePortfolioUI();
        });
    }

    updatePortfolioUI();
});


// ============================================================
//  CRAWLER / LÄNDER AKTUALISIEREN (PIN 2203 / 220363)
// ============================================================
const CRAWLER_PORT = 8765;

function initCrawlerFeature() {
    const crawlerBtn       = document.getElementById('crawler-btn');
    const crawlerModal     = document.getElementById('crawler-modal');
    const crawlerModalClose= document.getElementById('crawler-modal-close');
    const crawlerNoServer  = document.getElementById('crawler-no-server');
    const crawlerReady     = document.getElementById('crawler-ready');
    const crawlerRunning   = document.getElementById('crawler-running');
    const crawlerDone      = document.getElementById('crawler-done');
    const crawlerStartBtn  = document.getElementById('crawler-start-btn');
    const crawlerLog       = document.getElementById('crawler-log');
    const crawlerBar       = document.getElementById('crawler-progress-bar');
    const crawlerLabel     = document.getElementById('crawler-progress-label');
    const crawlerDoneMsg   = document.getElementById('crawler-done-msg');
    const copyCmdBtn       = document.getElementById('copy-cmd-btn');
    let pollInterval       = null;

    const SERVER = `http://localhost:${CRAWLER_PORT}`;

    const showPane = (pane) => {
        [crawlerNoServer, crawlerReady, crawlerRunning, crawlerDone]
            .forEach(el => el && (el.style.display = 'none'));
        if (pane) pane.style.display = 'block';
    };

    const appendLog = (msg, color = '#cdd6f4') => {
        if (!crawlerLog) return;
        const line = document.createElement('div');
        line.style.color = color;
        line.textContent = msg;
        crawlerLog.appendChild(line);
        crawlerLog.scrollTop = crawlerLog.scrollHeight;
    };

    const stopPolling = () => { if (pollInterval) { clearInterval(pollInterval); pollInterval = null; } };

    const pollStatus = () => {
        fetch(`${SERVER}/status`)
            .then(r => r.json())
            .then(data => {
                const steps = data.progress || [];
                const total = steps.length > 0 ? steps[steps.length-1].total || 1 : 1;
                const done  = steps.filter(s => s.status !== 'running').length;
                if (crawlerBar)   crawlerBar.style.width = `${Math.round((done/total)*100)}%`;
                if (crawlerLabel) crawlerLabel.textContent = `[${done}/${total}] ${steps.length > 0 ? steps[steps.length-1].fund : ''}...`;
                if (steps.length > (crawlerLog ? crawlerLog.children.length : 0)) {
                    const last = steps[steps.length-1];
                    const ok = ['ok','ok_fallback'].includes(last.status);
                    appendLog(`[${done}/${total}] ${last.fund}: ${ok ? '✅' : '❌'} ${last.status}`,
                              ok ? '#a6e3a1' : '#f38ba8');
                }
                if (data.done) {
                    stopPolling();
                    const okCount = steps.filter(s => ['ok','ok_fallback'].includes(s.status)).length;
                    if (crawlerDoneMsg) crawlerDoneMsg.textContent = `Fertig! ✅ ${okCount} Fonds aktualisiert, ❌ ${steps.length - okCount} nicht gefunden.`;
                    showPane(crawlerDone);
                }
                if (data.error) {
                    stopPolling();
                    appendLog(`Fehler: ${data.error}`, '#f38ba8');
                }
            })
            .catch(() => stopPolling());
    };

    if (crawlerBtn) {
        crawlerBtn.addEventListener('click', () => {
            const pwOverlay = document.createElement('div');
            pwOverlay.style.cssText = `
                position:fixed; inset:0; background:rgba(0,0,0,0.65);
                display:flex; align-items:center; justify-content:center; z-index:9999;`;
            pwOverlay.innerHTML = `
                <div style="background:#1e2130; border:1px solid #3d4258; border-radius:14px;
                    padding:36px 40px; min-width:320px; text-align:center; box-shadow:0 8px 40px rgba(0,0,0,0.5);">
                    <div style="font-size:2em; margin-bottom:12px;">🔒</div>
                    <div style="color:#e0e4f0; font-size:1.1em; font-weight:600; margin-bottom:6px;">Zugang geschützt</div>
                    <div style="color:#9aa0bc; font-size:0.88em; margin-bottom:22px;">Bitte PIN eingeben, um die Aktualisierung zu starten.</div>
                    <input id="pw-input" type="password" maxlength="20"
                        placeholder="PIN eingeben"
                        style="width:100%; box-sizing:border-box; padding:10px 14px; font-size:1.1em;
                            border:1.5px solid #3d4258; border-radius:8px; background:#131520;
                            color:#e0e4f0; outline:none; text-align:center; letter-spacing:4px;"
                    />
                    <div id="pw-error" style="color:#f38ba8; font-size:0.85em; margin-top:10px; min-height:18px;"></div>
                    <div style="display:flex; gap:12px; margin-top:20px; justify-content:center;">
                        <button id="pw-cancel" style="padding:9px 24px; border-radius:8px; border:1px solid #3d4258;
                            background:transparent; color:#9aa0bc; cursor:pointer; font-size:0.95em;">Abbrechen</button>
                        <button id="pw-confirm" style="padding:9px 28px; border-radius:8px; border:none;
                            background:#034d6e; color:#fff; cursor:pointer; font-size:0.95em; font-weight:600;">Bestätigen</button>
                    </div>
                </div>`;
            document.body.appendChild(pwOverlay);
            const pwInput   = pwOverlay.querySelector('#pw-input');
            const pwError   = pwOverlay.querySelector('#pw-error');
            const pwConfirm = pwOverlay.querySelector('#pw-confirm');
            const pwCancel  = pwOverlay.querySelector('#pw-cancel');
            setTimeout(() => pwInput.focus(), 50);

            const checkPw = () => {
                const val = pwInput.value.trim();
                if (val === '2203' || val === '220363') {
                    document.body.removeChild(pwOverlay);
                    if (crawlerModal) crawlerModal.style.display = 'flex';
                    showPane(null);
                    fetch(`${SERVER}/ping`, { signal: AbortSignal.timeout(2500) })
                        .then(r => r.json())
                        .then(() => showPane(crawlerReady))
                        .catch(() => showPane(crawlerNoServer));
                } else {
                    pwError.textContent = 'Falsche PIN. Bitte erneut versuchen.';
                    pwInput.value = '';
                    pwInput.focus();
                    pwInput.style.borderColor = '#f38ba8';
                    setTimeout(() => { pwInput.style.borderColor = '#3d4258'; pwError.textContent = ''; }, 2000);
                }
            };

            pwConfirm.addEventListener('click', checkPw);
            pwInput.addEventListener('keydown', e => { if (e.key === 'Enter') checkPw(); });
            pwCancel.addEventListener('click', () => document.body.removeChild(pwOverlay));
            pwOverlay.addEventListener('click', e => { if (e.target === pwOverlay) document.body.removeChild(pwOverlay); });
        });
    }

    if (crawlerModalClose) {
        crawlerModalClose.addEventListener('click', () => {
            if (crawlerModal) crawlerModal.style.display = 'none';
            stopPolling();
        });
    }
    window.addEventListener('click', e => {
        if (e.target === crawlerModal) { crawlerModal.style.display = 'none'; stopPolling(); }
    });

    if (copyCmdBtn) {
        copyCmdBtn.addEventListener('click', () => {
            const dir = window.location.href.replace('index.html','').replace('file://','');
            const cmd = `cd "${decodeURIComponent(dir)}" && python3 crawler_server.py`;
            navigator.clipboard.writeText(cmd).then(() => { copyCmdBtn.textContent = '✅ Kopiert!'; setTimeout(() => { copyCmdBtn.textContent = '📋 Kopieren'; }, 2000); });
        });
    }

    if (crawlerStartBtn) {
        crawlerStartBtn.addEventListener('click', () => {
            showPane(crawlerRunning);
            if (crawlerLog) crawlerLog.innerHTML = '';
            if (crawlerBar) crawlerBar.style.width = '0%';
            if (crawlerLabel) crawlerLabel.textContent = 'Starte Crawler...';
            appendLog('Verbinde mit Crawle-Server...');
            fetch(`${SERVER}/run`, { method: 'POST' })
                .then(r => r.json())
                .then(d => {
                    if (d.ok) {
                        appendLog('Crawler gestartet 🚀', '#89b4fa');
                        pollInterval = setInterval(pollStatus, 2000);
                    } else {
                        appendLog(`Fehler: ${d.error}`, '#f38ba8');
                    }
                })
                .catch(e => appendLog(`Verbindungsfehler: ${e}`, '#f38ba8'));
        });
    }
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initCrawlerFeature);
} else {
    initCrawlerFeature();
}
