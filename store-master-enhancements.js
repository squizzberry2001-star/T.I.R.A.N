(function () {
    const STORE_MASTER_STORAGE_KEY = 'regional-bestie-store-master-v1';
    const STORE_MASTER_META_KEY = 'regional-bestie-store-master-meta-v1';
    const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

    const DEFAULT_MASTER_META = window.DEFAULT_STORE_MASTER_META || {
        label: 'Master Store Default',
        source: 'default',
        recordCount: 0
    };

    let storeMasterRecords = [];
    let storeMasterMeta = { ...DEFAULT_MASTER_META };
    let storeMasterIndexByCode = {};
    let storeMasterIndexByName = {};

    function safeText(value) {
        if (value === undefined || value === null) return '';
        const text = String(value).replace(/\u00A0/g, ' ').trim();
        return text === '0' ? '' : text;
    }

    function normalizeLookup(value) {
        if (typeof window.normalizeLookupValue === 'function') {
            return window.normalizeLookupValue(value);
        }
        return safeText(value)
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .toUpperCase()
            .replace(/[^A-Z0-9\s]/g, ' ')
            .replace(/\s+/g, ' ')
            .trim();
    }

    function extractDigits(value) {
        return safeText(value).replace(/\D/g, '');
    }

    function formatSiteCode(value) {
        const digits = extractDigits(value);
        if (!digits) return '';
        return digits.length >= 4 ? digits : digits.padStart(4, '0');
    }

    function standardizeRecord(record) {
        const siteCode = extractDigits(record.siteCode || record.site || record.Site || record.code || record.storeCode);
        return {
            siteCode,
            siteCode4: formatSiteCode(record.siteCode4 || siteCode),
            siteDescr: safeText(record.siteDescr || record.SiteDescr || record.name || record.storeName),
            type: safeText(record.type || record.Type),
            city: safeText(record.city || record.City),
            address: safeText(record.address || record.Address || record.storeAddress),
            emailStore: safeText(record.emailStore || record.email || record.Email),
            storeHead: safeText(record.storeHead || record['Store Head'] || record.storeHeadName),
            areaManager: safeText(record.areaManager || record['Area Manager']),
            areaManagerEmail: safeText(record.areaManagerEmail || record['Area Manager Email']),
            regionalManager: safeText(record.regionalManager || record['Regional Manager']),
            regionalManagerEmail: safeText(record.regionalManagerEmail || record['Regional Manager Email'])
        };
    }

    function mergeRecords(target, source) {
        const merged = { ...target };
        Object.keys(source || {}).forEach((key) => {
            if (safeText(source[key])) {
                merged[key] = source[key];
            }
        });
        return standardizeRecord(merged);
    }

    function buildMasterState(records) {
        const uniqueRecords = [];
        const byCode = {};
        const byName = {};

        (Array.isArray(records) ? records : []).forEach((row) => {
            const normalized = standardizeRecord(row);
            if (!normalized.siteCode4 && !normalized.siteDescr) return;
            const nameKey = normalizeLookup(normalized.siteDescr);
            const existing = (normalized.siteCode4 && byCode[normalized.siteCode4]) || (nameKey && byName[nameKey]) || null;

            if (existing) {
                const merged = mergeRecords(existing, normalized);
                Object.keys(existing).forEach((key) => delete existing[key]);
                Object.assign(existing, merged);
                if (merged.siteCode4) byCode[merged.siteCode4] = existing;
                if (nameKey) byName[nameKey] = existing;
                return;
            }

            const entry = { ...normalized };
            uniqueRecords.push(entry);
            if (entry.siteCode4) byCode[entry.siteCode4] = entry;
            if (nameKey) byName[nameKey] = entry;
        });

        return { records: uniqueRecords, byCode, byName };
    }

    function loadPersistedMaster() {
        let persistedRecords = null;
        let persistedMeta = null;

        try {
            const raw = localStorage.getItem(STORE_MASTER_STORAGE_KEY);
            persistedRecords = raw ? JSON.parse(raw) : null;
        } catch (error) {
            persistedRecords = null;
        }

        try {
            const rawMeta = localStorage.getItem(STORE_MASTER_META_KEY);
            persistedMeta = rawMeta ? JSON.parse(rawMeta) : null;
        } catch (error) {
            persistedMeta = null;
        }

        const baseRecords = Array.isArray(persistedRecords) && persistedRecords.length
            ? persistedRecords
            : (Array.isArray(window.DEFAULT_STORE_MASTER_DATA) ? window.DEFAULT_STORE_MASTER_DATA : []);

        const state = buildMasterState(baseRecords);
        storeMasterRecords = state.records;
        storeMasterIndexByCode = state.byCode;
        storeMasterIndexByName = state.byName;
        storeMasterMeta = Array.isArray(persistedRecords) && persistedRecords.length
            ? {
                label: safeText(persistedMeta && persistedMeta.label) || 'Import Excel',
                fileName: safeText(persistedMeta && persistedMeta.fileName),
                importedAt: safeText(persistedMeta && persistedMeta.importedAt),
                source: 'imported',
                recordCount: state.records.length
            }
            : {
                ...DEFAULT_MASTER_META,
                source: 'default',
                recordCount: state.records.length
            };
    }

    function persistImportedMaster(records, meta) {
        const state = buildMasterState(records);
        storeMasterRecords = state.records;
        storeMasterIndexByCode = state.byCode;
        storeMasterIndexByName = state.byName;
        storeMasterMeta = {
            label: safeText(meta && meta.label) || 'Import Excel',
            fileName: safeText(meta && meta.fileName),
            importedAt: safeText(meta && meta.importedAt) || new Date().toISOString(),
            source: 'imported',
            recordCount: state.records.length
        };
        localStorage.setItem(STORE_MASTER_STORAGE_KEY, JSON.stringify(state.records));
        localStorage.setItem(STORE_MASTER_META_KEY, JSON.stringify(storeMasterMeta));
    }

    function resetPersistedMaster() {
        localStorage.removeItem(STORE_MASTER_STORAGE_KEY);
        localStorage.removeItem(STORE_MASTER_META_KEY);
        loadPersistedMaster();
    }

    function getAssignmentFallback(storeName) {
        if (typeof window.getStoreDetailByName === 'function') {
            try {
                return window.getStoreDetailByName(storeName) || null;
            } catch (error) {
                return null;
            }
        }
        return null;
    }

    function findStoreMasterDetail(storeName, fallbackDetail) {
        const codeCandidates = [
            fallbackDetail && (fallbackDetail.siteCode4 || fallbackDetail.siteCode || fallbackDetail.code || fallbackDetail.storeCode),
            storeName
        ];

        for (const candidate of codeCandidates) {
            const code4 = formatSiteCode(candidate);
            if (code4 && storeMasterIndexByCode[code4]) {
                return storeMasterIndexByCode[code4];
            }
        }

        const nameCandidates = [
            storeName,
            fallbackDetail && (fallbackDetail.siteDescr || fallbackDetail.name || fallbackDetail.storeName || fallbackDetail.assignmentStoreName)
        ];

        for (const candidate of nameCandidates) {
            const key = normalizeLookup(candidate);
            if (key && storeMasterIndexByName[key]) {
                return storeMasterIndexByName[key];
            }
        }

        return null;
    }

    function buildWebStoreDetail(storeName, fallbackDetail) {
        const fallback = fallbackDetail || getAssignmentFallback(storeName) || {};
        const master = findStoreMasterDetail(storeName, fallback) || {};
        return {
            siteCode4: safeText(master.siteCode4 || formatSiteCode(master.siteCode || fallback.code || fallback.storeCode)),
            siteDescr: safeText(master.siteDescr || fallback.name || storeName),
            type: safeText(master.type),
            city: safeText(master.city || fallback.city),
            address: safeText(master.address || fallback.address),
            emailStore: safeText(master.emailStore),
            storeHead: safeText(master.storeHead || fallback.storeHead),
            areaManager: safeText(master.areaManager || fallback.areaManager),
            areaManagerEmail: safeText(master.areaManagerEmail),
            regionalManager: safeText(master.regionalManager || fallback.regionalManager),
            regionalManagerEmail: safeText(master.regionalManagerEmail),
            hasMasterRecord: Boolean(master && (master.siteCode4 || master.siteDescr)),
            source: storeMasterMeta.source || 'default'
        };
    }

    window.getStoreWebDetail = function getStoreWebDetail(storeName) {
        return buildWebStoreDetail(storeName, getAssignmentFallback(storeName) || {});
    };

    window.getStoreMasterRecords = function getStoreMasterRecords() {
        return storeMasterRecords.slice();
    };

    function formatDateTimeText(value) {
        if (!value) return '-';
        try {
            return new Date(value).toLocaleString('id-ID', {
                day: '2-digit',
                month: '2-digit',
                year: 'numeric',
                hour: '2-digit',
                minute: '2-digit'
            });
        } catch (error) {
            return safeText(value) || '-';
        }
    }

    function ensureStoreDetailModal() {
        if (document.getElementById('storeDetailModal')) return;

        const overlay = document.createElement('div');
        overlay.id = 'storeDetailModal';
        overlay.className = 'secret-modal-overlay';
        overlay.hidden = true;
        overlay.innerHTML = [
            '<div class="secret-card store-detail-card">',
            '  <button type="button" class="secret-close" id="storeDetailCloseBtn" aria-label="Tutup">&times;</button>',
            '  <div class="secret-badge">Detail Store</div>',
            '  <h2 id="storeDetailTitle" style="margin:0; color:#132238;">-</h2>',
            '  <div class="store-detail-grid">',
            '    <div class="store-detail-item"><div class="store-detail-label">Site</div><div class="store-detail-value" id="storeDetailSite">-</div></div>',
            '    <div class="store-detail-item"><div class="store-detail-label">SiteDescr</div><div class="store-detail-value" id="storeDetailSiteDescr">-</div></div>',
            '    <div class="store-detail-item"><div class="store-detail-label">Type</div><div class="store-detail-value" id="storeDetailType">-</div></div>',
            '    <div class="store-detail-item"><div class="store-detail-label">City</div><div class="store-detail-value" id="storeDetailCity">-</div></div>',
            '    <div class="store-detail-item store-detail-item-wide"><div class="store-detail-label">Address</div><div class="store-detail-value" id="storeDetailAddress">-</div></div>',
            '    <div class="store-detail-item"><div class="store-detail-label">Email Store</div><div class="store-detail-value" id="storeDetailEmailStore">-</div></div>',
            '    <div class="store-detail-item"><div class="store-detail-label">Store Head</div><div class="store-detail-value" id="storeDetailStoreHead">-</div></div>',
            '    <div class="store-detail-item"><div class="store-detail-label">Area Manager</div><div class="store-detail-value" id="storeDetailAreaManager">-</div></div>',
            '    <div class="store-detail-item"><div class="store-detail-label">Area Manager Email</div><div class="store-detail-value" id="storeDetailAreaManagerEmail">-</div></div>',
            '    <div class="store-detail-item"><div class="store-detail-label">Regional Manager</div><div class="store-detail-value" id="storeDetailRegionalManager">-</div></div>',
            '    <div class="store-detail-item"><div class="store-detail-label">Regional Manager Email</div><div class="store-detail-value" id="storeDetailRegionalManagerEmail">-</div></div>',
            '  </div>',
            '</div>'
        ].join('');

        document.body.appendChild(overlay);

        document.getElementById('storeDetailCloseBtn')?.addEventListener('click', closeStoreDetailModal);
        overlay.addEventListener('click', (event) => {
            if (event.target && event.target.id === 'storeDetailModal') {
                closeStoreDetailModal();
            }
        });
    }

    function populateStoreDetailModal(detail) {
        const mapping = {
            storeDetailTitle: detail.siteDescr || '-',
            storeDetailSite: detail.siteCode4 || '-',
            storeDetailSiteDescr: detail.siteDescr || '-',
            storeDetailType: detail.type || '-',
            storeDetailCity: detail.city || '-',
            storeDetailAddress: detail.address || '-',
            storeDetailEmailStore: detail.emailStore || '-',
            storeDetailStoreHead: detail.storeHead || '-',
            storeDetailAreaManager: detail.areaManager || '-',
            storeDetailAreaManagerEmail: detail.areaManagerEmail || '-',
            storeDetailRegionalManager: detail.regionalManager || '-',
            storeDetailRegionalManagerEmail: detail.regionalManagerEmail || '-'
        };

        Object.entries(mapping).forEach(([id, value]) => {
            const element = document.getElementById(id);
            if (element) element.textContent = value || '-';
        });
    }

    function getCurrentSelectedStoreDetail() {
        const storeInput = document.getElementById('store');
        const rawStore = safeText(storeInput && storeInput.value);
        if (!rawStore) return null;

        let resolved = null;
        if (typeof window.resolveStoreName === 'function') {
            try {
                resolved = window.resolveStoreName(rawStore);
            } catch (error) {
                resolved = null;
            }
        }

        const canonicalName = resolved && resolved.matched ? resolved.canonicalName : rawStore;
        const fallback = (resolved && resolved.detail) || getAssignmentFallback(canonicalName) || getAssignmentFallback(rawStore) || {};
        const detail = buildWebStoreDetail(canonicalName, fallback);

        if (!detail.siteDescr && !detail.siteCode4 && !detail.address) {
            return null;
        }

        return detail;
    }

    function openStoreDetailModal() {
        const detail = getCurrentSelectedStoreDetail();
        if (!detail) return;
        ensureStoreDetailModal();
        populateStoreDetailModal(detail);
        const modal = document.getElementById('storeDetailModal');
        if (modal) modal.hidden = false;
    }

    function closeStoreDetailModal() {
        const modal = document.getElementById('storeDetailModal');
        if (modal) modal.hidden = true;
    }

    function ensureStoreInfoButton() {
        const combo = document.getElementById('store') && document.getElementById('store').closest('.custom-combobox');
        if (!combo || document.getElementById('storeDetailBtn')) return;
        combo.classList.add('store-combobox');

        const menu = document.getElementById('storeDropdown');
        const button = document.createElement('button');
        button.type = 'button';
        button.id = 'storeDetailBtn';
        button.className = 'combo-info-btn';
        button.setAttribute('aria-label', 'Lihat detail store');
        button.textContent = '!';
        button.addEventListener('click', openStoreDetailModal);

        if (menu && menu.parentNode === combo) {
            combo.insertBefore(button, menu);
        } else {
            combo.appendChild(button);
        }
    }

    function syncStoreInfoButtonState() {
        const button = document.getElementById('storeDetailBtn');
        if (!button) return;
        const detail = getCurrentSelectedStoreDetail();
        const enabled = Boolean(detail && (detail.siteDescr || detail.siteCode4 || detail.address));
        button.disabled = !enabled;
        button.classList.toggle('is-ready', enabled);
        button.title = enabled ? 'Detail store' : 'Pilih store dahulu';
    }

    function normalizeHeader(value) {
        return safeText(value).toUpperCase().replace(/[^A-Z0-9]/g, '');
    }

    function getTextNodesJoined(node, tagName) {
        return Array.from(node.getElementsByTagName(tagName || 't'))
            .map((item) => item.textContent || '')
            .join('');
    }

    function columnLettersToIndex(ref) {
        let index = 0;
        for (let i = 0; i < ref.length; i += 1) {
            index = (index * 26) + (ref.charCodeAt(i) - 64);
        }
        return Math.max(index - 1, 0);
    }

    function resolveZipPath(baseDir, targetPath) {
        const base = safeText(baseDir).replace(/\\/g, '/').replace(/\/+$/, '');
        const target = safeText(targetPath).replace(/\\/g, '/');
        if (!target) return '';
        if (target.startsWith('/')) return target.replace(/^\/+/, '');
        const parts = base ? base.split('/') : [];
        target.split('/').forEach((part) => {
            if (!part || part === '.') return;
            if (part === '..') {
                parts.pop();
            } else {
                parts.push(part);
            }
        });
        return parts.join('/');
    }

    async function parseSharedStrings(zip, parser) {
        const file = zip.file('xl/sharedStrings.xml');
        if (!file) return [];
        const xml = await file.async('string');
        const doc = parser.parseFromString(xml, 'application/xml');
        return Array.from(doc.getElementsByTagName('si')).map((node) => getTextNodesJoined(node, 't'));
    }

    function readCellValue(cellNode, sharedStrings) {
        const type = cellNode.getAttribute('t') || '';
        if (type === 'inlineStr') {
            return getTextNodesJoined(cellNode, 't');
        }
        const valueNode = cellNode.getElementsByTagName('v')[0];
        const raw = valueNode ? valueNode.textContent || '' : '';
        if (type === 's') {
            const index = Number(raw || 0);
            return sharedStrings[index] || '';
        }
        if (type === 'b') {
            return raw === '1' ? 'TRUE' : 'FALSE';
        }
        return raw;
    }

    function parseSheetRows(sheetXml, sharedStrings, parser) {
        const doc = parser.parseFromString(sheetXml, 'application/xml');
        const rows = [];
        Array.from(doc.getElementsByTagName('row')).forEach((rowNode) => {
            const row = [];
            Array.from(rowNode.getElementsByTagName('c')).forEach((cellNode) => {
                const ref = cellNode.getAttribute('r') || '';
                const colRef = (ref.match(/[A-Z]+/i) || ['A'])[0].toUpperCase();
                const columnIndex = columnLettersToIndex(colRef);
                row[columnIndex] = readCellValue(cellNode, sharedStrings);
            });
            rows.push(row);
        });
        return rows;
    }

    const HEADER_ALIASES = {
        siteCode: ['SITE', 'STORECODE', 'KODETOKO'],
        siteDescr: ['SITEDESCR', 'STORENAME', 'STORE', 'STOREYANGDITANGGUNGJAWABKAN'],
        type: ['TYPE'],
        city: ['CITY', 'KOTA'],
        address: ['ADDRESS', 'ALAMAT', 'ALAMATSTORE'],
        emailStore: ['EMAIL', 'EMAILSTORE', 'STOREEMAIL'],
        storeHead: ['STOREHEAD'],
        areaManager: ['AREAMANAGER'],
        areaManagerEmail: ['AREAMANAGEREMAIL'],
        regionalManager: ['REGIONALMANAGER'],
        regionalManagerEmail: ['REGIONALMANAGEREMAIL']
    };

    function detectHeaderMap(rows) {
        let best = null;

        rows.slice(0, 40).forEach((row, rowIndex) => {
            const map = {};
            row.forEach((cellValue, columnIndex) => {
                const header = normalizeHeader(cellValue);
                Object.entries(HEADER_ALIASES).forEach(([field, aliases]) => {
                    if (!map[field] && aliases.includes(header)) {
                        map[field] = columnIndex;
                    }
                });
            });

            const score = Object.keys(map).length;
            if (!best || score > best.score) {
                best = { rowIndex, map, score };
            }
        });

        if (!best || best.score < 5 || best.map.siteCode === undefined || best.map.siteDescr === undefined) {
            return null;
        }

        return best;
    }

    function rowValue(row, columnIndex) {
        return columnIndex === undefined ? '' : safeText(row[columnIndex]);
    }

    async function parseStoreMasterExcel(file) {
        if (!window.JSZip) {
            throw new Error('Mesin baca Excel belum tersedia.');
        }

        const zip = await window.JSZip.loadAsync(await file.arrayBuffer());
        const parser = new DOMParser();
        const workbookFile = zip.file('xl/workbook.xml');
        if (!workbookFile) {
            throw new Error('File Excel tidak valid. Gunakan file .xlsx.');
        }

        const workbookXml = await workbookFile.async('string');
        const workbookDoc = parser.parseFromString(workbookXml, 'application/xml');
        const relsFile = zip.file('xl/_rels/workbook.xml.rels');
        const relsXml = relsFile ? await relsFile.async('string') : '';
        const relsDoc = relsXml ? parser.parseFromString(relsXml, 'application/xml') : null;
        const relsMap = {};

        if (relsDoc) {
            Array.from(relsDoc.getElementsByTagName('Relationship')).forEach((node) => {
                const id = node.getAttribute('Id');
                const target = node.getAttribute('Target');
                if (id && target) relsMap[id] = target;
            });
        }

        const sharedStrings = await parseSharedStrings(zip, parser);
        const sheets = Array.from(workbookDoc.getElementsByTagName('sheet')).map((sheetNode) => {
            const relId = sheetNode.getAttribute('r:id') || sheetNode.getAttributeNS(REL_NS, 'id') || sheetNode.getAttribute('id');
            const name = sheetNode.getAttribute('name') || 'Sheet';
            const target = relsMap[relId] || '';
            return { name, target };
        });

        let bestResult = null;

        for (const sheet of sheets) {
            const path = resolveZipPath('xl', sheet.target);
            const sheetFile = path ? zip.file(path) : null;
            if (!sheetFile) continue;
            const sheetXml = await sheetFile.async('string');
            const rows = parseSheetRows(sheetXml, sharedStrings, parser);
            const detectedHeader = detectHeaderMap(rows);
            if (!detectedHeader) continue;

            const parsedRows = rows.slice(detectedHeader.rowIndex + 1)
                .map((row) => standardizeRecord({
                    siteCode: rowValue(row, detectedHeader.map.siteCode),
                    siteDescr: rowValue(row, detectedHeader.map.siteDescr),
                    type: rowValue(row, detectedHeader.map.type),
                    city: rowValue(row, detectedHeader.map.city),
                    address: rowValue(row, detectedHeader.map.address),
                    emailStore: rowValue(row, detectedHeader.map.emailStore),
                    storeHead: rowValue(row, detectedHeader.map.storeHead),
                    areaManager: rowValue(row, detectedHeader.map.areaManager),
                    areaManagerEmail: rowValue(row, detectedHeader.map.areaManagerEmail),
                    regionalManager: rowValue(row, detectedHeader.map.regionalManager),
                    regionalManagerEmail: rowValue(row, detectedHeader.map.regionalManagerEmail)
                }))
                .filter((record) => record.siteCode4 || record.siteDescr);

            const result = {
                sheetName: sheet.name,
                score: detectedHeader.score,
                records: buildMasterState(parsedRows).records
            };

            if (!bestResult || result.score > bestResult.score || result.records.length > bestResult.records.length) {
                bestResult = result;
            }
        }

        if (!bestResult || !bestResult.records.length) {
            throw new Error('Header master store tidak ditemukan. Pastikan file punya kolom Site, SiteDescr, Type, City, Address, Email, Store Head, Area Manager, Area Manager Email, Regional Manager, dan Regional Manager Email.');
        }

        return bestResult;
    }

    function ensureStoreMasterPanel() {
        const secretModal = document.getElementById('secretMonitorModal');
        if (!secretModal || document.getElementById('storeMasterImportPanel')) return;
        const toolbar = secretModal.querySelector('.secret-toolbar');
        if (!toolbar) return;

        const panel = document.createElement('div');
        panel.id = 'storeMasterImportPanel';
        panel.className = 'secret-import-box';
        panel.innerHTML = [
            '<div class="secret-import-head">',
            '  <div>',
            '    <h3 class="secret-import-title">Master Data Store</h3>',
            '    <p class="secret-import-subtitle">Upload Excel master store untuk detail store di web.</p>',
            '  </div>',
            '  <span id="storeMasterRecordCount" class="secret-pill neutral">0 store</span>',
            '</div>',
            '<div class="secret-import-actions">',
            '  <input type="file" id="storeMasterFileInput" accept=".xlsx" hidden>',
            '  <button type="button" class="btn btn-secondary" id="storeMasterImportBtn">Import Excel</button>',
            '  <button type="button" class="btn btn-secondary" id="storeMasterResetBtn">Kembali ke Default</button>',
            '</div>',
            '<div id="storeMasterImportStatus" class="secret-import-status"></div>'
        ].join('');

        toolbar.insertAdjacentElement('afterend', panel);

        document.getElementById('storeMasterImportBtn')?.addEventListener('click', () => {
            document.getElementById('storeMasterFileInput')?.click();
        });

        document.getElementById('storeMasterResetBtn')?.addEventListener('click', () => {
            resetPersistedMaster();
            refreshStoreMasterPanel();
            syncStoreInfoButtonState();
        });

        document.getElementById('storeMasterFileInput')?.addEventListener('change', async (event) => {
            const file = event.target && event.target.files ? event.target.files[0] : null;
            if (!file) return;
            setStoreMasterStatus('Membaca file Excel...', false, true);
            try {
                const parsed = await parseStoreMasterExcel(file);
                persistImportedMaster(parsed.records, {
                    label: 'Import Excel Master Store',
                    fileName: file.name,
                    importedAt: new Date().toISOString()
                });
                refreshStoreMasterPanel();
                syncStoreInfoButtonState();
            } catch (error) {
                setStoreMasterStatus(error && error.message ? error.message : 'Import Excel gagal.', true, false);
            } finally {
                if (event.target) event.target.value = '';
            }
        });
    }

    function setStoreMasterStatus(message, isError, isBusy) {
        const statusEl = document.getElementById('storeMasterImportStatus');
        if (!statusEl) return;
        statusEl.textContent = message;
        statusEl.classList.remove('is-error', 'is-busy');
        if (isError) statusEl.classList.add('is-error');
        if (isBusy) statusEl.classList.add('is-busy');
    }

    function refreshStoreMasterPanel() {
        const countEl = document.getElementById('storeMasterRecordCount');
        const resetBtn = document.getElementById('storeMasterResetBtn');
        if (countEl) {
            countEl.textContent = `${storeMasterRecords.length} store`;
        }
        if (resetBtn) {
            resetBtn.disabled = storeMasterMeta.source !== 'imported';
        }

        if (storeMasterMeta.source === 'imported') {
            const sourceName = storeMasterMeta.fileName || storeMasterMeta.label || 'Import Excel';
            setStoreMasterStatus(`Aktif: ${sourceName} | ${formatDateTimeText(storeMasterMeta.importedAt)}`, false, false);
        } else {
            const sourceName = storeMasterMeta.label || 'Master Store Default';
            setStoreMasterStatus(`Aktif: ${sourceName} | ${storeMasterRecords.length} store`, false, false);
        }
    }

    function patchStoreStatusMessage() {
        const originalValidateStoreSelection = typeof window.validateStoreSelection === 'function'
            ? window.validateStoreSelection
            : null;

        if (originalValidateStoreSelection && !window.__storeMasterValidatePatched) {
            window.__storeMasterValidatePatched = true;
            window.validateStoreSelection = function patchedValidateStoreSelection(showFeedback) {
                const valid = originalValidateStoreSelection.call(this, showFeedback);
                syncStoreInfoButtonState();
                if (valid && showFeedback && typeof window.setStoreStatus === 'function') {
                    const storeInput = document.getElementById('store');
                    const resolved = typeof window.resolveStoreName === 'function' ? window.resolveStoreName(storeInput && storeInput.value) : null;
                    if (resolved && resolved.matched) {
                        const detail = buildWebStoreDetail(resolved.canonicalName, resolved.detail || {});
                        const codeLabel = detail.siteCode4 ? `${detail.siteCode4} | ` : '';
                        window.setStoreStatus('match-success', `Store dipilih: ${codeLabel}${resolved.canonicalName}`);
                    }
                }
                return valid;
            };
        }

        const originalRenderSecretMonitorPayload = typeof window.renderSecretMonitorPayload === 'function'
            ? window.renderSecretMonitorPayload
            : null;

        if (originalRenderSecretMonitorPayload && !window.__storeMasterMonitorPatched) {
            window.__storeMasterMonitorPatched = true;
            window.renderSecretMonitorPayload = function patchedRenderSecretMonitorPayload(payload, sourceLabel) {
                const nextPayload = payload && typeof payload === 'object'
                    ? {
                        ...payload,
                        rows: Array.isArray(payload.rows)
                            ? payload.rows.map((item) => ({
                                ...item,
                                storeCode: formatSiteCode(item.storeCode) || safeText(item.storeCode)
                            }))
                            : []
                    }
                    : payload;
                return originalRenderSecretMonitorPayload.call(this, nextPayload, sourceLabel);
            };
        }
    }

    function bindStoreSelectionListeners() {
        const storeInput = document.getElementById('store');
        const visitorInput = document.getElementById('nama');
        const dateInput = document.getElementById('tanggal');

        [storeInput, visitorInput, dateInput].forEach((element) => {
            if (!element) return;
            element.addEventListener('input', syncStoreInfoButtonState);
            element.addEventListener('change', syncStoreInfoButtonState);
        });

        document.addEventListener('keydown', (event) => {
            if (event.key === 'Escape') {
                closeStoreDetailModal();
            }
        });

        window.addEventListener('storage', (event) => {
            if (event.key === STORE_MASTER_STORAGE_KEY || event.key === STORE_MASTER_META_KEY) {
                loadPersistedMaster();
                refreshStoreMasterPanel();
                syncStoreInfoButtonState();
            }
        });
    }

    function initializeStoreMasterFeature() {
        loadPersistedMaster();
        ensureStoreDetailModal();
        ensureStoreInfoButton();
        ensureStoreMasterPanel();
        patchStoreStatusMessage();
        bindStoreSelectionListeners();
        refreshStoreMasterPanel();
        syncStoreInfoButtonState();
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initializeStoreMasterFeature);
    } else {
        initializeStoreMasterFeature();
    }
})();
