(function () {
    const PPT = {
        W: 13.333,
        H: 7.5,
        HEADER_H: 0.78,
        M: 0.42,
        GAP: 0.12,
        TEAL: '1B8F96',
        TEAL_DARK: '14636A',
        WHITE: 'FFFFFF',
        LIGHT: 'F6F8FB',
        ALT: 'F3F6FA',
        BORDER: 'D7E3F0',
        TEXT: '132238',
        MUTED: '607D8B',
        LINK: '0B57D0',
        ERROR: 'C62828'
    };

    const PHOTO_PLACEHOLDER_HTML = '<span class="upload-icon">📷</span><span class="upload-text">Klik untuk upload foto</span>';
    const definedMasters = new Set();

    function dispatchInput(target) {
        if (!target) return;
        target.dispatchEvent(new Event('input', { bubbles: true }));
        target.dispatchEvent(new Event('change', { bubbles: true }));
    }

    function normalizeText(value, fallback = '—') {
        const str = value == null ? '' : String(value).trim();
        return str && str !== '-' ? str : fallback;
    }

    function sanitizeFileName(value) {
        return String(value || 'Regional_Bestie_Visit_Report')
            .replace(/[^\w\s-]/g, '')
            .trim()
            .replace(/\s+/g, '_') || 'Regional_Bestie_Visit_Report';
    }

    function isMeaningfulPhoto(photo) {
        if (!photo || typeof photo !== 'object') return false;
        const hasImage = !!photo.image;
        const desc = String(photo.description || '').trim();
        return hasImage || !!(desc && desc !== '-');
    }

    function resetPreview(previewEl, emptyText = PHOTO_PLACEHOLDER_HTML) {
        if (!previewEl) return;
        previewEl.innerHTML = emptyText;
        previewEl.classList.remove('has-image');
    }

    window.insertBullet = function insertBullet(button) {
        const textarea = button?.parentElement?.querySelector('textarea');
        if (!textarea) return;

        const value = textarea.value || '';
        const start = Number.isFinite(textarea.selectionStart) ? textarea.selectionStart : value.length;
        const end = Number.isFinite(textarea.selectionEnd) ? textarea.selectionEnd : value.length;
        const before = value.slice(0, start);
        const prefix = before.length === 0 || before.endsWith('\n') ? '• ' : '\n• ';
        const nextValue = value.slice(0, start) + prefix + value.slice(end);

        textarea.value = nextValue;
        const caret = start + prefix.length;
        textarea.focus();
        textarea.selectionStart = caret;
        textarea.selectionEnd = caret;

        if (typeof window.handleTextareaInput === 'function') {
            window.handleTextareaInput(textarea);
        }
        dispatchInput(textarea);
    };

    window.clearSinglePhoto = function clearSinglePhoto(inputId, previewId, descId) {
        const input = document.getElementById(inputId);
        const preview = document.getElementById(previewId);
        const desc = document.getElementById(descId);
        if (input) input.value = '';
        resetPreview(preview, previewId === 'qscResultPreview' ? '<span class="upload-icon">📷</span><span class="upload-text">Klik untuk upload foto QSC</span>' : PHOTO_PLACEHOLDER_HTML);
        if (desc) {
            desc.value = '';
            dispatchInput(desc);
        }
    };

    window.clearPhotoCell = function clearPhotoCell(trigger) {
        const cell = trigger?.closest('.photo-cell');
        if (!cell) return;
        const input = cell.querySelector('input[type="file"]');
        const preview = cell.querySelector('.photo-preview');
        const desc = cell.querySelector('.photo-description');
        if (input) input.value = '';
        resetPreview(preview);
        if (desc) {
            desc.value = '';
            dispatchInput(desc);
        }
    };

    window.createPhotoCell = function createPhotoCell(index) {
        const uniqueId = `${Date.now()}_${index}_${Math.random().toString(36).slice(2, 8)}`;
        return `
            <div class="photo-cell">
                <div class="photo-frame">
                    <button type="button" class="photo-delete-btn" aria-label="Hapus foto" onclick="clearPhotoCell(this)">×</button>
                    <div class="photo-upload-box" onclick="document.getElementById('photo${uniqueId}').click()">
                        <input type="file" id="photo${uniqueId}" accept="image/*" style="display:none"
                            onchange="handlePhotoCellUpload(this, 'preview${uniqueId}')">
                        <div id="preview${uniqueId}" class="photo-preview">
                            ${PHOTO_PLACEHOLDER_HTML}
                        </div>
                    </div>
                </div>
                <input type="text" id="desc${uniqueId}" class="photo-description" placeholder="Deskripsi foto...">
            </div>
        `;
    };

    if (typeof window.getFormData === 'function') {
        const originalGetFormData = window.getFormData;
        window.getFormData = function patchedGetFormData() {
            const data = originalGetFormData();
            data.findingEvidencePhotos = (data.findingEvidencePhotos || []).filter(isMeaningfulPhoto);
            data.correctiveActionPhotos = (data.correctiveActionPhotos || []).filter(isMeaningfulPhoto);
            if (!isMeaningfulPhoto(data.qscResultPhoto)) {
                data.qscResultPhoto = { image: null, description: '-' };
            }
            return data;
        };
    }

    function getPptxCtor() {
        if (typeof window.PptxGenJS !== 'undefined') return window.PptxGenJS;
        if (typeof window.pptxgen !== 'undefined') return window.pptxgen;
        if (typeof PptxGenJS !== 'undefined') return PptxGenJS;
        return null;
    }

    function setButtonBusy(button, isBusy, busyLabel) {
        if (!button) return;
        if (isBusy) {
            if (!button.dataset.originalHtml) {
                button.dataset.originalHtml = button.innerHTML;
            }
            button.disabled = true;
            button.innerHTML = `<span>⏳</span> ${busyLabel}`;
        } else {
            button.disabled = false;
            if (button.dataset.originalHtml) {
                button.innerHTML = button.dataset.originalHtml;
            }
        }
    }

    function ensureHeaderMaster(pptx, masterName, titleText) {
        if (definedMasters.has(masterName)) return;
        definedMasters.add(masterName);
        pptx.defineSlideMaster({
            title: masterName,
            margin: [0.25, 0.25, 0.25, 0.25],
            background: { color: PPT.WHITE },
            objects: [
                {
                    rect: {
                        x: 0,
                        y: 0,
                        w: PPT.W,
                        h: PPT.HEADER_H,
                        line: { color: PPT.TEAL, pt: 1 },
                        fill: { color: PPT.TEAL }
                    }
                },
                {
                    text: {
                        text: titleText,
                        options: {
                            x: 0.45,
                            y: 0.16,
                            w: PPT.W - 0.9,
                            h: 0.36,
                            margin: 0,
                            fontFace: 'Arial',
                            fontSize: 19,
                            bold: true,
                            color: PPT.WHITE,
                            fit: 'shrink'
                        }
                    }
                }
            ]
        });
    }

    function addFullBackground(slide, pptx, color) {
        slide.addShape(pptx.ShapeType.rect, {
            x: 0,
            y: 0,
            w: PPT.W,
            h: PPT.H,
            line: { color },
            fill: { color }
        });
    }

    function addTitleSlide(slide, pptx, title, subtitle) {
        addFullBackground(slide, pptx, PPT.WHITE);
        slide.addShape(pptx.ShapeType.rect, {
            x: 0.48,
            y: 2.15,
            w: 0.18,
            h: 1.55,
            line: { color: PPT.TEAL, pt: 1 },
            fill: { color: PPT.TEAL }
        });
        slide.addText(title, {
            x: 0.9,
            y: 2.25,
            w: 11.8,
            h: 0.62,
            margin: 0,
            fontFace: 'Arial',
            fontSize: 28,
            bold: true,
            color: PPT.TEAL,
            fit: 'shrink'
        });
        slide.addText(subtitle, {
            x: 0.9,
            y: 3.0,
            w: 11.4,
            h: 0.4,
            margin: 0,
            fontFace: 'Arial',
            fontSize: 14,
            color: PPT.MUTED,
            fit: 'shrink'
        });
    }

    function addEmptyStateSlide(slide, pptx, masterName, message) {
        addFullBackground(slide, pptx, PPT.WHITE);
        slide.addText(message, {
            x: 0.8,
            y: 3.0,
            w: PPT.W - 1.6,
            h: 0.6,
            margin: 0,
            fontFace: 'Arial',
            fontSize: 20,
            bold: true,
            color: PPT.MUTED,
            align: 'center',
            valign: 'mid',
            fit: 'shrink'
        });
    }

    function createTempTableHost(headers, rows, widths, options = {}) {
        const host = document.createElement('div');
        const table = document.createElement('table');
        const colgroup = document.createElement('colgroup');
        const thead = document.createElement('thead');
        const tbody = document.createElement('tbody');
        const tableId = `ppt-table-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;

        host.style.position = 'fixed';
        host.style.left = '-20000px';
        host.style.top = '0';
        host.style.width = `${options.widthPx || 1600}px`;
        host.style.background = '#ffffff';
        host.style.padding = '0';
        host.style.zIndex = '-1';
        host.style.pointerEvents = 'none';

        table.id = tableId;
        table.style.width = '100%';
        table.style.borderCollapse = 'collapse';
        table.style.tableLayout = 'fixed';
        table.style.fontFamily = 'Arial, sans-serif';
        table.style.background = '#ffffff';

        widths.forEach((width) => {
            const col = document.createElement('col');
            col.style.width = `${width * 100}%`;
            colgroup.appendChild(col);
        });

        const trHead = document.createElement('tr');
        headers.forEach((headerText) => {
            const th = document.createElement('th');
            th.textContent = headerText;
            th.style.padding = options.headerPadding || '12px 10px';
            th.style.background = '#1B8F96';
            th.style.color = '#ffffff';
            th.style.border = '1px solid #1B8F96';
            th.style.fontSize = options.headerFontSize || '15px';
            th.style.fontWeight = '700';
            th.style.lineHeight = '1.25';
            th.style.textAlign = 'left';
            th.style.verticalAlign = 'middle';
            trHead.appendChild(th);
        });
        thead.appendChild(trHead);

        rows.forEach((row, rowIndex) => {
            const tr = document.createElement('tr');
            row.forEach((value, colIndex) => {
                const td = document.createElement('td');
                td.innerText = normalizeText(value, '—');
                td.style.padding = options.cellPadding || '10px 10px';
                td.style.border = '1px solid #D7E3F0';
                td.style.fontSize = options.bodyFontSize || '14px';
                td.style.color = '#132238';
                td.style.lineHeight = '1.3';
                td.style.whiteSpace = 'pre-wrap';
                td.style.wordBreak = 'break-word';
                td.style.textAlign = (options.centerColumns || []).includes(colIndex) ? 'center' : 'left';
                td.style.verticalAlign = 'top';
                td.style.background = rowIndex % 2 === 0 ? '#FFFFFF' : '#F5F7FA';
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        });

        table.appendChild(colgroup);
        table.appendChild(thead);
        table.appendChild(tbody);
        host.appendChild(table);
        document.body.appendChild(host);

        return {
            tableId,
            cleanup() {
                try {
                    host.remove();
                } catch (err) {
                    /* no-op */
                }
            }
        };
    }

    async function toDataUrl(src) {
        if (!src) return null;
        if (String(src).startsWith('data:')) return src;
        return new Promise((resolve) => {
            const img = new Image();
            img.crossOrigin = 'anonymous';
            img.onload = () => {
                try {
                    const canvas = document.createElement('canvas');
                    canvas.width = img.naturalWidth || img.width;
                    canvas.height = img.naturalHeight || img.height;
                    const ctx = canvas.getContext('2d');
                    ctx.drawImage(img, 0, 0);
                    resolve(canvas.toDataURL('image/png'));
                } catch (err) {
                    resolve(null);
                }
            };
            img.onerror = () => resolve(null);
            img.src = src;
        });
    }

    async function getContainedPlacement(src, x, y, w, h, pad = 0.05) {
        return new Promise((resolve) => {
            const img = new Image();
            img.onload = () => {
                const boxW = Math.max(0.01, w - pad * 2);
                const boxH = Math.max(0.01, h - pad * 2);
                const ar = (img.naturalWidth || img.width || 1) / (img.naturalHeight || img.height || 1);
                let drawW = boxW;
                let drawH = boxH;
                if (ar > boxW / boxH) {
                    drawH = boxW / ar;
                } else {
                    drawW = boxH * ar;
                }
                resolve({
                    x: x + (w - drawW) / 2,
                    y: y + (h - drawH) / 2,
                    w: drawW,
                    h: drawH
                });
            };
            img.onerror = () => resolve({ x: x + pad, y: y + pad, w: w - pad * 2, h: h - pad * 2 });
            img.src = src;
        });
    }

    async function addCoverSlide(pptx, data, storeDetail) {
        const slide = pptx.addSlide();
        addFullBackground(slide, pptx, PPT.WHITE);

        const heroImage = await toDataUrl('storefoto2.png');
        if (heroImage) {
            slide.addImage({ data: heroImage, x: 0, y: 0, w: PPT.W, h: 3.25 });
        }

        slide.addShape(pptx.ShapeType.rect, {
            x: 0,
            y: 3.25,
            w: PPT.W,
            h: PPT.H - 3.25,
            line: { color: 'F2F4F7', pt: 1 },
            fill: { color: 'F5F5F5' }
        });

        slide.addText('Regional Bestie Visit Report', {
            x: 0.55,
            y: 4.05,
            w: 8.9,
            h: 0.62,
            margin: 0,
            fontFace: 'Arial',
            fontSize: 28,
            bold: true,
            color: PPT.TEAL,
            fit: 'shrink'
        });

        const chips = [
            `Nama: ${normalizeText(data.nama)}`,
            `Store: ${normalizeText(data.store)}`,
            `Kode Store: ${normalizeText(storeDetail?.code || '-')}`,
            `Tanggal: ${typeof window.formatDateLong === 'function' ? window.formatDateLong(data.tanggal) : normalizeText(data.tanggal)}`
        ];

        let chipY = 4.95;
        chips.forEach((text, index) => {
            const chipW = index % 2 === 0 ? 5.8 : 4.6;
            const chipX = index % 2 === 0 ? 0.55 : 6.55;
            slide.addShape(pptx.ShapeType.roundRect, {
                x: chipX,
                y: chipY,
                w: chipW,
                h: 0.55,
                rectRadius: 0.08,
                line: { color: index < 2 ? PPT.BORDER : 'D8E9EA', pt: 1 },
                fill: { color: index < 2 ? PPT.WHITE : 'EEF8F8' }
            });
            slide.addText(text, {
                x: chipX + 0.14,
                y: chipY + 0.11,
                w: chipW - 0.28,
                h: 0.28,
                margin: 0,
                fontFace: 'Arial',
                fontSize: 13,
                bold: index < 2,
                color: PPT.TEXT,
                fit: 'shrink'
            });
            if (index % 2 === 1) chipY += 0.72;
        });
    }

    function addGeneralInfoSlides(pptx, data, storeDetail) {
        ensureHeaderMaster(pptx, 'master-general-information', 'GENERAL INFORMATION');

        const rows = [
            ['Store Leader', normalizeText(data.storeLeader), normalizeText(data.storeLeaderLevel)],
            ['Shift Leader', normalizeText(data.shiftLeader), normalizeText(data.shiftLeaderLevel)]
        ];
        (data.crewList || []).forEach((crew, index) => {
            rows.push([
                `Crew ${index + 1}`,
                normalizeText(crew.name),
                normalizeText(crew.level)
            ]);
        });

        const tableHost = createTempTableHost(
            ['Role', 'Name', 'Job Level'],
            rows,
            [0.24, 0.56, 0.20],
            { widthPx: 1400, headerFontSize: '16px', bodyFontSize: '15px', centerColumns: [2] }
        );

        try {
            const infoLines = [
                `Hari, Tanggal : ${typeof window.formatDateLong === 'function' ? window.formatDateLong(data.tanggal) : normalizeText(data.tanggal)}`,
                `Kode Store : ${normalizeText(storeDetail?.code || '-')}`,
            ].join('\n');

            pptx.tableToSlides(tableHost.tableId, {
                masterSlideName: 'master-general-information',
                x: 0.42,
                y: 1.55,
                w: PPT.W - 0.84,
                slideMargin: [1.55, 0.42, 0.38, 0.42],
                autoPageSlideStartY: 1.55,
                autoPageCharWeight: -0.05,
                autoPageLineWeight: 0.1,
                addText: {
                    text: infoLines,
                    options: {
                        x: 0.48,
                        y: 0.95,
                        w: 6.1,
                        h: 0.66,
                        margin: 0,
                        fontFace: 'Arial',
                        fontSize: 10,
                        color: PPT.TEXT,
                        breakLine: false
                    }
                }
            });
        } finally {
            tableHost.cleanup();
        }
    }

    function addObservationSlides(pptx, rows, opts) {
        const { title, subtitle, masterName, masterTitle } = opts;
        const titleSlide = pptx.addSlide();
        addTitleSlide(titleSlide, pptx, title, subtitle);

        if (!(rows || []).length) {
            ensureHeaderMaster(pptx, masterName, masterTitle);
            const emptySlide = pptx.addSlide({ masterName });
            emptySlide.addText('Belum ada data observasi.', {
                x: 0.8,
                y: 3.0,
                w: PPT.W - 1.6,
                h: 0.5,
                margin: 0,
                fontFace: 'Arial',
                fontSize: 20,
                bold: true,
                color: PPT.MUTED,
                align: 'center',
                fit: 'shrink'
            });
            return;
        }

        ensureHeaderMaster(pptx, masterName, masterTitle);
        const tableRows = rows.map((row) => [
            normalizeText(row.temuan),
            normalizeText(row.dampak),
            normalizeText(row.penyebab),
            normalizeText(row.tindakan),
            row.deadline && row.deadline !== '-' && typeof window.formatDateLong === 'function' ? window.formatDateLong(row.deadline) : normalizeText(row.deadline),
            normalizeText(row.hasil)
        ]);

        const tableHost = createTempTableHost(
            ['Temuan', 'Dampak', 'Penyebab', 'Tindakan Perbaikan', 'Tanggal Perbaikan / Deadline', 'Hasil'],
            tableRows,
            [0.16, 0.16, 0.16, 0.22, 0.14, 0.16],
            { widthPx: 1700, headerFontSize: '15px', bodyFontSize: '14px', centerColumns: [4] }
        );

        try {
            pptx.tableToSlides(tableHost.tableId, {
                masterSlideName: masterName,
                x: 0.32,
                y: 1.02,
                w: PPT.W - 0.64,
                slideMargin: [1.02, 0.32, 0.32, 0.32],
                autoPageSlideStartY: 1.02,
                autoPageCharWeight: -0.05,
                autoPageLineWeight: 0.12
            });
        } finally {
            tableHost.cleanup();
        }
    }

    async function addQscResultSlides(pptx, data) {
        ensureHeaderMaster(pptx, 'master-qsc-result', 'QSC / FAMITRACK RESULT');
        const slide = pptx.addSlide({ masterName: 'master-qsc-result' });
        slide.addShape(pptx.ShapeType.rect, {
            x: 0.42,
            y: 1.02,
            w: PPT.W - 0.84,
            h: PPT.H - 1.46,
            line: { color: PPT.TEAL, pt: 1.1 },
            fill: { color: 'FBFCFE' }
        });

        if (data.qscResultPhoto?.image) {
            const imageData = await toDataUrl(data.qscResultPhoto.image);
            if (imageData) {
                const placement = await getContainedPlacement(imageData, 0.46, 1.06, PPT.W - 0.92, PPT.H - 1.62, 0.06);
                slide.addImage({ data: imageData, ...placement });
            }
        } else {
            slide.addText('Belum ada foto QSC.', {
                x: 0.8,
                y: 3.0,
                w: PPT.W - 1.6,
                h: 0.5,
                margin: 0,
                fontFace: 'Arial',
                fontSize: 20,
                bold: true,
                color: PPT.MUTED,
                align: 'center',
                fit: 'shrink'
            });
        }

        const desc = normalizeText(data.qscResultPhoto?.description || '-', '');
        if (desc) {
            slide.addShape(pptx.ShapeType.roundRect, {
                x: 0.62,
                y: PPT.H - 0.78,
                w: PPT.W - 1.24,
                h: 0.34,
                line: { color: 'DCE7EF', pt: 1 },
                fill: { color: 'F5F8FC' }
            });
            slide.addText(desc, {
                x: 0.78,
                y: PPT.H - 0.71,
                w: PPT.W - 1.56,
                h: 0.18,
                margin: 0,
                fontFace: 'Arial',
                fontSize: 10.5,
                color: PPT.TEXT,
                align: 'center',
                fit: 'shrink'
            });
        }
    }

    function chunkPhotos(photos, size) {
        const items = (photos || []).filter(isMeaningfulPhoto);
        if (!items.length) return [[]];
        const chunks = [];
        for (let i = 0; i < items.length; i += size) {
            chunks.push(items.slice(i, i + size));
        }
        return chunks;
    }

    async function addPhotoGridSlides(pptx, photos, opts) {
        const titleSlide = pptx.addSlide();
        addTitleSlide(titleSlide, pptx, opts.title, opts.subtitle);

        ensureHeaderMaster(pptx, opts.masterName, opts.masterTitle);
        const chunks = chunkPhotos(photos, 8);

        for (const chunk of chunks) {
            const slide = pptx.addSlide({ masterName: opts.masterName });
            if (!chunk.length) {
                slide.addText('Belum ada foto yang diunggah.', {
                    x: 0.8,
                    y: 3.0,
                    w: PPT.W - 1.6,
                    h: 0.5,
                    margin: 0,
                    fontFace: 'Arial',
                    fontSize: 20,
                    bold: true,
                    color: PPT.MUTED,
                    align: 'center',
                    fit: 'shrink'
                });
                continue;
            }

            const cols = 4;
            const rows = 2;
            const areaX = 0.34;
            const areaY = 1.04;
            const areaW = PPT.W - 0.68;
            const areaH = PPT.H - 1.36;
            const gap = PPT.GAP;
            const cellW = (areaW - gap * (cols - 1)) / cols;
            const cellH = (areaH - gap * (rows - 1)) / rows;
            const photoH = cellH - 0.55;

            for (let index = 0; index < cols * rows; index++) {
                const col = index % cols;
                const row = Math.floor(index / cols);
                const x = areaX + col * (cellW + gap);
                const y = areaY + row * (cellH + gap);
                const item = chunk[index];

                slide.addShape(pptx.ShapeType.rect, {
                    x,
                    y,
                    w: cellW,
                    h: cellH,
                    line: { color: PPT.BORDER, pt: 1 },
                    fill: { color: PPT.WHITE }
                });
                slide.addShape(pptx.ShapeType.rect, {
                    x,
                    y,
                    w: cellW,
                    h: photoH,
                    line: { color: 'E8EEF5', pt: 1 },
                    fill: { color: 'FBFCFE' }
                });
                slide.addShape(pptx.ShapeType.rect, {
                    x,
                    y: y + photoH,
                    w: cellW,
                    h: cellH - photoH,
                    line: { color: 'E8EEF5', pt: 1 },
                    fill: { color: 'F5F7FA' }
                });

                if (item?.image) {
                    const imageData = await toDataUrl(item.image);
                    if (imageData) {
                        const placement = await getContainedPlacement(imageData, x + 0.03, y + 0.03, cellW - 0.06, photoH - 0.06, 0.04);
                        slide.addImage({ data: imageData, ...placement });
                    }
                } else {
                    slide.addText('📷', {
                        x: x + cellW / 2 - 0.2,
                        y: y + photoH / 2 - 0.18,
                        w: 0.4,
                        h: 0.3,
                        margin: 0,
                        fontFace: 'Arial',
                        fontSize: 18,
                        color: 'C5CED8',
                        align: 'center'
                    });
                }

                slide.addText(normalizeText(item?.description || '-', '—'), {
                    x: x + 0.08,
                    y: y + photoH + 0.06,
                    w: cellW - 0.16,
                    h: cellH - photoH - 0.12,
                    margin: 0,
                    fontFace: 'Arial',
                    fontSize: 9.5,
                    color: PPT.TEXT,
                    fit: 'shrink',
                    valign: 'mid',
                    align: 'left'
                });
            }
        }
    }

    function addStoreAssignmentSlides(pptx, data) {
        const titleSlide = pptx.addSlide();
        addTitleSlide(titleSlide, pptx, 'Store Assignment', 'Corrective Action Purpose');

        const slide = pptx.addSlide();
        addFullBackground(slide, pptx, PPT.WHITE);
        slide.addShape(pptx.ShapeType.roundRect, {
            x: 0.45,
            y: 0.9,
            w: PPT.W - 0.9,
            h: 0.58,
            line: { color: PPT.TEAL, pt: 1 },
            fill: { color: PPT.TEAL }
        });
        slide.addText(normalizeText(data.storeAssignmentLink || '-', '-'), {
            x: 0.7,
            y: 1.05,
            w: PPT.W - 1.4,
            h: 0.18,
            margin: 0,
            fontFace: 'Arial',
            fontSize: 14,
            bold: true,
            color: PPT.WHITE,
            align: 'center',
            fit: 'shrink'
        });

        slide.addText('Mekanisme pelaporan', {
            x: 0.55,
            y: 1.9,
            w: 4.4,
            h: 0.34,
            margin: 0,
            fontFace: 'Arial',
            fontSize: 18,
            bold: true,
            color: PPT.TEXT
        });

        const stepsText = [
            '1. Unduh file pada link di atas.',
            '2. Tim store mengisi form tersebut berdasarkan temuan yang telah ditunjukkan pada file laporan ini.',
            '3. Tindakan perbaikan WAJIB dilakukan sebelum deadline yang diberikan oleh Regional Bestie.',
            '4. Form tindakan perbaikan yang telah dibuat WAJIB dikirimkan kembali via email dengan terusan:',
            '   a. Regional Manager',
            '   b. Area Manager',
            '   c. Regional Bestie',
            '   d. FMCU (Bu Sari, Pak Ami, Pak Aufar)'
        ].join('\n');

        slide.addShape(pptx.ShapeType.roundRect, {
            x: 0.52,
            y: 2.4,
            w: PPT.W - 1.04,
            h: 3.85,
            line: { color: PPT.BORDER, pt: 1 },
            fill: { color: 'FBFCFE' }
        });
        slide.addText(stepsText, {
            x: 0.82,
            y: 2.78,
            w: PPT.W - 1.64,
            h: 3.1,
            margin: 0,
            fontFace: 'Arial',
            fontSize: 16,
            color: PPT.TEXT,
            fit: 'shrink',
            breakLine: false
        });
    }

    window.downloadPPT = async function downloadPPT(triggerButton) {
        const storeInput = document.getElementById('store');
        if (typeof window.validateStoreSelection === 'function' && !window.validateStoreSelection(true)) {
            if (storeInput) {
                storeInput.reportValidity();
                storeInput.focus();
            }
            return;
        }

        const PptxCtor = getPptxCtor();
        if (!PptxCtor) {
            alert('Mesin PPT belum siap di halaman ini. Silakan refresh lalu coba lagi.');
            return;
        }

        setButtonBusy(triggerButton, true, 'Menyiapkan PPT...');

        try {
            const data = typeof window.getFormData === 'function' ? window.getFormData() : {};
            const storeResult = typeof window.resolveStoreName === 'function' ? window.resolveStoreName(data.store || '') : null;
            const storeDetail = storeResult?.detail || (typeof window.getStoreDetailByName === 'function' ? window.getStoreDetailByName(data.store || '') : null) || {};

            const pptx = new PptxCtor();
            pptx.layout = 'LAYOUT_WIDE';
            pptx.author = 'Regional Bestie Report System';
            pptx.company = 'Marketty';
            pptx.subject = 'Regional Bestie Visit Report';
            pptx.title = `Regional Bestie Visit Report - ${normalizeText(data.store, 'Store')}`;

            await addCoverSlide(pptx, data, storeDetail);
            addGeneralInfoSlides(pptx, data, storeDetail);

            if (data.showQSCResult) {
                await addQscResultSlides(pptx, data);
            }

            if (data.showOPITable) {
                addObservationSlides(pptx, data.opiData || [], {
                    title: 'OPI PROJECT OBSERVATION',
                    subtitle: 'Findings & Root Cause Analysis',
                    masterName: 'master-opi-observation',
                    masterTitle: 'OBSERVATION - Findings & Root Cause Analysis'
                });
            }

            if (data.showQSCTable) {
                addObservationSlides(pptx, data.qscData || [], {
                    title: 'QSC OBSERVATION',
                    subtitle: 'Findings & Root Cause Analysis',
                    masterName: 'master-qsc-observation',
                    masterTitle: 'QSC OBSERVATION - Findings & Root Cause Analysis'
                });
            }

            if (data.showFindingEvidence) {
                await addPhotoGridSlides(pptx, data.findingEvidencePhotos || [], {
                    title: 'FINDING EVIDENCE',
                    subtitle: 'of OPI & QSC Observation',
                    masterName: 'master-finding-evidence',
                    masterTitle: 'FINDING EVIDENCE of OPI & QSC Observation'
                });
            }

            if (data.showCorrectiveAction) {
                await addPhotoGridSlides(pptx, data.correctiveActionPhotos || [], {
                    title: 'Corrective Action',
                    subtitle: 'Evidence & Result by Regional Bestie',
                    masterName: 'master-corrective-action',
                    masterTitle: 'Corrective Action Evidence & Result by Regional Bestie'
                });
            }

            addStoreAssignmentSlides(pptx, data);

            const fileName = `${sanitizeFileName(`Regional_Bestie_Visit_Report_${data.store || 'Store'}`)}.pptx`;
            await pptx.writeFile({ fileName, compression: true });
        } catch (error) {
            console.error('PPT generation error:', error);
            alert('Gagal generate PPT. Silakan coba lagi.');
        } finally {
            setButtonBusy(triggerButton, false, 'Menyiapkan PPT...');
        }
    };
})();
