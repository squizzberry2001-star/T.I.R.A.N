(function () {
    const TEMPLATE_BASE64 = "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";
    const SHEET_PATH = 'xl/worksheets/sheet1.xml';

    function safeText(value) {
        if (value === undefined || value === null) return '';
        return String(value).replace(/\u00a0/g, ' ').trim();
    }

    function isMeaningful(value) {
        const text = safeText(value);
        return !!text && text !== '-';
    }

    function escapeXml(value) {
        return String(value === undefined || value === null ? '' : value)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&apos;');
    }

    function stripFormatMarkers(value) {
        let text = String(value || '');
        if (typeof window.stripObservationFormatMarkers === 'function') {
            text = window.stripObservationFormatMarkers(text);
        }
        if (/<[^>]+>/.test(text) && typeof document !== 'undefined') {
            const scratch = document.createElement('div');
            scratch.innerHTML = text.replace(/<br\s*\/?\s*>/gi, '\n').replace(/<\/p>|<\/div>|<\/li>/gi, '\n');
            text = scratch.textContent || scratch.innerText || '';
        } else if (/<[^>]+>/.test(text)) {
            text = text
                .replace(/<br\s*\/?\s*>/gi, '\n')
                .replace(/<\/p>|<\/div>|<\/li>/gi, '\n')
                .replace(/<[^>]+>/g, '');
        }
        return text
            .replace(/<u>([\s\S]*?)<\/u>/gi, '$1')
            .replace(/\*\*([\s\S]*?)\*\*/g, '$1')
            .replace(/(^|[^*])\*([^*\n]+?)\*/g, '$1$2')
            .replace(/\r/g, '')
            .replace(/[ \t]+\n/g, '\n')
            .replace(/\n{3,}/g, '\n\n')
            .trim();
    }

    function formatDateId(value) {
        const raw = safeText(value);
        if (!raw || raw === '-') return '';
        const date = new Date(raw);
        if (Number.isNaN(date.getTime())) return raw;
        const monthNames = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
        const day = String(date.getDate()).padStart(2, '0');
        const month = monthNames[date.getMonth()] || String(date.getMonth() + 1).padStart(2, '0');
        const year = String(date.getFullYear());
        return day + ' ' + month + ' ' + year;
    }

    function sanitizeFileName(value) {
        return String(value || 'CA_Store_Assignment')
            .replace(/[^\w\s-]/g, '')
            .trim()
            .replace(/\s+/g, '_') || 'CA_Store_Assignment';
    }

    function getStoreDetail(data) {
        if (typeof window.getStoreWebDetail === 'function') {
            try {
                return window.getStoreWebDetail(data.store || '') || {};
            } catch (error) {
                return {};
            }
        }
        return {};
    }

    function normalizeObservationRows(rows) {
        return (Array.isArray(rows) ? rows : [])
            .map(function (row) {
                return {
                    temuan: stripFormatMarkers(row && row.temuan),
                    deadline: safeText(row && row.deadline)
                };
            })
            .filter(function (row) {
                return isMeaningful(row.temuan);
            });
    }

    function composeFindingText(row) {
        return stripFormatMarkers(row && row.temuan);
    }

    function collectCAAssignmentRows(data) {
        const opiRows = normalizeObservationRows(data && data.opiData);
        const qscRows = normalizeObservationRows(data && data.qscData);
        return opiRows.concat(qscRows);
    }

    function cellBlank(ref, style) {
        return '<c r="' + ref + '" s="' + style + '"/>';
    }

    function cellShared(ref, style, index) {
        return '<c r="' + ref + '" s="' + style + '" t="s"><v>' + index + '</v></c>';
    }

    function cellInline(ref, style, value) {
        const text = safeText(value);
        if (!text) return cellBlank(ref, style);
        return '<c r="' + ref + '" s="' + style + '" t="inlineStr"><is><t xml:space="preserve">' + escapeXml(text) + '</t></is></c>';
    }

    function cellNumber(ref, style, value) {
        if (value === undefined || value === null || value === '') return cellBlank(ref, style);
        return '<c r="' + ref + '" s="' + style + '"><v>' + Number(value) + '</v></c>';
    }

    function rowXml(rowNumber, height, cells) {
        const span = rowNumber <= 32 ? ' spans="1:6"' : '';
        return '<row r="' + rowNumber + '"' + span + ' ht="' + height + '" customHeight="1" x14ac:dyDescent="0.2">' + cells.join('') + '</row>';
    }

    function estimateRowHeight(findingText) {
        const text = safeText(findingText);
        const manualLines = text ? text.split('\n').length : 1;
        const estimatedWrapLines = text ? Math.ceil(text.length / 48) : 1;
        const totalLines = Math.max(manualLines, estimatedWrapLines);
        const calculated = 54 + totalLines * 18;
        return Math.max(187.5, Math.min(360, calculated));
    }

    function buildHeaderRows(meta) {
        return [
            rowXml(1, 45.75, [cellShared('A1', 6, 0), cellBlank('B1', 7), cellBlank('C1', 7), cellBlank('D1', 7), cellBlank('E1', 7), cellBlank('F1', 8)]),
            rowXml(2, 15.75, [cellBlank('A2', 1), cellBlank('B2', 1), cellBlank('C2', 1), cellBlank('D2', 1), cellBlank('E2', 1)]),
            rowXml(3, 27.75, [cellShared('A3', 9, 1), cellBlank('B3', 7), cellBlank('C3', 8), cellInline('D3', 10, meta.storeCode), cellBlank('E3', 8)]),
            rowXml(4, 27.75, [cellShared('A4', 9, 2), cellBlank('B4', 7), cellBlank('C4', 8), cellInline('D4', 10, meta.storeName), cellBlank('E4', 8)]),
            rowXml(5, 27.75, [cellShared('A5', 9, 3), cellBlank('B5', 7), cellBlank('C5', 8), cellInline('D5', 10, meta.bestieName), cellBlank('E5', 8)]),
            rowXml(6, 27.75, [cellShared('A6', 9, 4), cellBlank('B6', 7), cellBlank('C6', 8), cellInline('D6', 10, meta.visitDate), cellBlank('E6', 8)]),
            rowXml(7, 27.75, [cellShared('A7', 9, 5), cellBlank('B7', 7), cellBlank('C7', 8), cellInline('D7', 10, meta.picStore), cellBlank('E7', 8)]),
            rowXml(8, 27.75, []),
            rowXml(9, 50.25, [cellShared('A9', 2, 6), cellShared('B9', 11, 7), cellBlank('C9', 7), cellBlank('D9', 8), cellShared('E9', 3, 8), cellShared('F9', 3, 9)])
        ];
    }

    function buildDataRows(rows) {
        const source = rows.length ? rows : [{ blank: true }];
        return source.map(function (row, index) {
            const excelRow = 10 + index;
            const findingText = row.blank ? '' : composeFindingText(row);
            const deadline = row.blank ? '' : formatDateId(row.deadline);
            const isFirst = index === 0;
            return rowXml(excelRow, estimateRowHeight(findingText), [
                cellNumber('A' + excelRow, isFirst ? 4 : 5, index + 1),
                cellInline('B' + excelRow, isFirst ? 12 : 13, findingText),
                cellBlank('C' + excelRow, 7),
                cellBlank('D' + excelRow, 8),
                cellInline('E' + excelRow, 14, deadline),
                cellBlank('F' + excelRow, 5)
            ]);
        });
    }

    function buildMergeCells(lastDataRow) {
        const refs = [
            'A5:C5', 'D5:E5', 'A6:C6', 'D6:E6', 'A7:C7', 'D7:E7',
            'A1:F1', 'A3:C3', 'D3:E3', 'A4:C4', 'D4:E4', 'B9:D9'
        ];
        for (let row = 10; row <= lastDataRow; row += 1) refs.push('B' + row + ':D' + row);
        return '<mergeCells count="' + refs.length + '">' + refs.map(function (ref) {
            return '<mergeCell ref="' + ref + '"/>';
        }).join('') + '</mergeCells>';
    }

    function createWorksheetXml(data) {
        const detail = getStoreDetail(data || {});
        const rows = collectCAAssignmentRows(data || {});
        const lastDataRow = 10 + Math.max(rows.length, 1) - 1;
        const meta = {
            storeCode: detail.siteCode4 || detail.siteCode || '',
            storeName: detail.siteDescr || (data && data.store) || '',
            bestieName: data && data.nama || '',
            visitDate: formatDateId(data && data.tanggal),
            picStore: data && data.storeLeader || detail.storeHead || ''
        };
        const sheetRows = buildHeaderRows(meta).concat(buildDataRows(rows)).join('');
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac xr xr2 xr3" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" xr:uid="{00000000-0001-0000-0000-000000000000}">' +
            '<sheetPr><outlinePr summaryBelow="0" summaryRight="0"/></sheetPr>' +
            '<dimension ref="A1:F' + lastDataRow + '"/>' +
            '<sheetViews><sheetView tabSelected="1" zoomScale="55" zoomScaleNormal="55" workbookViewId="0"><selection activeCell="B10" sqref="B10:D10"/></sheetView></sheetViews>' +
            '<sheetFormatPr defaultColWidth="0" defaultRowHeight="15.75" customHeight="1" x14ac:dyDescent="0.2"/>' +
            '<cols><col min="1" max="1" width="5.140625" customWidth="1"/><col min="2" max="2" width="3.85546875" customWidth="1"/><col min="3" max="3" width="12.42578125" customWidth="1"/><col min="4" max="4" width="22.42578125" customWidth="1"/><col min="5" max="5" width="19.7109375" customWidth="1"/><col min="6" max="6" width="91.5703125" customWidth="1"/><col min="7" max="16384" width="12.5703125" hidden="1"/></cols>' +
            '<sheetData>' + sheetRows + '</sheetData>' +
            buildMergeCells(lastDataRow) +
            '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>' +
            '</worksheet>';
    }


    function enhanceCAStylesXml(xml) {
        const findingStyleFirst = '<xf numFmtId="0" fontId="4" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>';
        const findingStyleNext = '<xf numFmtId="0" fontId="4" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"/>';
        const centeredWrappedFirst = '<xf numFmtId="0" fontId="4" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>';
        const centeredWrappedNext = '<xf numFmtId="0" fontId="4" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>';
        const deadlineBoldCenter = '<xf numFmtId="0" fontId="3" fillId="0" borderId="4" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>';
        let nextXml = safeText(xml)
            .replace(findingStyleFirst, centeredWrappedFirst)
            .replace(findingStyleNext, centeredWrappedNext);
        if (!nextXml.includes(deadlineBoldCenter)) {
            nextXml = nextXml.replace(/<cellXfs count="(\d+)">/, function (match, count) {
                return '<cellXfs count="' + (Number(count) + 1) + '">';
            }).replace('</cellXfs>', deadlineBoldCenter + '</cellXfs>');
        }
        return nextXml;
    }

    function saveBlob(blob, fileName) {
        const url = URL.createObjectURL(blob);
        const anchor = document.createElement('a');
        anchor.href = url;
        anchor.download = fileName;
        document.body.appendChild(anchor);
        anchor.click();
        anchor.remove();
        window.setTimeout(function () { URL.revokeObjectURL(url); }, 1200);
    }

    async function buildCAAssignmentWorkbook(data) {
        if (typeof JSZip === 'undefined') {
            throw new Error('Mesin Excel belum siap. File jszip.min.js tidak ditemukan.');
        }
        const zip = await JSZip.loadAsync(TEMPLATE_BASE64, { base64: true });
        const stylesFile = zip.file('xl/styles.xml');
        if (stylesFile) {
            zip.file('xl/styles.xml', enhanceCAStylesXml(await stylesFile.async('string')));
        }
        zip.file(SHEET_PATH, createWorksheetXml(data || {}));
        return zip.generateAsync({
            type: 'blob',
            mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            compression: 'DEFLATE'
        });
    }

    window.downloadCAAssignmentExcel = async function downloadCAAssignmentExcel(button) {
        const storeInput = document.getElementById('store');
        if (typeof window.validateStoreSelection === 'function' && !window.validateStoreSelection(true)) {
            if (storeInput) {
                storeInput.reportValidity();
                storeInput.focus();
            }
            return;
        }

        const originalText = button ? button.innerHTML : '';
        if (button) {
            button.disabled = true;
            button.innerHTML = '<span>⏳</span> Exporting Excel...';
        }

        try {
            const data = typeof window.getFormData === 'function' ? window.getFormData() : {};
            const rows = collectCAAssignmentRows(data);
            if (!rows.length) {
                const proceed = window.confirm('Data OPI/QSC belum terisi. Tetap export template CA Assignment kosong?');
                if (!proceed) return;
            }
            const blob = await buildCAAssignmentWorkbook(data);
            const fileName = sanitizeFileName('CA_Store_Assignment_' + (data.store || 'Store')) + '.xlsx';
            saveBlob(blob, fileName);
        } catch (error) {
            console.error('Export CA Assignment Excel gagal:', error);
            window.alert(error && error.message ? error.message : 'Gagal export Excel CA Assignment.');
        } finally {
            if (button) {
                button.disabled = false;
                button.innerHTML = originalText;
            }
        }
    };

    window.__caAssignmentExport = {
        createWorksheetXml: createWorksheetXml,
        collectRows: collectCAAssignmentRows,
        buildWorkbook: buildCAAssignmentWorkbook
    };
})();
