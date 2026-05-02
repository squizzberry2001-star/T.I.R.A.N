(function () {
  'use strict';

  const TEMPLATE = window.RBV_PDF_TEMPLATE_ASSETS || {};

  const PDF_SETTINGS_KEY = 'rbv_pdf_settings_v2';
  const DEFAULT_ASSIGNMENT_LINK = 'https://tinyurl.com/store-caassignment';
  const ASSIGNMENT_CONFIG_KEY = 'rbv_assignment_link_config_v1';
  const DEFAULT_PDF_SETTINGS = {
    tableFontSize: 9.4,
    tableTitleFontSize: 9.8,
    evidenceFontSize: 8.9,
    tableExtraRows: 0,
    photoGridPerPage: 6
  };

  function clampNumber(value, min, max, fallback) {
    const number = Number(value);
    if (!Number.isFinite(number)) return fallback;
    return Math.min(max, Math.max(min, number));
  }

  function normalizePdfPhotoGridPerPage(value, fallback) {
    const allowed = [4, 6, 8];
    const number = Number(value);
    if (!Number.isFinite(number)) return fallback || DEFAULT_PDF_SETTINGS.photoGridPerPage;
    return allowed.reduce(function (closest, item) { return Math.abs(item - number) < Math.abs(closest - number) ? item : closest; }, allowed.indexOf(fallback) !== -1 ? fallback : 6);
  }

  function readPdfSettings() {
    try {
      const raw = JSON.parse(localStorage.getItem(PDF_SETTINGS_KEY) || '{}') || {};
      return {
        tableFontSize: clampNumber(raw.tableFontSize, 8, 13, DEFAULT_PDF_SETTINGS.tableFontSize),
        tableTitleFontSize: clampNumber(raw.tableTitleFontSize, 8, 14, DEFAULT_PDF_SETTINGS.tableTitleFontSize),
        evidenceFontSize: clampNumber(raw.evidenceFontSize, 8, 12, DEFAULT_PDF_SETTINGS.evidenceFontSize),
        tableExtraRows: Math.round(clampNumber(raw.tableExtraRows, 0, 4, DEFAULT_PDF_SETTINGS.tableExtraRows)),
        photoGridPerPage: normalizePdfPhotoGridPerPage(raw.photoGridPerPage, DEFAULT_PDF_SETTINGS.photoGridPerPage)
      };
    } catch (error) {
      return Object.assign({}, DEFAULT_PDF_SETTINGS);
    }
  }

  function pdfTableFontSize() {
    return readPdfSettings().tableFontSize;
  }

  function pdfTableTitleFontSize() {
    return readPdfSettings().tableTitleFontSize;
  }

  function pdfFieldTitleGap() {
    const titleSize = pdfTableTitleFontSize();
    return Math.max(4.2, titleSize * 0.52);
  }

  function pdfEvidenceFontSize() {
    return readPdfSettings().evidenceFontSize;
  }

  function pdfTableExtraRows() {
    return readPdfSettings().tableExtraRows;
  }

  function pdfPhotoGridPerPage() {
    return readPdfSettings().photoGridPerPage;
  }

  function readAssignmentLinkConfig() {
    try {
      const parsed = JSON.parse(localStorage.getItem(ASSIGNMENT_CONFIG_KEY) || '{}') || {};
      return plainText(parsed.link, DEFAULT_ASSIGNMENT_LINK);
    } catch (error) {
      return DEFAULT_ASSIGNMENT_LINK;
    }
  }

  function getAssignmentLink(data) {
    const hiddenLink = readAssignmentLinkConfig();
    return hiddenLink || text(data && data.storeAssignmentLink, DEFAULT_ASSIGNMENT_LINK);
  }


  function plainText(value, fallback) {
    let raw = value === undefined || value === null ? '' : String(value);
    raw = raw
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<\/p>/gi, '\n')
      .replace(/<\/div>/gi, '\n')
      .replace(/<li[^>]*>/gi, '- ')
      .replace(/<\/li>/gi, '\n')
      .replace(/<[^>]+>/g, '')
      .replace(/&nbsp;/gi, ' ')
      .replace(/&amp;/gi, '&')
      .replace(/&lt;/gi, '<')
      .replace(/&gt;/gi, '>')
      .replace(/&quot;/gi, '"')
      .replace(/&#39;/gi, "'")
      .replace(/\r/g, '')
      .replace(/[ \t]+\n/g, '\n')
      .replace(/\n{3,}/g, '\n\n')
      .trim();
    return raw || (fallback === undefined ? '-' : fallback);
  }

  function text(value, fallback) {
    return plainText(value, fallback);
  }

  function sanitizeFileName(value) {
    return String(value || 'Regional_Bestie_Visit_Report')
      .replace(/[^\w\s-]/g, '')
      .trim()
      .replace(/\s+/g, '_') || 'Regional_Bestie_Visit_Report';
  }

  function formatDate(value) {
    const raw = text(value, '');
    if (!raw || raw === '-') return '-';
    const date = new Date(raw);
    if (Number.isNaN(date.getTime())) return raw;
    return date.toLocaleDateString('id-ID', { day: '2-digit', month: 'short', year: 'numeric' });
  }

  function buildFileName(data) {
    return sanitizeFileName('Regional_Bestie_Visit_Report_' + text(data && data.store, 'Store')) + '.pdf';
  }

  function getStoreDetail(storeName, data) {
    let detail = {};
    if (typeof window.getStoreWebDetail === 'function') {
      try { detail = window.getStoreWebDetail(storeName) || {}; } catch (error) { detail = {}; }
    }
    const manual = data && data.manualStoreDetail && typeof data.manualStoreDetail === 'object' ? data.manualStoreDetail : {};
    return Object.assign({}, detail, manual);
  }

  function addWrapped(doc, value, x, y, width, lineHeight, options) {
    const lines = doc.splitTextToSize(text(value), width);
    doc.text(lines, x, y, options || {});
    return y + Math.max(1, lines.length) * lineHeight;
  }

  function richTokens(value) {
    const raw = value === undefined || value === null ? '' : String(value);
    const tokens = [];
    function push(value, style) {
      const safe = String(value || '').replace(/\u00a0/g, ' ');
      if (safe) tokens.push(Object.assign({ text: safe }, style || {}));
    }
    function newline() { tokens.push({ text: '\n' }); }
    if (!/<[a-z][\s\S]*>/i.test(raw) || typeof document === 'undefined') {
      text(raw, '').split('\n').forEach(function (line, index) {
        if (index) newline();
        push(line, {});
      });
      return tokens.length ? tokens : [{ text: '-' }];
    }
    const root = document.createElement('div');
    root.innerHTML = raw;
    function walk(node, style, list) {
      if (node.nodeType === 3) { push(node.nodeValue, style); return; }
      if (node.nodeType !== 1) return;
      const tag = String(node.tagName || '').toLowerCase();
      if (tag === 'br') { newline(); return; }
      const nextStyle = Object.assign({}, style);
      if (tag === 'b' || tag === 'strong') nextStyle.bold = true;
      if (tag === 'i' || tag === 'em') nextStyle.italic = true;
      if (tag === 'u') nextStyle.underline = true;
      if (tag === 'ul' || tag === 'ol') {
        Array.from(node.children || []).forEach(function (child, index) { walk(child, nextStyle, { type: tag, index: index + 1 }); });
        newline();
        return;
      }
      if (tag === 'li') {
        if (tokens.length && tokens[tokens.length - 1].text !== '\n') newline();
        push(list && list.type === 'ol' ? String(list.index || 1) + '. ' : '• ', nextStyle);
        Array.from(node.childNodes || []).forEach(function (child) { walk(child, nextStyle, list); });
        newline();
        return;
      }
      const block = ['p', 'div'].indexOf(tag) !== -1;
      if (block && tokens.length && tokens[tokens.length - 1].text !== '\n') newline();
      Array.from(node.childNodes || []).forEach(function (child) { walk(child, nextStyle, list); });
      if (block) newline();
    }
    Array.from(root.childNodes || []).forEach(function (child) { walk(child, {}, null); });
    while (tokens.length && tokens[0].text === '\n') tokens.shift();
    while (tokens.length && tokens[tokens.length - 1].text === '\n') tokens.pop();
    return tokens.length ? tokens : [{ text: '-' }];
  }

  function setRichFont(doc, token, baseBold) {
    const bold = Boolean(baseBold || token.bold);
    const italic = Boolean(token.italic);
    const style = bold && italic ? 'bolditalic' : bold ? 'bold' : italic ? 'italic' : 'normal';
    try { doc.setFont('helvetica', style); } catch (error) { doc.setFont('helvetica', bold ? 'bold' : 'normal'); }
  }

  function measureRich(doc, token, baseBold) {
    setRichFont(doc, token, baseBold);
    return doc.getTextWidth(token.text || '');
  }

  function splitRichTextToLines(doc, value, width, baseBold) {
    const source = richTokens(value);
    const lines = [[]];
    function current() { return lines[lines.length - 1]; }
    function lineWidth(line) { return line.reduce(function (sum, token) { return sum + measureRich(doc, token, baseBold); }, 0); }
    function pushLine() { if (current().length || lines.length === 0) lines.push([]); }
    source.forEach(function (token) {
      String(token.text || '').replace(/\r/g, '').split(/(\n|\s+)/).forEach(function (piece) {
        if (!piece) return;
        if (piece === '\n') { pushLine(); return; }
        const isSpace = /^\s+$/.test(piece);
        if (isSpace) { if (!current().length) return; piece = ' '; }
        const next = Object.assign({}, token, { text: piece });
        if (!isSpace && current().length && lineWidth(current()) + measureRich(doc, next, baseBold) > width) pushLine();
        current().push(next);
      });
    });
    const clean = lines.filter(function (line) { return line.length; });
    return clean.length ? clean : [[{ text: '-' }]];
  }

  function richLineToText(line) {
    return (line || []).map(function (token) { return token.text || ''; }).join('').trim() || '-';
  }

  function drawRichLines(doc, lines, x, y, lineHeight, options) {
    const opts = options || {};
    const baseBold = Boolean(opts.bold);
    const textColor = opts.textColor || null;
    const fontSize = opts.fontSize || doc.getFontSize();
    doc.setFontSize(fontSize);
    if (textColor) doc.setTextColor.apply(doc, textColor);
    (lines && lines.length ? lines : [[{ text: '-' }]]).forEach(function (line, lineIndex) {
      let cx = x;
      const cy = y + lineIndex * lineHeight;
      line.forEach(function (token) {
        const chunk = token.text || '';
        if (!chunk) return;
        setRichFont(doc, token, baseBold);
        if (textColor) doc.setTextColor.apply(doc, textColor);
        doc.text(chunk, cx, cy, opts.textOptions || {});
        const w = doc.getTextWidth(chunk);
        if (token.underline) {
          doc.setDrawColor.apply(doc, textColor || [39, 39, 42]);
          doc.setLineWidth(0.18);
          doc.line(cx, cy + 0.8, cx + w, cy + 0.8);
        }
        cx += w;
      });
    });
  }


  function fitRichLinesToBox(doc, lines, maxLines, width, baseBold) {
    const sourceLines = lines && lines.length ? lines : [[{ text: '-' }]];
    const source = sourceLines.slice(0, Math.max(1, maxLines));
    if (!source.length) return [[{ text: '-' }]];
    const hasMore = sourceLines.length > source.length;
    if (!hasMore) return source;
    const last = (source[source.length - 1] || []).map(function (token) { return Object.assign({}, token); });
    function lineWidth(line) {
      return (line || []).reduce(function (sum, token) { return sum + measureRich(doc, token, baseBold); }, 0);
    }
    if (!last.length) {
      source[source.length - 1] = [{ text: '…' }];
      return source;
    }
    let lastToken = last[last.length - 1];
    lastToken.text = String(lastToken.text || '').replace(/\s+$/g, '');
    while (last.length && lineWidth(last) + doc.getTextWidth('…') > width) {
      lastToken = last[last.length - 1];
      const txt = String(lastToken.text || '');
      if (txt.length > 1) {
        lastToken.text = txt.slice(0, -1).replace(/\s+$/g, '');
      } else {
        last.pop();
      }
    }
    if (!last.length) last.push({ text: '…' });
    else last.push({ text: '…' });
    source[source.length - 1] = last;
    return source;
  }

  function drawRichLinesInBox(doc, lines, x, y, width, height, lineHeight, options) {
    const opts = options || {};
    const fontSize = opts.fontSize || doc.getFontSize();
    doc.setFontSize(fontSize);
    const usableHeight = Math.max(0, height || 0);
    const maxLines = Math.max(1, Math.floor(usableHeight / Math.max(1, lineHeight)));
    const fitted = fitRichLinesToBox(doc, lines, maxLines, width, Boolean(opts.bold));
    drawRichLines(doc, fitted, x, y, lineHeight, opts);
  }

  async function prepareImage(src) {
    return new Promise(function (resolve) {
      if (!src) return resolve(null);
      const img = new Image();
      img.onload = function () {
        try {
          const canvas = document.createElement('canvas');
          const maxSide = 2200;
          let width = img.naturalWidth || img.width || 1;
          let height = img.naturalHeight || img.height || 1;
          const ratio = Math.min(1, maxSide / Math.max(width, height));
          width = Math.max(1, Math.round(width * ratio));
          height = Math.max(1, Math.round(height * ratio));
          canvas.width = width;
          canvas.height = height;
          const context = canvas.getContext('2d');
          context.drawImage(img, 0, 0, width, height);
          resolve({ dataUrl: canvas.toDataURL('image/jpeg', 0.88), width: width, height: height, canvas: canvas });
        } catch (error) { resolve(null); }
      };
      img.onerror = function () { resolve(null); };
      img.crossOrigin = 'anonymous';
      img.src = src;
    });
  }

  async function addImageInBox(doc, src, x, y, width, height, mode) {
    const prepared = await prepareImage(src);
    if (!prepared) return false;

    if (mode === 'coverCrop') {
      try {
        const targetAspect = width / height;
        let sx = 0;
        let sy = 0;
        let sw = prepared.width;
        let sh = prepared.height;
        const sourceAspect = prepared.width / prepared.height;
        if (sourceAspect > targetAspect) {
          sw = prepared.height * targetAspect;
          sx = (prepared.width - sw) / 2;
        } else {
          sh = prepared.width / targetAspect;
          sy = (prepared.height - sh) / 2;
        }
        const canvas = document.createElement('canvas');
        const outW = Math.max(480, Math.round(width * 8));
        const outH = Math.max(480, Math.round(height * 8));
        canvas.width = outW;
        canvas.height = outH;
        const context = canvas.getContext('2d');
        context.imageSmoothingEnabled = true;
        context.imageSmoothingQuality = 'high';
        context.drawImage(prepared.canvas, sx, sy, sw, sh, 0, 0, outW, outH);
        doc.addImage(canvas.toDataURL('image/jpeg', 0.9), 'JPEG', x, y, width, height);
        return true;
      } catch (error) {
        doc.addImage(prepared.dataUrl, 'JPEG', x, y, width, height);
        return true;
      }
    }

    const aspect = prepared.width / prepared.height;
    let targetWidth = width;
    let targetHeight = height;
    if (mode === 'cover') {
      if (aspect > targetWidth / targetHeight) targetWidth = targetHeight * aspect;
      else targetHeight = targetWidth / aspect;
    } else if (aspect > targetWidth / targetHeight) {
      targetHeight = targetWidth / aspect;
    } else {
      targetWidth = targetHeight * aspect;
    }
    const tx = x + (width - targetWidth) / 2;
    const ty = y + (height - targetHeight) / 2;
    doc.addImage(prepared.dataUrl, 'JPEG', tx, ty, targetWidth, targetHeight);
    return true;
  }

  function addBackground(doc, key, pageWidth, pageHeight) {
    if (!TEMPLATE[key]) return false;
    try {
      doc.addImage(TEMPLATE[key], 'JPEG', 0, 0, pageWidth, pageHeight);
      return true;
    } catch (error) {
      return false;
    }
  }

  function normalizeRows(rows) {
    return (Array.isArray(rows) ? rows : [])
      .map(function (row) {
        return {
          temuan: row && row.temuan !== undefined && row.temuan !== null ? String(row.temuan) : '',
          kondisiIdeal: row && row.kondisiIdeal !== undefined && row.kondisiIdeal !== null ? String(row.kondisiIdeal) : '',
          dampak: row && row.dampak !== undefined && row.dampak !== null ? String(row.dampak) : '',
          penyebab: row && row.penyebab !== undefined && row.penyebab !== null ? String(row.penyebab) : '',
          tindakan: row && row.tindakan !== undefined && row.tindakan !== null ? String(row.tindakan) : '',
          deadline: text(row && row.deadline, ''),
          hasil: row && row.hasil !== undefined && row.hasil !== null ? String(row.hasil) : ''
        };
      })
      .filter(function (row) {
        return Object.keys(row).some(function (key) {
          const value = text(row[key], '');
          return value && value !== '-';
        });
      });
  }

  function normalizePhotos(photos) {
    return (Array.isArray(photos) ? photos : [])
      .map(function (photo) {
        return { image: photo && photo.image ? photo.image : '', description: photo && photo.description !== undefined && photo.description !== null ? String(photo.description) : '' };
      })
      .filter(function (photo) { return photo.image || photo.description; });
  }

  function normalizeQscPhotos(data) {
    const modern = Array.isArray(data && data.qscResultPhotos) ? data.qscResultPhotos : [];
    const legacy = data && data.qscResultPhoto ? [data.qscResultPhoto] : [];
    const source = modern.length ? modern : legacy;
    return [0, 1].map(function (index) {
      const item = source[index] || {};
      return { image: item.image || '', description: item.description !== undefined && item.description !== null ? String(item.description) : '' };
    });
  }

  function drawTopBar(doc, title, palette, pageWidth) {
    doc.setFillColor.apply(doc, palette.primary);
    doc.rect(0, 0, pageWidth, 18, 'F');
    doc.setTextColor(255, 255, 255);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(12.5);
    doc.text(title, 12, 11.8);
  }

  function drawSubtlePageTitle(doc, title, subtitle, palette, pageWidth, margin) {
    doc.setTextColor.apply(doc, palette.primary);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(24);
    doc.text(title, margin, 24);
    if (subtitle) {
      doc.setTextColor(82, 82, 91);
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(10.5);
      doc.text(subtitle, margin, 32);
    }
    doc.setDrawColor.apply(doc, palette.primary);
    doc.setLineWidth(0.4);
    doc.line(margin, 37, pageWidth - margin, 37);
  }

  function addFooter(doc, pageWidth, pageHeight) {
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(11.5);
    doc.setTextColor(148, 163, 184);
    doc.text('GENERATE BY BESTIE VISIT WEB REPORT', pageWidth / 2, pageHeight - 4.8, { align: 'center' });
  }

  function applyFooterAllPages(doc, pageWidth, pageHeight) {
    const totalPages = typeof doc.internal.getNumberOfPages === 'function' ? doc.internal.getNumberOfPages() : 1;
    for (let page = 1; page <= totalPages; page += 1) {
      doc.setPage(page);
      addFooter(doc, pageWidth, pageHeight);
    }
  }

  function drawLabelValue(doc, label, value, x, y, labelWidth, maxWidth, palette, options) {
    const opts = options || {};
    doc.setFont('helvetica', opts.boldLabel ? 'bold' : 'normal');
    doc.setFontSize(opts.fontSize || 12.5);
    doc.setTextColor(82, 82, 91);
    doc.text(label + ' :', x, y);
    doc.setFont('helvetica', opts.boldValue ? 'bold' : 'normal');
    doc.setTextColor.apply(doc, opts.valueColor || palette.ink);
    doc.text(doc.splitTextToSize(text(value), maxWidth), x + labelWidth, y);
  }

  function drawCover(doc, data, detail, palette, pageWidth, pageHeight, margin) {
    if (!addBackground(doc, 'cover', pageWidth, pageHeight)) {
      doc.setFillColor(255, 255, 255);
      doc.rect(0, 0, pageWidth, pageHeight, 'F');
      doc.setFillColor(229, 247, 244);
      doc.rect(0, 0, pageWidth, 102, 'F');
    }

    doc.setFillColor(255, 255, 255);
    doc.rect(0, 103, pageWidth, pageHeight - 103, 'F');

    doc.setTextColor.apply(doc, palette.primary);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(26);
    doc.text('Regional Bestie Visit Report', margin + 2, 126);

    const storeCode = detail.siteCode4 || detail.siteCode || detail.storeCode || data.storeCode || '-';
    const storeHead = text(data && data.storeLeader, '-');

    const leftX = margin + 2;
    let y = 144;
    drawLabelValue(doc, 'Nama', data && data.nama, leftX, y, 31, 120, palette, { fontSize: 12.5, boldValue: true });
    y += 9;
    drawLabelValue(doc, 'Store', data && data.store, leftX, y, 31, 120, palette, { fontSize: 12.5, boldValue: true });
    y += 9;
    drawLabelValue(doc, 'Store Head', storeHead, leftX, y, 31, 120, palette, { fontSize: 11.5, boldValue: true });

    const cardX = 180;
    const cardY = 134;
    const cardW = pageWidth - cardX - margin;
    const cardH = 34;
    doc.setFillColor(248, 250, 252);
    doc.setDrawColor(226, 232, 240);
    doc.roundedRect(cardX, cardY, cardW, cardH, 3, 3, 'FD');

    const items = [
      ['Tanggal Visit', formatDate(data && data.tanggal)],
      ['Kode Store', storeCode],
      ['Area Manager', detail.areaManager || data.areaManager || '-'],
      ['Regional Manager', detail.regionalManager || data.regionalManager || '-']
    ];
    items.forEach(function (item, index) {
      const col = index % 2;
      const row = Math.floor(index / 2);
      const x = cardX + 8 + col * (cardW / 2);
      const yy = cardY + 9 + row * 12;
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(7.7);
      doc.setTextColor(100, 116, 139);
      doc.text(item[0].toUpperCase(), x, yy);
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(clampNumber(pdfTableFontSize() - 0.2, 8.8, 11.6, 9.2));
      doc.setTextColor.apply(doc, palette.ink);
      doc.text(doc.splitTextToSize(text(item[1]), cardW / 2 - 12).slice(0, 1), x, yy + 5);
    });

    // Alamat sengaja tidak ditampilkan pada hasil PDF.
  }

  async function drawQscResultSlide(doc, data, palette, pageWidth, pageHeight, margin) {
    if (data && data.showQSCResult !== true) return;
    doc.addPage();
    doc.setFillColor(255, 255, 255);
    doc.rect(0, 0, pageWidth, pageHeight, 'F');
    drawSubtlePageTitle(doc, 'QSC/ Famitrack Result', '', palette, pageWidth, margin);

    const photos = normalizeQscPhotos(data || {});
    const gap = 10;
    const cardY = 42;
    const cardW = (pageWidth - margin * 2 - gap) / 2;
    const imageH = Math.round((cardW - 6) * 3 / 4);
    const cardH = imageH + 15;

    for (let index = 0; index < 2; index += 1) {
      const photo = photos[index];
      const x = margin + index * (cardW + gap);
      doc.setFillColor(255, 255, 255);
      doc.setDrawColor(203, 213, 225);
      doc.roundedRect(x, cardY, cardW, cardH, 4, 4, 'FD');
      const added = await addImageInBox(doc, photo.image, x + 3, cardY + 3, cardW - 6, imageH, 'contain');
      if (!added) {
        doc.setFillColor(248, 250, 252);
        doc.roundedRect(x + 3, cardY + 3, cardW - 6, imageH, 3, 3, 'F');
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(13);
        doc.setTextColor(148, 163, 184);
        doc.text('Belum ada foto', x + cardW / 2, cardY + 60, { align: 'center' });
      }
      doc.setTextColor(15, 23, 42);
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(8.5);
      const lines = doc.splitTextToSize(text(photo.description, index === 0 ? 'Foto QSC / Famitrack 1' : 'Foto QSC / Famitrack 2'), cardW - 10);
      doc.text(lines.slice(0, 2), x + 5, cardY + imageH + 8);
    }
  }

  function drawCrewSlide(doc, data, palette, pageWidth, pageHeight, margin) {
    doc.addPage();
    doc.setFillColor(255, 255, 255);
    doc.rect(0, 0, pageWidth, pageHeight, 'F');
    doc.setTextColor.apply(doc, palette.ink);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(19);
    doc.text('General Information', margin, 25);

    doc.setFont('helvetica', 'normal');
    doc.setFontSize(12.5);
    drawLabelValue(doc, 'Hari, Tanggal', formatDate(data && data.tanggal), margin, 42, 40, 150, palette, { fontSize: 12.5, boldValue: true });

    const crew = Array.isArray(data && data.crewList) ? data.crewList.filter(function (item) {
      return text(item && item.name, '') || text(item && item.level, '');
    }) : [];
    const rows = [
      ['Store Leader', text(data && data.storeLeader, '-'), text(data && data.storeLeaderLevel, '-')],
      ['Shift Leader', text(data && data.shiftLeader, '-'), text(data && data.shiftLeaderLevel, '-')]
    ];
    (crew.length ? crew : [{ name: '-', level: '-' }]).forEach(function (item, index) {
      rows.push(['Crew Store ' + String(index + 1), text(item.name, '-'), text(item.level, '-')]);
    });

    doc.autoTable({
      startY: 52,
      head: [['Role', 'Nama', 'Job Level']],
      body: rows,
      theme: 'grid',
      styles: {
        font: 'helvetica',
        fontSize: 12.5,
        cellPadding: 3.1,
        overflow: 'linebreak',
        valign: 'middle',
        textColor: palette.ink,
        lineColor: [203, 213, 225],
        lineWidth: 0.25,
        minCellHeight: 11
      },
      headStyles: { fillColor: palette.primary, textColor: [255, 255, 255], fontStyle: 'bold', fontSize: 12.5, halign: 'center' },
      alternateRowStyles: { fillColor: [248, 250, 252] },
      columnStyles: {
        0: { cellWidth: 58, fontStyle: 'bold' },
        1: { cellWidth: pageWidth - margin * 2 - 108 },
        2: { cellWidth: 50, halign: 'center' }
      },
      margin: { left: margin, right: margin, top: 32, bottom: 15 },
      tableWidth: pageWidth - margin * 2,
      showHead: 'everyPage',
      pageBreak: 'auto',
      rowPageBreak: 'auto'
    });
  }

  function drawStaticTitleSlide(doc, key, fallbackTitle, fallbackSubtitle, palette, pageWidth, pageHeight, margin) {
    doc.addPage();
    const ok = addBackground(doc, key, pageWidth, pageHeight);
    if (!ok) {
      doc.setFillColor(255, 255, 255);
      doc.rect(0, 0, pageWidth, pageHeight, 'F');
      doc.setTextColor.apply(doc, palette.primary);
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(28);
      doc.text(fallbackTitle, margin, pageHeight / 2 - 8);
      doc.setTextColor(82, 82, 91);
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(14);
      doc.text(fallbackSubtitle, margin, pageHeight / 2 + 8);
    }
  }

  function calcObservationColumnStyles(doc, rows, pageWidth, margin) {
    const keys = ['temuan', 'kondisiIdeal', 'dampak', 'penyebab', 'tindakan', 'deadline', 'hasil'];
    const header = ['Temuan', 'Kondisi Ideal', 'Dampak', 'Penyebab', 'Tindakan Perbaikan', 'Tanggal Perbaikan/Deadline', 'Hasil'];
    const minWidths = [32, 34, 29, 29, 42, 30, 30];
    const maxWidths = [62, 62, 52, 52, 76, 44, 56];
    const available = pageWidth - margin * 2;
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(12.5);
    const ideal = keys.map(function (key, index) {
      let widestWord = doc.getTextWidth(header[index]);
      let totalLength = header[index].length;
      rows.forEach(function (row) {
        const value = text(row[key], '');
        totalLength += value.length;
        value.split(/\s+/).forEach(function (word) {
          if (word) widestWord = Math.max(widestWord, doc.getTextWidth(word));
        });
      });
      const contentBias = Math.sqrt(Math.max(12, totalLength / Math.max(1, rows.length))) * 1.7;
      const actionBias = index === 4 ? 8 : 0;
      return Math.min(maxWidths[index], Math.max(minWidths[index], widestWord + 8 + contentBias + actionBias));
    });
    let total = ideal.reduce(function (sum, value) { return sum + value; }, 0);
    let widths = ideal.slice();
    if (total > available) {
      const shrinkable = widths.map(function (value, index) { return Math.max(0, value - minWidths[index]); });
      const shrinkTotal = shrinkable.reduce(function (sum, value) { return sum + value; }, 0) || 1;
      const need = total - available;
      widths = widths.map(function (value, index) {
        return Math.max(minWidths[index], value - need * (shrinkable[index] / shrinkTotal));
      });
    } else if (total < available) {
      let spare = available - total;
      const growOrder = [4, 0, 1, 6, 2, 3, 5];
      growOrder.forEach(function (index) {
        if (spare <= 0) return;
        const room = Math.max(0, maxWidths[index] - widths[index]);
        const add = Math.min(room, spare * (index === 4 ? 0.5 : 0.22));
        widths[index] += add;
        spare -= add;
      });
      if (spare > 0) widths[4] += spare;
    }
    const after = widths.reduce(function (sum, value) { return sum + value; }, 0);
    if (Math.abs(after - available) > 0.5) {
      const scale = available / after;
      widths = widths.map(function (value) { return value * scale; });
    }
    return widths.reduce(function (styles, width, index) {
      styles[index] = { cellWidth: width };
      if (index === 5) styles[index].halign = 'center';
      return styles;
    }, {});
  }

  function parseDeadlineDate(value) {
    const raw = text(value, '');
    if (!raw || raw === '-') return null;
    const isoLike = /^\d{4}-\d{2}-\d{2}$/.test(raw) ? raw + 'T00:00:00' : raw;
    const date = new Date(isoLike);
    if (Number.isNaN(date.getTime())) return null;
    date.setHours(0, 0, 0, 0);
    return date;
  }

  function deadlineHighlight(value) {
    const due = parseDeadlineDate(value);
    if (!due) return null;
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const days = Math.ceil((due.getTime() - today.getTime()) / 86400000);
    if (days <= 3) return { fill: [254, 226, 226], text: [153, 27, 27] };
    if (days <= 7) return { fill: [220, 252, 231], text: [22, 101, 52] };
    return { fill: [219, 234, 254], text: [30, 64, 175] };
  }

  function drawObservationStackPageHeader(doc, title, rowIndex, totalRows, row, palette, pageWidth, pageHeight, margin, continuation) {
    doc.addPage();
    doc.setFillColor(255, 255, 255);
    doc.rect(0, 0, pageWidth, pageHeight, 'F');
    drawTopBar(doc, title, palette, pageWidth);

    const cardX = margin;
    const cardW = pageWidth - margin * 2;
    let y = 29;
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(9.2);
    doc.setTextColor(100, 116, 139);
    doc.text('Temuan ' + String(rowIndex + 1) + ' dari ' + String(totalRows) + (continuation ? ' - lanjutan' : ''), cardX, 25.2);

    const titleLines = doc.splitTextToSize(text(row.temuan, 'Temuan ' + String(rowIndex + 1)), cardW - 18).slice(0, 3);
    const titleH = Math.max(14, 8 + titleLines.length * 5.1);
    doc.setFillColor.apply(doc, palette.primary);
    doc.roundedRect(cardX, y, cardW, titleH, 4, 4, 'F');
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(12);
    doc.setTextColor(255, 255, 255);
    doc.text(String(rowIndex + 1) + '. ' + titleLines[0], cardX + 7, y + 8.5);
    if (titleLines.length > 1) {
      doc.setFontSize(10.5);
      doc.text(titleLines.slice(1), cardX + 7, y + 13.5);
    }
    return y + titleH + 5;
  }

  function drawObservationStackRow(doc, label, lines, x, y, width, labelWidth, rowHeight, palette, options) {
    const opts = options || {};
    const valueFill = opts.valueFill || [255, 255, 255];
    const valueText = opts.valueText || palette.ink;
    doc.setDrawColor(203, 213, 225);
    doc.setLineWidth(0.22);
    doc.setFillColor(236, 253, 245);
    doc.roundedRect(x, y, labelWidth, rowHeight, 2.4, 2.4, 'FD');
    doc.setFillColor.apply(doc, valueFill);
    doc.roundedRect(x + labelWidth, y, width - labelWidth, rowHeight, 2.4, 2.4, 'FD');

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(8.2);
    doc.setTextColor.apply(doc, palette.primary);
    doc.text(doc.splitTextToSize(label, labelWidth - 6), x + 3.2, y + 6.3);

    doc.setFont('helvetica', opts.boldValue ? 'bold' : 'normal');
    doc.setFontSize(8.8);
    doc.setTextColor.apply(doc, valueText);
    doc.text(lines.length ? lines : ['-'], x + labelWidth + 4, y + 6.1);
  }

  function fitLinesToBox(doc, value, width, maxLines) {
    const lines = doc.splitTextToSize(text(value, '-'), width);
    if (lines.length <= maxLines) return lines.length ? lines : ['-'];
    const clipped = lines.slice(0, Math.max(1, maxLines));
    const lastIndex = clipped.length - 1;
    let last = clipped[lastIndex] || '';
    while (last.length > 4 && doc.getTextWidth(last + '...') > width) last = last.slice(0, -1);
    clipped[lastIndex] = (last || clipped[lastIndex]).replace(/[\s.,;:-]+$/g, '') + '...';
    return clipped;
  }

  function makeObservationCompactRows(doc, row, tableWidth) {
    const cellGap = 1.0;
    const cellPadX = 2.2;
    const fullW = tableWidth;
    const halfW = (tableWidth - cellGap) / 2;
    const valueFontSize = pdfTableFontSize();

    function buildField(label, value, width, options) {
      const opts = options || {};
      doc.setFontSize(valueFontSize);
      doc.setFont('helvetica', opts.boldValue ? 'bold' : 'normal');
      const richLines = splitRichTextToLines(doc, value, width - cellPadX * 2, opts.boldValue);
      return Object.assign({}, opts, {
        label: label,
        value: value,
        rawValue: value,
        richLines: richLines.length ? richLines : [[{ text: '-' }]],
        lines: (richLines.length ? richLines : [[{ text: '-' }]]).map(richLineToText)
      });
    }

    function chunkField(field, maxLines) {
      const lines = field.richLines && field.richLines.length ? field.richLines : [[{ text: '-' }]];
      const chunks = [];
      for (let i = 0; i < lines.length; i += maxLines) {
        chunks.push(Object.assign({}, field, {
          label: field.label + (i > 0 ? ' (lanjutan)' : ''),
          richLines: lines.slice(i, i + maxLines),
          lines: lines.slice(i, i + maxLines).map(richLineToText)
        }));
      }
      return chunks.length ? chunks : [Object.assign({}, field, { richLines: [[{ text: '-' }]], lines: ['-'] })];
    }

    function addFull(rows, field, maxLines) {
      chunkField(field, maxLines).forEach(function (chunk) {
        rows.push({ cells: [chunk], full: true });
      });
    }

    function addPair(rows, leftField, rightField, maxLines) {
      const leftChunks = chunkField(leftField, maxLines);
      const rightChunks = chunkField(rightField, maxLines);
      const total = Math.max(leftChunks.length, rightChunks.length);
      for (let i = 0; i < total; i += 1) {
        rows.push({ cells: [leftChunks[i] || null, rightChunks[i] || null], full: false });
      }
    }

    const rows = [];
    addFull(rows, buildField('Temuan', row.temuan || '-', fullW, { highlight: { fill: [254, 249, 195], text: [66, 32, 6] }, boldValue: true, labelColor: [133, 77, 14] }), 5);
    addPair(
      rows,
      buildField('Kondisi Ideal', row.kondisiIdeal || '-', halfW),
      buildField('Dampak', row.dampak || '-', halfW),
      5
    );
    addFull(rows, buildField('Penyebab', row.penyebab || '-', fullW), 5);
    addFull(rows, buildField('Tindakan Perbaikan', row.tindakan || '-', fullW), 6);
    addFull(rows, buildField('Deadline', row.deadline ? formatDate(row.deadline) : '-', fullW, { highlight: deadlineHighlight(row.deadline), boldValue: true }), 2);
    addFull(rows, buildField('Hasil', row.hasil || '-', fullW), 5);
    return rows;
  }

  function observationCompactCellHeight(field) {
    if (!field) return 0;
    const fontSize = pdfTableFontSize();
    const labelFontSize = pdfTableTitleFontSize();
    const lineHeight = Math.max(3.35, fontSize * 0.405);
    const lineCount = Math.max(1, (field.lines || ['-']).length);
    const labelBlock = Math.max(7.6, pdfFieldTitleGap() + labelFontSize * 0.18);
    return Math.max(fontSize + labelFontSize * 0.62, 2.4 + labelBlock + lineCount * lineHeight);
  }

  function observationCompactRowHeight(row) {
    return Math.max.apply(null, row.cells.map(observationCompactCellHeight));
  }

  function drawObservationCompactCell(doc, field, x, y, width, height, palette, fillColor) {
    if (!field) {
      doc.setDrawColor(226, 232, 240);
      doc.setFillColor(255, 255, 255);
      doc.roundedRect(x, y, width, height, 1.2, 1.2, 'FD');
      return;
    }
    const valueFill = field.highlight ? field.highlight.fill : fillColor;
    const valueText = field.highlight ? field.highlight.text : [15, 23, 42];
    doc.setDrawColor(148, 163, 184);
    doc.setLineWidth(0.28);
    doc.setFillColor.apply(doc, valueFill);
    doc.roundedRect(x, y, width, height, 1.2, 1.2, 'FD');

    const valueFontSize = pdfTableFontSize();
    const labelFontSize = pdfTableTitleFontSize();
    const lineHeight = Math.max(3.35, valueFontSize * 0.405);
    const labelY = y + Math.max(4.2, labelFontSize * 0.44);
    const valueY = labelY + pdfFieldTitleGap();
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(labelFontSize);
    doc.setTextColor.apply(doc, field.labelColor || [30, 64, 175]);
    doc.text(field.label, x + 2.2, labelY, { baseline: 'alphabetic' });

    doc.setFont('helvetica', field.boldValue ? 'bold' : 'normal');
    doc.setFontSize(valueFontSize);
    doc.setTextColor.apply(doc, valueText);
    drawRichLines(doc, field.richLines || (field.lines || ['-']).map(function (line) { return [{ text: line }]; }), x + 2.2, valueY, lineHeight, { bold: field.boldValue, textColor: valueText, fontSize: valueFontSize });
  }

  function addObservationListPage(doc, title, palette, pageWidth, pageHeight, margin) {
    doc.addPage();
    doc.setFillColor(255, 255, 255);
    doc.rect(0, 0, pageWidth, pageHeight, 'F');
    drawTopBar(doc, title, palette, pageWidth);
    return 22;
  }

  function observationCompactSegmentHeight(rows) {
    const density = pdfTableExtraRows();
    const titleFontSize = clampNumber(pdfTableTitleFontSize() - 1.3, 7.4, 10.8, 8.6);
    const titleH = Math.max(5.8, titleFontSize * 0.62);
    const titleGap = Math.max(1.6, 2.4 - density * 0.08);
    const pad = Math.max(1.45, 2.2 - density * 0.16);
    const rowGap = Math.max(0.35, 0.75 - density * 0.08);
    let bodyH = pad * 2;
    rows.forEach(function (row, index) {
      bodyH += observationCompactRowHeight(row) + (index ? rowGap : 0);
    });
    return titleH + titleGap + bodyH;
  }

  function drawObservationCompactSegment(doc, rowIndex, totalRows, rows, x, y, width, palette, continuation) {
    const density = pdfTableExtraRows();
    const titleFontSize = clampNumber(pdfTableTitleFontSize() - 1.3, 7.4, 10.8, 8.6);
    const titleH = Math.max(5.8, titleFontSize * 0.62);
    const titleGap = Math.max(1.6, 2.4 - density * 0.08);
    const pad = Math.max(1.45, 2.2 - density * 0.16);
    const rowGap = Math.max(0.35, 0.75 - density * 0.08);
    const cellGap = 1.0;
    const tableY = y + titleH + titleGap;
    const tableW = width;
    const innerW = tableW - pad * 2;
    const halfW = (innerW - cellGap) / 2;
    let bodyH = pad * 2;
    rows.forEach(function (row, index) {
      bodyH += observationCompactRowHeight(row) + (index ? rowGap : 0);
    });

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(titleFontSize);
    doc.setTextColor.apply(doc, palette.primary);
    doc.text('Temuan ' + String(rowIndex + 1) + ' dari ' + String(totalRows) + (continuation ? ' - lanjutan' : ''), x, y + Math.max(4.8, titleFontSize * 0.42));

    doc.setDrawColor(148, 163, 184);
    doc.setLineWidth(0.34);
    doc.setFillColor(255, 255, 255);
    doc.roundedRect(x, tableY, tableW, bodyH, 2.2, 2.2, 'FD');
    doc.setFillColor.apply(doc, palette.primary);
    doc.roundedRect(x, tableY, tableW, 0.8, 2.2, 2.2, 'F');

    let cy = tableY + pad;
    rows.forEach(function (row, index) {
      if (index) cy += rowGap;
      const rowH = observationCompactRowHeight(row);
      const fill = index % 2 ? [249, 250, 251] : [255, 255, 255];
      if (row.full) {
        drawObservationCompactCell(doc, row.cells[0], x + pad, cy, innerW, rowH, palette, fill);
      } else {
        drawObservationCompactCell(doc, row.cells[0], x + pad, cy, halfW, rowH, palette, fill);
        drawObservationCompactCell(doc, row.cells[1], x + pad + halfW + cellGap, cy, halfW, rowH, palette, fill);
      }
      cy += rowH;
    });
    return observationCompactSegmentHeight(rows);
  }

  function drawObservationTable(doc, title, rows, palette, pageWidth, pageHeight, margin) {
    const cleanRows = normalizeRows(rows);
    if (!cleanRows.length) return;
    const tableMargin = 3.2;
    const x = tableMargin;
    const width = pageWidth - tableMargin * 2;
    const density = pdfTableExtraRows();
    const bottom = pageHeight - Math.max(5, 12 - density * 1.6);
    const gap = Math.max(1.2, 2.4 - density * 0.22);
    const innerTableWidth = width - 3.6;
    let y = addObservationListPage(doc, title, palette, pageWidth, pageHeight, margin);

    cleanRows.forEach(function (row, rowIndex) {
      const allRows = makeObservationCompactRows(doc, row, innerTableWidth);
      let remaining = allRows.slice();
      let continuation = false;
      while (remaining.length) {
        const available = Math.max(36, bottom - y);
        const segment = [];
        let used = observationCompactSegmentHeight(segment);
        for (let i = 0; i < remaining.length; i += 1) {
          const candidate = segment.concat([remaining[i]]);
          const candidateHeight = observationCompactSegmentHeight(candidate);
          if (segment.length && candidateHeight > available) break;
          segment.push(remaining[i]);
          used = candidateHeight;
          if (used > available) break;
        }
        if (!segment.length) {
          y = addObservationListPage(doc, title, palette, pageWidth, pageHeight, margin);
          continue;
        }
        if (y + used > bottom && y > 26) {
          y = addObservationListPage(doc, title, palette, pageWidth, pageHeight, margin);
          continue;
        }
        const drawnHeight = drawObservationCompactSegment(doc, rowIndex, cleanRows.length, segment, x, y, width, palette, continuation);
        y += drawnHeight + gap;
        remaining = remaining.slice(segment.length);
        continuation = true;
        if (remaining.length && y > bottom - 18) {
          y = addObservationListPage(doc, title, palette, pageWidth, pageHeight, margin);
        }
      }
      if (y > bottom - 22 && rowIndex < cleanRows.length - 1) {
        y = addObservationListPage(doc, title, palette, pageWidth, pageHeight, margin);
      }
    });
  }

  function buildPhotoGridItems(doc, photos, cardWidth, cardHeight) {
    const evidenceFontSize = pdfEvidenceFontSize();
    const lineHeight = Math.max(3.6, evidenceFontSize * 0.45);
    const imageHeight = Math.max(34, Math.min(cardHeight * 0.72, cardHeight - 18));
    const maxLinesWithImage = Math.max(1, Math.floor((cardHeight - imageHeight - 10) / lineHeight));
    const maxLinesTextOnly = Math.max(2, Math.floor((cardHeight - 12) / lineHeight));
    const items = [];
    photos.forEach(function (photo, index) {
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(evidenceFontSize);
      const richLines = splitRichTextToLines(doc, photo.description, Math.max(12, cardWidth - 12), false);
      const first = richLines.slice(0, maxLinesWithImage);
      items.push({ image: photo.image, rawDescription: photo.description, richLines: first, lines: first.map(richLineToText), lineHeight: lineHeight, imageHeight: imageHeight, title: 'Foto ' + String(index + 1), continuation: false });
      let rest = richLines.slice(maxLinesWithImage);
      let continuationIndex = 1;
      while (rest.length) {
        const chunk = rest.slice(0, maxLinesTextOnly);
        items.push({ image: '', rawDescription: chunk.map(richLineToText).join('\n'), richLines: chunk, lines: chunk.map(richLineToText), lineHeight: lineHeight, imageHeight: 0, imageSize: 0, title: 'Lanjutan deskripsi Foto ' + String(index + 1) + '.' + String(continuationIndex), continuation: true });
        rest = rest.slice(maxLinesTextOnly);
        continuationIndex += 1;
      }
    });
    return items;
  }

  async function drawPhotoGridCard(doc, item, x, y, width, height, palette) {
    doc.setFillColor(255, 255, 255);
    doc.setDrawColor(148, 163, 184);
    doc.setLineWidth(0.24);
    doc.roundedRect(x, y, width, height, 3, 3, 'FD');

    let descY = y + 5;
    let descHeight = height - 8;
    if (item.image) {
      const imgX = x + 3;
      const imgY = y + 3;
      const imgW = Math.max(18, width - 5);
      const imgH = Math.max(30, Math.min(item.imageHeight || 46, height - 15));
      doc.setFillColor(248, 250, 252);
      doc.roundedRect(imgX, imgY, imgW, imgH, 2.5, 2.5, 'F');
      const added = await addImageInBox(doc, item.image, imgX + 0.6, imgY + 0.6, imgW - 1.2, imgH - 1.2, 'contain');
      if (!added) {
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(7.8);
        doc.setTextColor(100, 116, 139);
        doc.text('No photo', x + width / 2, imgY + imgH / 2, { align: 'center' });
      }
      descY = imgY + imgH + 3.2;
      descHeight = height - (descY - y) - 3;
    }

    doc.setFillColor(248, 250, 252);
    doc.roundedRect(x + 3, descY - 1.4, width - 6, Math.max(8, descHeight), 2.5, 2.5, 'F');
    const evidenceFontSize = pdfEvidenceFontSize();
    const evidenceLineHeight = item.lineHeight || Math.max(3.6, evidenceFontSize * 0.45);
    if (item.continuation) {
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(clampNumber(evidenceFontSize - 1.1, 7.2, 9.2, 7.6));
      const continuationTitleColor = palette && palette.primary ? palette.primary : [30, 64, 175];
      doc.setTextColor.apply(doc, continuationTitleColor);
      doc.text(item.title, x + 5, descY + 1.8, { baseline: 'top' });
      descY += Math.max(4.2, evidenceLineHeight);
    }
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(evidenceFontSize);
    doc.setTextColor(15, 23, 42);
    drawRichLines(doc, item.richLines || (item.lines || ['-']).map(function (line) { return [{ text: line }]; }), x + 5, descY + 2.1, evidenceLineHeight, { textColor: [15, 23, 42], fontSize: evidenceFontSize, textOptions: { baseline: 'top' } });
  }

  async function drawPhotoSlides(doc, title, subtitle, templateKey, photos, palette, pageWidth, pageHeight, margin) {
    const cleanPhotos = normalizePhotos(photos);
    if (!cleanPhotos.length) return;

    drawStaticTitleSlide(doc, templateKey, title, subtitle, palette, pageWidth, pageHeight, margin);

    const preferredGrid = pdfPhotoGridPerPage();
    const columns = preferredGrid === 4 ? 2 : preferredGrid === 8 ? 4 : 3;
    const rowsPerPage = 2;
    const maxPerPage = columns * rowsPerPage;
    const gap = preferredGrid === 8 ? 3.0 : 3.6;
    const contentTop = 22;
    const contentBottom = pageHeight - 9;
    const cardWidth = (pageWidth - margin * 2 - gap * (columns - 1)) / columns;
    const cardHeight = (contentBottom - contentTop - gap * (rowsPerPage - 1)) / rowsPerPage;
    const items = buildPhotoGridItems(doc, cleanPhotos, cardWidth, cardHeight);

    function newPhotoPage() {
      doc.addPage();
      doc.setFillColor(255, 255, 255);
      doc.rect(0, 0, pageWidth, pageHeight, 'F');
      drawTopBar(doc, title + ' - Evidence Photos', palette, pageWidth);
    }

    for (let index = 0; index < items.length; index += 1) {
      if (index % maxPerPage === 0) newPhotoPage();
      const slot = index % maxPerPage;
      const col = slot % columns;
      const row = Math.floor(slot / columns);
      const x = margin + col * (cardWidth + gap);
      const y = contentTop + row * (cardHeight + gap);
      await drawPhotoGridCard(doc, items[index], x, y, cardWidth, cardHeight, palette);
    }
  }

  function drawAssignmentSlide(doc, data, palette, pageWidth, pageHeight, margin) {
    drawStaticTitleSlide(doc, 'assignmentTitle', 'Store Assignment', 'Corrective Action Purpose', palette, pageWidth, pageHeight, margin);

    doc.addPage();
    doc.setFillColor(255, 255, 255);
    doc.rect(0, 0, pageWidth, pageHeight, 'F');
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(18);
    doc.setTextColor.apply(doc, palette.ink);
    doc.text('Mekanisme pelaporan', margin, 30);

    let y = 47;
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(14);
    doc.setTextColor.apply(doc, palette.primary);
    doc.text(getAssignmentLink(data), margin, y);
    y += 17;

    doc.setTextColor.apply(doc, palette.ink);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(12.5);
    const steps = [
      '1. Unduh file pada link di atas.',
      '2. Tim store mengisi form tersebut berdasarkan temuan yang telah ditunjukkan pada file laporan ini.',
      '3. Tindakan perbaikan WAJIB dilakukan sebelum deadline yang diberikan oleh Regional Bestie.',
      '4. Form tindakan perbaikan yang telah dibuat WAJIB dikirimkan kembali via email dengan terusan: Regional Manager, Area Manager, Regional Bestie, dan FMCU.'
    ];
    steps.forEach(function (step) {
      y = addWrapped(doc, step, margin, y, pageWidth - margin * 2, 7.2);
      y += 4;
    });
  }

  async function createDocument(data) {
    if (!window.jspdf || !window.jspdf.jsPDF) throw new Error('Library jsPDF belum tersedia. Pastikan koneksi internet aktif saat membuka halaman.');
    const doc = new window.jspdf.jsPDF({ orientation: 'landscape', unit: 'mm', format: [320, 180] });
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const margin = 10;
    const palette = { primary: [0, 153, 166], accent: [249, 115, 22], ink: [39, 39, 42] };
    const detail = getStoreDetail(data && data.store, data || {});

    drawCover(doc, data || {}, detail, palette, pageWidth, pageHeight, margin);
    await drawQscResultSlide(doc, data || {}, palette, pageWidth, pageHeight, margin);
    drawCrewSlide(doc, data || {}, palette, pageWidth, pageHeight, margin);

    if (data && data.showOPITable === true && normalizeRows(data.opiData).length) {
      drawStaticTitleSlide(doc, 'opiTitle', 'OPI Project Observation', 'Findings & Root Cause Analysis', palette, pageWidth, pageHeight, margin);
      drawObservationTable(doc, 'OPI Project Observation', data.opiData, palette, pageWidth, pageHeight, margin);
    }
    if (data && data.showQSCTable === true && normalizeRows(data.qscData).length) {
      drawStaticTitleSlide(doc, 'qscTitle', 'QSC Observation', 'Findings & Root Cause Analysis', palette, pageWidth, pageHeight, margin);
      drawObservationTable(doc, 'QSC Observation', data.qscData, palette, pageWidth, pageHeight, margin);
    }
    if (data && data.showFindingEvidence === true) await drawPhotoSlides(doc, 'Finding Evidence', 'of OPI & QSC Observation', 'findingTitle', data.findingEvidencePhotos, palette, pageWidth, pageHeight, margin);
    if (data && data.showCorrectiveAction === true) await drawPhotoSlides(doc, 'Corrective Action Evidence & Result', 'by Regional Bestie', 'correctiveTitle', data.correctiveActionPhotos, palette, pageWidth, pageHeight, margin);
    drawAssignmentSlide(doc, data || {}, palette, pageWidth, pageHeight, margin);
    applyFooterAllPages(doc, pageWidth, pageHeight);
    return doc;
  }

  async function createBlob(data) {
    const doc = await createDocument(data);
    return doc.output('blob');
  }

  async function save(data) {
    const doc = await createDocument(data);
    doc.save(buildFileName(data));
  }

  window.ReportVisitPDF = { createDocument: createDocument, createBlob: createBlob, save: save, buildFileName: buildFileName };
})();
