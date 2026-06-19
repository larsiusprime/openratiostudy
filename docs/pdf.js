const PAGE = {
  margin: 15,
  usableWidth: 180,
  textMaxWidth: 170,
  imageHeight: 90,
  bodyFontSize: 11,
  titleFontSize: 14,
};

function ensureSpace(doc, cursorY, neededHeight) {
  const pageHeight = doc.internal.pageSize.getHeight();
  if (cursorY + neededHeight > pageHeight - PAGE.margin) {
    doc.addPage();
    return PAGE.margin;
  }
  return cursorY;
}

function resolveJsPDF() {
  if (window.jspdf && window.jspdf.jsPDF) return window.jspdf.jsPDF;
  if (window.jsPDF) return window.jsPDF;
  throw new Error('jsPDF not found on window. Check jspdf.umd.min.js include path.');
}

function parseCssColorToRgb(input) {
  const s = String(input || '').trim().toLowerCase();
  if (!s) return [51, 65, 85];
  if (s.startsWith('#')) {
    const hex = s.slice(1);
    if (hex.length === 3) {
      return hex.split('').map((c) => parseInt(c + c, 16));
    }
    if (hex.length === 6) {
      return [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
    }
  }
  const m = s.match(/^rgb\(\s*([0-9]{1,3})\s*,\s*([0-9]{1,3})\s*,\s*([0-9]{1,3})\s*\)$/);
  if (m) return [Number(m[1]), Number(m[2]), Number(m[3])];
  return [51, 65, 85];
}

async function drawChartSection(doc, section, cursorY) {
  const title = typeof section.title === 'string' ? section.title : '';
  if (title) {
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(PAGE.titleFontSize);
    cursorY = ensureSpace(doc, cursorY, 10);
    doc.text(title, PAGE.margin, cursorY, { maxWidth: PAGE.usableWidth });
    cursorY += 8;
  }

  const chartNode = document.getElementById(section.plotlyDivId);
  if (!chartNode || !window.Plotly || typeof Plotly.toImage !== 'function') {
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(PAGE.bodyFontSize);
    cursorY = ensureSpace(doc, cursorY, 7);
    doc.text(`Chart unavailable: ${title}`, PAGE.margin, cursorY, { maxWidth: PAGE.usableWidth });
    return cursorY + 8;
  }

  try {
    const dataUrl = await Plotly.toImage(chartNode, {
      format: 'png',
      width: 900,
      height: 500,
    });
    const targetHeight = PAGE.imageHeight;
    const targetWidth = Math.min(PAGE.usableWidth, (900 / 500) * targetHeight);
    cursorY = ensureSpace(doc, cursorY, targetHeight);
    doc.addImage(dataUrl, 'PNG', PAGE.margin, cursorY, targetWidth, targetHeight);
    return cursorY + PAGE.imageHeight + 8;
  } catch (err) {
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(PAGE.bodyFontSize);
    cursorY = ensureSpace(doc, cursorY, 7);
    doc.text(`Chart unavailable: ${title}`, PAGE.margin, cursorY, { maxWidth: PAGE.usableWidth });
    return cursorY + 8;
  }
}

async function generatePDF(config) {
  const jsPDFCtor = resolveJsPDF();
  const doc = new jsPDFCtor({ unit: 'mm', format: 'a4' });
  const sections = Array.isArray(config && config.sections) ? config.sections : [];
  const filename = (config && config.filename) ? config.filename : 'report.pdf';
  const bw = !!(config && config.bw);
  let cursorY = PAGE.margin;

  for (const section of sections) {
    if (!section || !section.type) continue;
    if (section.pageBreakBefore) {
      doc.addPage();
      cursorY = PAGE.margin;
    }

    if (section.type === 'text') {
      const isTitle = Boolean(section.isTitle);
      const small = Boolean(section.small);
      const align = section.align === 'right' ? 'right' : (section.align === 'center' ? 'center' : 'left');
      doc.setFont('helvetica', isTitle ? 'bold' : 'normal');
      doc.setFontSize(isTitle ? PAGE.titleFontSize : (small ? PAGE.bodyFontSize - 2 : PAGE.bodyFontSize));
      if (small) doc.setTextColor(100, 116, 139); // muted gray for fine print
      const text = String(section.content || '');
      const wrapped = doc.splitTextToSize(text, PAGE.textMaxWidth);
      const lineHeight = isTitle ? 6 : (small ? 4.4 : 5.2);
      const estimatedHeight = Math.max(small ? 5 : 8, wrapped.length * lineHeight);
      cursorY = ensureSpace(doc, cursorY, estimatedHeight);
      const x = align === 'right' ? (PAGE.margin + PAGE.usableWidth)
        : align === 'center' ? (PAGE.margin + PAGE.usableWidth / 2)
        : PAGE.margin;
      doc.text(wrapped, x, cursorY, { maxWidth: PAGE.textMaxWidth, align });
      if (small) doc.setTextColor(31, 41, 55); // reset
      cursorY += estimatedHeight + 2;
      continue;
    }

    if (section.type === 'table') {
      const headers = Array.isArray(section.headers) ? section.headers : [];
      const rows = Array.isArray(section.rows) ? section.rows : [];
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(PAGE.bodyFontSize);
      const tableOptions = {
        body: rows,
        startY: cursorY,
        margin: { left: PAGE.margin, right: PAGE.margin, top: PAGE.margin, bottom: PAGE.margin },
        styles: { font: 'helvetica', fontSize: PAGE.bodyFontSize },
        headStyles: { fontStyle: 'bold' },
        didDrawPage: (data) => {
          cursorY = data.cursor && data.cursor.y ? data.cursor.y : PAGE.margin;
        },
      };
      if (bw) {
        // Black & white: bold black headers on white with a black outline around every cell.
        tableOptions.theme = 'grid';
        tableOptions.styles = { font: 'helvetica', fontSize: PAGE.bodyFontSize, textColor: [0, 0, 0], lineColor: [0, 0, 0], lineWidth: 0.2 };
        tableOptions.headStyles = { fontStyle: 'bold', fillColor: [255, 255, 255], textColor: [0, 0, 0], lineColor: [0, 0, 0], lineWidth: 0.2 };
      }
      if (headers.length) {
        tableOptions.head = [headers];
      }
      doc.autoTable(tableOptions);
      cursorY = (doc.lastAutoTable && doc.lastAutoTable.finalY ? doc.lastAutoTable.finalY : cursorY) + 8;
      continue;
    }

    if (section.type === 'fieldchip') {
      // A label followed by a rounded "capsule" containing a literal field name,
      // mirroring the .field-chip styling used on the page.
      const label = String(section.label || 'Chosen field:');
      const value = String(section.value || '');
      const labelSize = PAGE.bodyFontSize;
      const valSize = PAGE.bodyFontSize - 1;
      const padX = 2.4;
      const capH = 5.8;

      doc.setFont('helvetica', 'normal');
      doc.setFontSize(labelSize);
      const labelW = doc.getTextWidth(label + ' ');

      doc.setFont('courier', 'normal');
      doc.setFontSize(valSize);
      const capW = doc.getTextWidth(value) + padX * 2;

      cursorY = ensureSpace(doc, cursorY, capH + 6);
      const baseline = cursorY + 3;

      // Label
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(labelSize);
      doc.setTextColor(31, 41, 55);
      doc.text(label, PAGE.margin, baseline);

      // Capsule (white/black in B&W mode, indigo otherwise)
      const capX = PAGE.margin + labelW;
      const capY = baseline - 4.1;
      doc.setFillColor(255, 255, 255);
      doc.setDrawColor(bw ? 0 : 199, bw ? 0 : 210, bw ? 0 : 254);
      doc.setLineWidth(0.2);
      doc.roundedRect(capX, capY, capW, capH, 1.4, 1.4, 'FD');
      doc.setFont('courier', 'normal');
      doc.setFontSize(valSize);
      doc.setTextColor(bw ? 0 : 55, bw ? 0 : 55, bw ? 0 : 163);
      doc.text(value, capX + padX, baseline);

      // Reset
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(PAGE.bodyFontSize);
      doc.setTextColor(31, 41, 55);
      cursorY = capY + capH + 6;
      continue;
    }

    if (section.type === 'titlechip') {
      // A bold heading that leads with a capsule (the literal field name),
      // followed by the rest of the heading text.
      const value = String(section.value || '');
      const suffix = String(section.suffix || '').replace(/^\s+/, '');
      const size = PAGE.titleFontSize;
      const valSize = size - 2;
      const padX = 2.4;
      const capH = 6.8;

      doc.setFont('courier', 'bold');
      doc.setFontSize(valSize);
      const capW = doc.getTextWidth(value) + padX * 2;

      cursorY = ensureSpace(doc, cursorY, capH + 6);
      const baseline = cursorY + 5;
      const capX = PAGE.margin;
      const capY = baseline - 5;

      // Capsule (white/black in B&W mode, indigo otherwise)
      doc.setFillColor(255, 255, 255);
      doc.setDrawColor(bw ? 0 : 199, bw ? 0 : 210, bw ? 0 : 254);
      doc.setLineWidth(0.2);
      doc.roundedRect(capX, capY, capW, capH, 1.6, 1.6, 'FD');
      doc.setFont('courier', 'bold');
      doc.setFontSize(valSize);
      doc.setTextColor(bw ? 0 : 55, bw ? 0 : 55, bw ? 0 : 163);
      doc.text(value, capX + padX, baseline);

      // Suffix (rest of the heading)
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(size);
      doc.setTextColor(bw ? 0 : 31, bw ? 0 : 41, bw ? 0 : 55);
      doc.text(suffix, capX + capW + 1.8, baseline);

      // Reset
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(PAGE.bodyFontSize);
      doc.setTextColor(31, 41, 55);
      cursorY = capY + capH + 6;
      continue;
    }

    if (section.type === 'chart') {
      cursorY = await drawChartSection(doc, section, cursorY);
      continue;
    }

    if (section.type === 'image') {
      const title = typeof section.title === 'string' ? section.title : '';
      const dataUrl = typeof section.dataUrl === 'string' ? section.dataUrl : '';
      if (title) {
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(PAGE.titleFontSize);
        cursorY = ensureSpace(doc, cursorY, 10);
        doc.text(title, PAGE.margin, cursorY, { maxWidth: PAGE.usableWidth });
        cursorY += 8;
      }
      if (!dataUrl) {
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(PAGE.bodyFontSize);
        cursorY = ensureSpace(doc, cursorY, 7);
        doc.text('Image unavailable.', PAGE.margin, cursorY, { maxWidth: PAGE.usableWidth });
        cursorY += 8;
        continue;
      }
      const imgW = Math.max(1, Number(section.imageWidth || 1200));
      const imgH = Math.max(1, Number(section.imageHeight || 700));
      const imageFormat = String(section.imageFormat || (dataUrl.startsWith('data:image/jpeg') ? 'JPEG' : 'PNG')).toUpperCase();
      const targetWidth = PAGE.usableWidth;
      const targetHeight = targetWidth * (imgH / imgW);
      cursorY = ensureSpace(doc, cursorY, targetHeight);
      try {
        console.info('[pdf-map] addImage diagnostics', {
          dataUrlLength: dataUrl.length,
          imageFormat,
          sourceWidth: imgW,
          sourceHeight: imgH,
          targetWidth,
          targetHeight
        });
        doc.addImage(dataUrl, imageFormat, PAGE.margin, cursorY, targetWidth, targetHeight);
        cursorY += targetHeight + 8;
      } catch (err) {
        console.warn('[pdf-map] addImage failed', err);
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(PAGE.bodyFontSize);
        cursorY = ensureSpace(doc, cursorY, 7);
        doc.text('Map image could not be embedded in PDF.', PAGE.margin, cursorY, { maxWidth: PAGE.usableWidth });
        cursorY += 8;
      }
    }

    if (section.type === 'legend') {
      const rows = Array.isArray(section.rows) ? section.rows : [];
      const title = typeof section.title === 'string' ? section.title : 'Legend';
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(PAGE.titleFontSize);
      cursorY = ensureSpace(doc, cursorY, 10);
      doc.text(title, PAGE.margin, cursorY, { maxWidth: PAGE.usableWidth });
      cursorY += 7;
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(PAGE.bodyFontSize);
      for (const row of rows) {
        const color = Array.isArray(row) ? row[0] : '';
        const label = Array.isArray(row) ? String(row[1] || '') : '';
        cursorY = ensureSpace(doc, cursorY, 7);
        const [r, g, b] = parseCssColorToRgb(color);
        const cx = PAGE.margin + 2.8;
        const cy = cursorY - 1.6;
        doc.setDrawColor(51, 65, 85);
        doc.setFillColor(r, g, b);
        doc.circle(cx, cy, 1.7, 'FD');
        doc.setTextColor(31, 41, 55);
        doc.text(label, PAGE.margin + 8, cursorY);
        cursorY += 5.2;
      }
      cursorY += 2;
    }
  }

  doc.save(filename);
}

window.generatePDF = generatePDF;
