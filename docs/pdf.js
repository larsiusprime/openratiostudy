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

async function drawChartSection(doc, section, cursorY) {
  const title = section.title || 'Chart';
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(PAGE.titleFontSize);
  cursorY = ensureSpace(doc, cursorY, 10);
  doc.text(title, PAGE.margin, cursorY, { maxWidth: PAGE.usableWidth });
  cursorY += 8;

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

export async function generatePDF(config) {
  const jsPDFCtor = resolveJsPDF();
  const doc = new jsPDFCtor({ unit: 'mm', format: 'a4' });
  const sections = Array.isArray(config && config.sections) ? config.sections : [];
  const filename = (config && config.filename) ? config.filename : 'report.pdf';
  let cursorY = PAGE.margin;

  for (const section of sections) {
    if (!section || !section.type) continue;

    if (section.type === 'text') {
      const isTitle = Boolean(section.isTitle);
      doc.setFont('helvetica', isTitle ? 'bold' : 'normal');
      doc.setFontSize(isTitle ? PAGE.titleFontSize : PAGE.bodyFontSize);
      const text = String(section.content || '');
      const wrapped = doc.splitTextToSize(text, PAGE.textMaxWidth);
      const lineHeight = isTitle ? 6 : 5.2;
      const estimatedHeight = Math.max(8, wrapped.length * lineHeight);
      cursorY = ensureSpace(doc, cursorY, estimatedHeight);
      doc.text(wrapped, PAGE.margin, cursorY, { maxWidth: PAGE.textMaxWidth });
      cursorY += estimatedHeight + 2;
      continue;
    }

    if (section.type === 'table') {
      const headers = Array.isArray(section.headers) ? section.headers : [];
      const rows = Array.isArray(section.rows) ? section.rows : [];
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(PAGE.bodyFontSize);
      doc.autoTable({
        head: [headers],
        body: rows,
        startY: cursorY,
        margin: { left: PAGE.margin, right: PAGE.margin, top: PAGE.margin, bottom: PAGE.margin },
        styles: { font: 'helvetica', fontSize: PAGE.bodyFontSize },
        headStyles: { fontStyle: 'bold' },
        didDrawPage: (data) => {
          cursorY = data.cursor && data.cursor.y ? data.cursor.y : PAGE.margin;
        },
      });
      cursorY = (doc.lastAutoTable && doc.lastAutoTable.finalY ? doc.lastAutoTable.finalY : cursorY) + 8;
      continue;
    }

    if (section.type === 'chart') {
      cursorY = await drawChartSection(doc, section, cursorY);
    }
  }

  doc.save(filename);
}

window.generatePDF = generatePDF;
