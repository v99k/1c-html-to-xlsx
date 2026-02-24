const ExcelJS = require('exceljs');
const { parse } = require('node-html-parser');

const ENTITIES = { nbsp: '\u00A0', amp: '&', lt: '<', gt: '>', quot: '"' };
function decodeText(html) {
  if (!html) return '';
  return String(html)
    .replace(/&nbsp;/gi, '\u00A0')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#x([0-9a-f]+);/gi, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(parseInt(d, 10)));
}

function getText(node) {
  if (!node) return '';
  let out = '';
  for (const child of node.childNodes || []) {
    if (child.nodeType === 1) {
      const tag = child.tagName ? child.tagName.toLowerCase() : '';
      if (tag === 'br') out += '\n';
      else out += getText(child);
    } else if (child.nodeType === 3) out += child.text || '';
  }
  return decodeText(out);
}

function parseCSS(css) {
  const styles = { body: {}, table: {}, td: {}, rowHeight: {}, rowStyle: {}, cellStyle: {} };
  const bodyStyle = (css.match(/body\s*\{([^}]+)\}/i) || [])[1];
  const tableStyle = (css.match(/table\s*\{([^}]+)\}/i) || [])[1];
  const tdStyle = (css.match(/td\s*\{([^}]+)\}/i) || [])[1];
  const parseDecl = (s) => {
    if (!s) return {};
    const o = {};
    s.split(';').forEach((p) => {
      const i = p.indexOf(':');
      if (i <= 0) return;
      const k = p.slice(0, i).trim().toLowerCase();
      const v = p.slice(i + 1).trim();
      if (k === 'font-family') o.fontFamily = v;
      else if (k === 'font-size') {
        const m = v.match(/(\d+(?:\.\d+)?)\s*pt/i);
        if (m) o.fontSize = parseFloat(m[1]);
      } else if (k === 'font-weight' && /bold/i.test(v)) o.bold = true;
      else if (k === 'font-style' && /italic/i.test(v)) o.italic = true;
      else if (k === 'color') o.color = v.trim();
      else if (k === 'background-color') o.fill = v.trim();
      else if (k === 'text-align') o.hAlign = v.trim().toLowerCase();
      else if (k === 'vertical-align') o.vAlign = v.trim().toLowerCase();
      else if (k === 'height') {
        const m = v.match(/(\d+)\s*px/i);
        if (m) o.heightPx = parseInt(m[1], 10);
      } else if (k.startsWith('border-') && v.includes('solid')) {
        const side = k.replace('border-', '');
        const colorMatch = v.match(/#[0-9a-f]{3,6}/i);
        o['border' + side.charAt(0).toUpperCase() + side.slice(1)] = colorMatch ? colorMatch[0] : '#000000';
      }
    });
    return o;
  };
  Object.assign(styles.body, parseDecl(bodyStyle));
  Object.assign(styles.table, parseDecl(tableStyle));
  Object.assign(styles.td, parseDecl(tdStyle));

  const ruleRe = /tr\.(R\d+)\s*\{([^}]+)\}/gi;
  let m;
  while ((m = ruleRe.exec(css)) !== null) {
    const decl = parseDecl(m[2]);
    if (decl.heightPx) styles.rowHeight[m[1]] = decl.heightPx;
    styles.rowStyle[m[1]] = decl;
  }
  const cellRe = /tr\.(R\d+)\s+td\.(R\d+C\d+)\s*\{([^}]+)\}/gi;
  while ((m = cellRe.exec(css)) !== null) {
    const key = m[1] + ' ' + m[2];
    styles.cellStyle[key] = parseDecl(m[3]);
  }
  return styles;
}

function mergeStyle(base, ...over) {
  const o = { ...base };
  over.forEach((x) => {
    if (!x) return;
    Object.keys(x).forEach((k) => {
      if (x[k] !== undefined) o[k] = x[k];
    });
  });
  return o;
}

function pxToColWidth(px) {
  return Math.max(0, Math.min(255, (px || 64) / 7.5));
}
function pxToRowHeight(px) {
  return Math.max(0, (px || 15) * 0.75);
}

function toArgb(value) {
  if (!value || typeof value !== 'string') return null;
  const hex = value.trim().replace(/^#/, '');
  if (/^[0-9a-f]{6}$/i.test(hex)) return 'FF' + hex.toUpperCase();
  if (/^[0-9a-f]{8}$/i.test(hex)) return hex.toUpperCase();
  return null;
}

function htmlToXlsx(html, log) {
  const root = parse(html, { comment: false });
  const styleEl = root.querySelector('style');
  const cssText = (styleEl && styleEl.textContent) || '';
  const sheetStyles = parseCSS(cssText);

  const tables = root.querySelectorAll('table');
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1', { views: [{ state: 'frozen', ySplit: 0 }] });

  let globalRow = 0;
  const styleCache = new Map();
  const cacheKey = (s) => JSON.stringify(s);
  const getExcelStyle = (cellStyle) => {
    const key = cacheKey(cellStyle);
    if (styleCache.has(key)) return styleCache.get(key);
    const fontArgb = toArgb(cellStyle.color);
    const fillArgb = toArgb(cellStyle.fill);
    const o = {
      font: {
        name: cellStyle.fontFamily || 'Arial',
        size: cellStyle.fontSize || 8,
        bold: !!cellStyle.bold,
        italic: !!cellStyle.italic,
        color: fontArgb ? { argb: fontArgb } : undefined,
      },
      fill: fillArgb ? { type: 'pattern', pattern: 'solid', fgColor: { argb: fillArgb } } : undefined,
      alignment: {
        horizontal: cellStyle.hAlign === 'center' ? 'center' : cellStyle.hAlign === 'right' ? 'right' : 'left',
        vertical: cellStyle.vAlign === 'middle' || cellStyle.vAlign === 'center' ? 'middle' : cellStyle.vAlign === 'bottom' ? 'bottom' : 'top',
        wrapText: true,
      },
      border: {},
    };
    ['Left', 'Right', 'Top', 'Bottom'].forEach((side) => {
      const c = cellStyle['border' + side];
      const borderArgb = toArgb(c);
      if (borderArgb) o.border[side.toLowerCase()] = { style: 'thin', color: { argb: borderArgb } };
    });
    styleCache.set(key, o);
    return o;
  };
  const applyStyle = (cell, style) => {
    if (style?.font) cell.font = { ...(cell.font || {}), ...style.font };
    if (style?.fill) cell.fill = style.fill;
    if (style?.alignment) cell.alignment = { ...(cell.alignment || {}), ...style.alignment };
    if (style?.border) cell.border = { ...(cell.border || {}), ...style.border };
  };

  for (const table of tables) {
    const cols = table.querySelectorAll('col');
    const colWidths = [];
    cols.forEach((c) => {
      const w = parseInt(c.getAttribute('width'), 10);
      colWidths.push(isNaN(w) ? null : w);
    });
    const rows = table.querySelectorAll('tr');
    let maxCol = 0;
    const grid = [];
    let rowIndex = 0;
    for (const tr of rows) {
      const trClass = (tr.getAttribute('class') || '').trim();
      const rowStyles = mergeStyle(sheetStyles.body, sheetStyles.table, sheetStyles.td, sheetStyles.rowStyle[trClass]);
      const heightPx = sheetStyles.rowHeight[trClass];
      const cells = tr.querySelectorAll('td');
      let colIndex = 0;
      const rowData = [];
      for (const td of cells) {
        while (grid[rowIndex] && grid[rowIndex][colIndex] !== undefined) colIndex++;
        const rowSpan = parseInt(td.getAttribute('rowspan'), 10) || 1;
        const colSpan = parseInt(td.getAttribute('colspan'), 10) || 1;
        const tdClass = (td.getAttribute('class') || '').trim();
        const cellKey = trClass + ' ' + tdClass;
        const cellStyle = mergeStyle(rowStyles, sheetStyles.cellStyle[cellKey]);
        const text = getText(td).trim();
        rowData.push({
          r: globalRow + rowIndex,
          c: colIndex,
          rowSpan,
          colSpan,
          style: cellStyle,
          text,
        });
        for (let r = 0; r < rowSpan; r++) {
          if (!grid[rowIndex + r]) grid[rowIndex + r] = [];
          for (let c = 0; c < colSpan; c++) grid[rowIndex + r][colIndex + c] = r === 0 && c === 0 ? rowData[rowData.length - 1] : true;
        }
        colIndex += colSpan;
        maxCol = Math.max(maxCol, colIndex);
      }
      if (heightPx) sheet.getRow(globalRow + rowIndex + 1).height = pxToRowHeight(heightPx);
      rowIndex++;
    }
    for (let r = 0; r < grid.length; r++) {
      for (let c = 0; c < (grid[r]?.length || 0); c++) {
        const cell = grid[r][c];
        if (cell === true) continue;
        if (!cell) continue;
        const excelRow = globalRow + r + 1;
        const excelCol = c + 1;
        const exCell = sheet.getCell(excelRow, excelCol);
        exCell.value = cell.text || '';
        applyStyle(exCell, getExcelStyle(cell.style));
        if (cell.rowSpan > 1 || cell.colSpan > 1) {
          sheet.mergeCells(excelRow, excelCol, excelRow + cell.rowSpan - 1, excelCol + cell.colSpan - 1);
        }
      }
    }
    for (let c = 0; c < maxCol; c++) {
      const w = colWidths[c];
      if (w != null) sheet.getColumn(c + 1).width = pxToColWidth(w);
    }
    globalRow += grid.length;
  }

  return workbook;
}

async function convertFile(htmlPath, xlsxPath, log) {
  const fs = require('fs').promises;
  const html = await fs.readFile(htmlPath, 'utf8');
  const workbook = htmlToXlsx(html, log);
  await workbook.xlsx.writeFile(xlsxPath);
}

module.exports = { convertFile, htmlToXlsx };
