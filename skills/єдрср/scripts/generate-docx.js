#!/usr/bin/env node
/**
 * Шаблон генератора Word-документа для скіла /єдрср
 *
 * Використання:
 *   node generate-docx.js <шлях_до_input.json> <шлях_до_output.docx>
 *
 * Залежності: npm install -g docx
 *
 * Структура input.json:
 *
 * {
 *   "title": "Аналітична довідка",
 *   "subtitle": "Судова практика щодо ...",
 *   "intro": ["абзац 1", "абзац 2", ...],
 *   "questions": [
 *     {
 *       "title": "Питання 1. ...",
 *       "paragraphs": ["абзац", "абзац", ...]
 *     }
 *   ],
 *   "tableColumns": [
 *     { "name": "№", "width": 600 },
 *     { "name": "Реквізити + лінк ЄДРСР", "width": 3200 },
 *     { "name": "Позиція суду + мотиви", "width": 5560 }
 *   ],
 *   "cases": [
 *     {
 *       "num": 1,
 *       "cells": [
 *         [{ "text": "Справа № ...", "bold": true }],
 *         [{ "text": "Посилання на ЄДРСР", "hyperlink": "https://reyestr.court.gov.ua/Review/..." }],
 *         [{ "text": "Позиція: ...", "bold": true }, { "text": " (п. X.Y)" }],
 *         [{ "text": "✅ Верифіковано в ЄДРСР", "color": "006100" }]
 *       ]
 *     }
 *   ],
 *   "methodology": ["абзац", ...],  // опціонально
 *   "outputFileName": "2026.05.13 - тема.docx"
 * }
 */

const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, ExternalHyperlink, HeadingLevel, BorderStyle, WidthType, ShadingType,
  VerticalAlign,
} = require('docx');

// === Параметри документа ===
const FONT = 'Times New Roman';
const SIZE_BODY = 28;        // 14pt
const SIZE_H1 = 36;          // 18pt
const SIZE_H2 = 32;          // 16pt
const SIZE_TABLE = 22;       // 11pt
const PAGE_WIDTH = 11906;    // A4
const PAGE_HEIGHT = 16838;
const MARGIN_TOP = 1440;     // 1 inch
const MARGIN_BOTTOM = 1440;
const MARGIN_LEFT = 1440;
const MARGIN_RIGHT = 1080;   // 0.75 inch — щоб таблиця ширша

const border = { style: BorderStyle.SINGLE, size: 6, color: '808080' };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 100, bottom: 100, left: 120, right: 120 };
const headerShading = { fill: 'D9E2F3', type: ShadingType.CLEAR };

// === Утиліти ===

function bodyParagraph(text, opts = {}) {
  return new Paragraph({
    alignment: opts.alignment || AlignmentType.JUSTIFIED,
    spacing: { line: 360, after: opts.after !== undefined ? opts.after : 160 },
    children: [new TextRun({
      text, font: FONT, size: opts.size || SIZE_BODY,
      bold: opts.bold, italics: opts.italics, color: opts.color,
    })],
  });
}

function headingParagraph(text, level) {
  return new Paragraph({
    heading: level === 1 ? HeadingLevel.HEADING_1 : HeadingLevel.HEADING_2,
    spacing: { before: 360, after: 200 },
    children: [new TextRun({
      text, font: FONT,
      size: level === 1 ? SIZE_H1 : SIZE_H2,
      bold: true,
    })],
  });
}

function buildCellParagraphs(lines, opts = {}) {
  if (!Array.isArray(lines)) lines = [lines];
  if (lines.length === 0) return [new Paragraph({ children: [new TextRun({ text: '', font: FONT, size: SIZE_TABLE })] })];

  return lines.map((runsOrText, idx) => {
    const runs = Array.isArray(runsOrText) ? runsOrText : [{ text: String(runsOrText) }];
    const children = runs.map(r => {
      if (r.hyperlink) {
        return new ExternalHyperlink({
          link: r.hyperlink,
          children: [new TextRun({
            text: r.text, font: FONT, size: SIZE_TABLE,
            color: '0563C1', underline: {},
          })],
        });
      }
      return new TextRun({
        text: r.text, font: FONT, size: SIZE_TABLE,
        bold: r.bold, italics: r.italics, color: r.color,
      });
    });
    return new Paragraph({
      alignment: opts.alignment || AlignmentType.LEFT,
      spacing: { line: 260, after: idx === lines.length - 1 ? 0 : 60 },
      children,
    });
  });
}

function tableCell(lines, width, opts = {}) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    margins: cellMargins,
    shading: opts.shading,
    verticalAlign: VerticalAlign.TOP,
    children: buildCellParagraphs(lines, opts),
  });
}

function headerCell(text, width) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    margins: cellMargins,
    shading: headerShading,
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { line: 260, after: 0 },
      children: [new TextRun({ text, font: FONT, size: SIZE_TABLE, bold: true })],
    })],
  });
}

// === Побудова документа ===

function buildDocument(spec) {
  const children = [];

  // Заголовки
  if (spec.title) {
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 240 },
      children: [new TextRun({ text: spec.title, font: FONT, size: SIZE_H1, bold: true })],
    }));
  }
  if (spec.subtitle) {
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 400 },
      children: [new TextRun({ text: spec.subtitle, font: FONT, size: SIZE_H2, bold: true })],
    }));
  }

  // Вступ
  if (spec.intro && spec.intro.length > 0) {
    spec.intro.forEach(p => children.push(bodyParagraph(p)));
  }

  // Питання з відповідями
  if (spec.questions && spec.questions.length > 0) {
    spec.questions.forEach(q => {
      children.push(headingParagraph(q.title, 1));
      if (q.paragraphs) {
        q.paragraphs.forEach(p => {
          if (typeof p === 'string') {
            children.push(bodyParagraph(p));
          } else if (p.text) {
            children.push(bodyParagraph(p.text, p));
          }
        });
      }
    });
  }

  // Заголовок таблиці
  if (spec.tableHeading) {
    children.push(headingParagraph(spec.tableHeading, 1));
  }

  // Опис перед таблицею
  if (spec.tableIntro) {
    children.push(bodyParagraph(spec.tableIntro, { italics: true, after: 240 }));
  }

  // Таблиця
  if (spec.tableColumns && spec.cases) {
    const columnWidths = spec.tableColumns.map(c => c.width);
    const tableWidth = columnWidths.reduce((a, b) => a + b, 0);

    const headerRow = new TableRow({
      tableHeader: true,
      children: spec.tableColumns.map(c => headerCell(c.name, c.width)),
    });

    const dataRows = spec.cases.map((c, rowIdx) => {
      let cells = c.cells;
      if (c.num != null && cells.length < columnWidths.length) {
        cells = [[[{ text: String(c.num), bold: true }]], ...cells];
      }
      return new TableRow({
        children: cells.map((cellContent, idx) => tableCell(cellContent, columnWidths[idx])),
      });
    });

    children.push(new Table({
      width: { size: tableWidth, type: WidthType.DXA },
      columnWidths,
      rows: [headerRow, ...dataRows],
    }));
  }

  // Методологія (опціонально)
  if (spec.methodology && spec.methodology.length > 0) {
    children.push(new Paragraph({
      spacing: { before: 360, after: 120 },
      children: [new TextRun({
        text: 'Методологія та застереження',
        font: FONT, size: SIZE_H2, bold: true,
      })],
    }));
    spec.methodology.forEach(p => children.push(bodyParagraph(p, { italics: true })));
  }

  // Дата
  const today = new Date();
  const dateStr = `${today.getDate().toString().padStart(2, '0')} ${[
    'січня','лютого','березня','квітня','травня','червня',
    'липня','серпня','вересня','жовтня','листопада','грудня'
  ][today.getMonth()]} ${today.getFullYear()} року`;
  children.push(bodyParagraph(`Дата підготовки документа: ${dateStr}.`, { italics: true, after: 0 }));

  // Документ
  return new Document({
    styles: {
      default: { document: { run: { font: FONT, size: SIZE_BODY } } },
      paragraphStyles: [
        {
          id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal',
          quickFormat: true,
          run: { size: SIZE_H2, bold: true, font: FONT },
          paragraph: { spacing: { before: 240, after: 200 }, outlineLevel: 0 },
        },
      ],
    },
    sections: [{
      properties: {
        page: {
          size: { width: PAGE_WIDTH, height: PAGE_HEIGHT },
          margin: { top: MARGIN_TOP, right: MARGIN_RIGHT, bottom: MARGIN_BOTTOM, left: MARGIN_LEFT },
        },
      },
      children,
    }],
  });
}

// === Main ===

if (require.main === module) {
  const inputPath = process.argv[2];
  const outputPath = process.argv[3];

  if (!inputPath || !outputPath) {
    console.error('Usage: node generate-docx.js <input.json> <output.docx>');
    process.exit(1);
  }

  let spec;
  try {
    spec = JSON.parse(fs.readFileSync(inputPath, 'utf8'));
  } catch (e) {
    console.error('ERROR: Invalid JSON in', inputPath, '—', e.message);
    process.exit(1);
  }
  if (!spec.tableColumns || !spec.cases) {
    console.error('ERROR: JSON must contain "tableColumns" and "cases" fields.');
    process.exit(1);
  }
  const doc = buildDocument(spec);

  Packer.toBuffer(doc).then(buffer => {
    fs.writeFileSync(outputPath, buffer);
    console.log(`OK: ${outputPath}`);
  }).catch(err => {
    console.error('ERROR:', err);
    process.exit(1);
  });
}

module.exports = { buildDocument };
