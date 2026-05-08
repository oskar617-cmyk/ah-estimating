// js/pdf-tools.js
// Two related but distinct capabilities:
//   1. extractPdfText(arrayBuffer) — uses pdf.js loaded via CDN to pull
//      text out of a PDF attachment so we can do Tier 2 reply matching.
//   2. emailBodyToPdfBytes(html, meta) — uses jsPDF (CDN) to render the
//      reply email as a small PDF when the supplier didn't send a PDF
//      themselves. The result is what we save to SharePoint.

import { escapeHtml as _esc } from './ui.js';

// --------- PDF text extraction (pdf.js) ---------

// pdf.js is loaded via <script> in index.html and exposes window.pdfjsLib.
// We configure the worker URL once on first use.
let pdfjsConfigured = false;
function ensurePdfjsConfigured() {
  if (pdfjsConfigured) return;
  if (typeof pdfjsLib === 'undefined') {
    throw new Error('pdf.js not loaded — check index.html CDN script tag');
  }
  // Use the matching worker from the same CDN/version.
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    'https://cdn.jsdelivr.net/npm/pdfjs-dist@4.0.379/build/pdf.worker.min.js';
  pdfjsConfigured = true;
}

// Extract concatenated visible text from all pages of a PDF.
// Returns a single string. Throws if the PDF can't be parsed.
export async function extractPdfText(arrayBuffer) {
  ensurePdfjsConfigured();
  const loadingTask = pdfjsLib.getDocument({ data: new Uint8Array(arrayBuffer) });
  const pdf = await loadingTask.promise;
  const parts = [];
  // Cap pages we read at 20 to keep things snappy on monster PDFs.
  const maxPages = Math.min(pdf.numPages, 20);
  for (let i = 1; i <= maxPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    parts.push(content.items.map(it => it.str).join(' '));
  }
  return parts.join('\n');
}

// --------- Email body → PDF (jsPDF) ---------

// jsPDF is loaded via <script> in index.html and exposes window.jspdf.
// We render the email body as plain text wrapped to the page width.
//
// `options`:
//   subject:   email subject (used as the PDF heading)
//   from:      "Name <addr@x>"
//   to:        "Name <addr@x>"
//   receivedAt: ISO string
//   bodyHtml:  the message body HTML (we strip tags to plain text)
// Returns ArrayBuffer of the generated PDF.
export function emailBodyToPdfBytes({ subject, from, to, receivedAt, bodyHtml }) {
  if (!window.jspdf || !window.jspdf.jsPDF) {
    throw new Error('jsPDF not loaded — check index.html CDN script tag');
  }
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: 'pt', format: 'a4' });
  const pageW = doc.internal.pageSize.getWidth();
  const pageH = doc.internal.pageSize.getHeight();
  const margin = 48;
  const usableW = pageW - margin * 2;
  let y = margin;

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(13);
  doc.text(safeStr(subject || '(no subject)'), margin, y, { maxWidth: usableW });
  y += 22;

  doc.setFont('helvetica', 'normal');
  doc.setFontSize(10);
  doc.setTextColor(100);
  if (from) { doc.text(`From: ${safeStr(from)}`, margin, y); y += 14; }
  if (to) { doc.text(`To: ${safeStr(to)}`, margin, y); y += 14; }
  if (receivedAt) {
    const d = new Date(receivedAt);
    doc.text(`Received: ${d.toLocaleString()}`, margin, y);
    y += 14;
  }
  doc.setTextColor(0);
  y += 8;
  // Divider
  doc.setDrawColor(180);
  doc.line(margin, y, pageW - margin, y);
  y += 14;

  // Body — strip HTML to plain text, preserve paragraph breaks.
  const bodyText = htmlToPlainText(bodyHtml || '');
  doc.setFontSize(11);
  const lines = doc.splitTextToSize(bodyText, usableW);
  for (const line of lines) {
    if (y > pageH - margin) {
      doc.addPage();
      y = margin;
    }
    doc.text(line, margin, y);
    y += 14;
  }

  // Output as ArrayBuffer (jsPDF returns Blob; convert to ArrayBuffer)
  const blob = doc.output('blob');
  return blob.arrayBuffer();
}

function safeStr(s) {
  if (s == null) return '';
  return String(s).replace(/[\u0000-\u001f]/g, ' ');
}

// Convert HTML to plain text preserving paragraph breaks.
// We don't ship a full DOM parser; this heuristic handles the common cases
// (Outlook / Gmail reply formats).
function htmlToPlainText(html) {
  if (!html) return '';
  let s = String(html);
  // Drop scripts/styles entirely
  s = s.replace(/<script[\s\S]*?<\/script>/gi, '');
  s = s.replace(/<style[\s\S]*?<\/style>/gi, '');
  // Treat common block-level closing tags as paragraph breaks
  s = s.replace(/<\/(p|div|br\s*\/?|li|tr|h[1-6])>/gi, '\n');
  s = s.replace(/<br\s*\/?>/gi, '\n');
  // Strip remaining tags
  s = s.replace(/<[^>]+>/g, '');
  // Decode common HTML entities
  s = s
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'");
  // Collapse whitespace but keep paragraph breaks
  s = s.replace(/[ \t]+/g, ' ').replace(/\n{3,}/g, '\n\n').trim();
  return s;
}
