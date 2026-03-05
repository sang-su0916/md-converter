import { NextRequest, NextResponse } from 'next/server';
import { writeFile, readFile, unlink, mkdir } from 'fs/promises';
import { join } from 'path';
import { execFile } from 'child_process';
import { promisify } from 'util';
import { tmpdir } from 'os';
import { randomUUID } from 'crypto';

const execFileAsync = promisify(execFile);

// CORS: allow Vercel frontend to call Render API directly
const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

// If RENDER_API_URL is set, proxy to Render backend (Vercel mode)
const RENDER_API_URL = process.env.RENDER_API_URL;

const SUPPORTED_EXTENSIONS = [
  '.pdf', '.docx', '.doc', '.pptx', '.ppt', '.xlsx', '.xls',
  '.html', '.htm', '.csv', '.json', '.xml', '.rtf', '.epub',
  '.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.webp',
  '.mp3', '.wav', '.m4a',
  '.zip',
  '.hwp', '.hwpx',
  '.txt', '.md', '.rst', '.log',
];

const HWP_EXTENSIONS = ['.hwp', '.hwpx'];
const TEXT_EXTENSIONS = ['.txt', '.md', '.rst', '.log'];
const PDF_EXTENSIONS = ['.pdf'];
const OFFICE_EXTENSIONS = ['.docx', '.doc', '.pptx', '.ppt', '.xlsx', '.xls', '.rtf', '.epub'];
const HTML_EXTENSIONS = ['.html', '.htm', '.xml'];

function getExtension(filename: string): string {
  const idx = filename.lastIndexOf('.');
  return idx >= 0 ? filename.slice(idx).toLowerCase() : '';
}

const HOME = process.env.HOME || '/root';
const ENV = {
  ...process.env,
  HOME,
  PATH: `${HOME}/.local/bin:/opt/homebrew/bin:/usr/local/bin:${process.env.PATH || '/usr/bin:/bin'}`,
};

const RENDER_BACKEND = process.env.RENDER_BACKEND_URL || 'https://md-converter-ghdf.onrender.com';

/**
 * Proxy conversion to Render backend (Docker with full Python tools)
 */
async function proxyToRender(file: File, ext?: string): Promise<Response> {
  const formData = new FormData();
  formData.append('file', file);

  const res = await fetch(`${RENDER_BACKEND}/api/convert`, {
    method: 'POST',
    body: formData,
    signal: AbortSignal.timeout(180000), // 3min timeout for cold start
  });

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const data = await res.json() as any;

  // Post-process Render response for PDF and HTML
  if (data.markdown && ext) {
    if (PDF_EXTENSIONS.includes(ext)) {
      data.markdown = postProcessPdfMarkdown(data.markdown);
      data.lineCount = data.markdown.split('\n').length;
      data.charCount = data.markdown.length;
    } else if (HTML_EXTENSIONS.includes(ext)) {
      data.markdown = postProcessHtmlMarkdown(data.markdown);
      data.lineCount = data.markdown.split('\n').length;
      data.charCount = data.markdown.length;
    }
  }

  return jsonWithCors(data, res.status);
}

let markitdownAvailable: boolean | null = null;

async function checkMarkitdown(): Promise<string | null> {
  if (markitdownAvailable === false) return null;
  const candidates = [
    join(HOME, '.local', 'bin', 'markitdown'),
    '/opt/homebrew/bin/markitdown',
    '/usr/local/bin/markitdown',
    'markitdown',
  ];
  for (const p of candidates) {
    try {
      await execFileAsync(p, ['--version'], { timeout: 5000 });
      markitdownAvailable = true;
      return p;
    } catch { continue; }
  }
  markitdownAvailable = false;
  return null;
}

async function convertPdfToMarkdown(filePath: string): Promise<string> {
  let rawText = '';

  // Strategy 1: pdf-parse (stable on serverless)
  try {
    // eslint-disable-next-line @typescript-eslint/no-require-imports
    const pdfParse = require('pdf-parse');
    const buffer = await readFile(filePath);
    const result = await pdfParse(buffer);
    if (result.text && result.text.trim().length > 0) {
      rawText = result.text;
    }
  } catch (e: unknown) {
    console.error('pdf-parse failed:', e instanceof Error ? e.message : e);
  }

  // Strategy 2: unpdf fallback
  if (!rawText) {
    try {
      const { extractText, getDocumentProxy } = await import('unpdf');
      const buffer = await readFile(filePath);
      const pdf = await getDocumentProxy(new Uint8Array(buffer));
      const { text } = await extractText(pdf, { mergePages: true });
      if (text && text.trim().length > 0) rawText = text;
    } catch (e: unknown) {
      console.error('unpdf failed:', e instanceof Error ? e.message : e);
    }
  }

  if (!rawText) return '';

  // Post-process: structure the raw text into readable markdown
  const formatted = formatPdfText(rawText);
  // DEBUG: check if headings were inserted
  const debugHeadings = formatted.split('\n').filter(l => l.startsWith('#')).length;
  console.log(`[PDF DEBUG] rawText lines: ${rawText.split('\n').length}, formatted lines: ${formatted.split('\n').length}, headings: ${debugHeadings}`);
  return formatted;
}

/**
 * Format raw PDF text into structured markdown
 * Handles: paragraph breaks, headings, list items, tables
 */
function formatPdfText(text: string): string {
  // Normalize line endings
  let result = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');

  // If text is essentially one huge line, split on sentence boundaries
  const lines = result.split('\n');
  if (lines.length <= 3 && result.length > 1000) {
    // Single-line dump: split on double spaces or period+space patterns
    result = result
      .replace(/\s{3,}/g, '\n\n')  // Triple+ spaces → paragraph break
      .replace(/\s{2}/g, '\n')      // Double spaces → line break
      .replace(/([.!?])\s+([A-Z가-힣ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ①②③④⑤⑥⑦⑧⑨⑩])/g, '$1\n\n$2')  // Sentence end + capital/Korean start
      .replace(/(·{3,}|\.{3,})\s*/g, '\n')  // Dots/middots as separators
      .replace(/(\d+)\s*페이지/g, '\n---\n')  // Page numbers
      ;
  }

  // Split into lines for processing
  const outputLines = result.split('\n');
  const formatted: string[] = [];

  for (const line of outputLines) {
    const trimmed = line.trim();
    if (!trimmed) {
      formatted.push('');
      continue;
    }

    // Roman numeral headings: Ⅰ. Ⅱ. etc. (anywhere in line)
    if (/[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ][\.\s]/.test(trimmed) && trimmed.length < 80) {
      formatted.push('', `# ${trimmed}`, '');
      continue;
    }

    // Numeric major headings: "1. 제목" (short, likely heading)
    if (/^\d+\.\s/.test(trimmed) && trimmed.length < 80 && !/[,;]/.test(trimmed)) {
      formatted.push('', `## ${trimmed}`, '');
      continue;
    }

    // Korean chapter/article patterns
    if (/제\s*\d+\s*장/.test(trimmed) && trimmed.length < 80) {
      formatted.push('', `## ${trimmed}`, '');
      continue;
    }
    if (/제\s*\d+\s*조[\s(]/.test(trimmed) && trimmed.length < 100) {
      formatted.push('', `### ${trimmed}`, '');
      continue;
    }

    // CONTENTS/목차/차례 - any line containing these keywords
    if (/CONTENTS|목차|차례|TABLE OF CONTENTS/i.test(trimmed) && trimmed.length < 100) {
      formatted.push('', `# ${trimmed}`, '');
      continue;
    }

    // ALL-CAPS or short bold-like lines (likely section titles)
    if (trimmed.length < 50 && /^[A-Z\s]+$/.test(trimmed) && trimmed.length > 3) {
      formatted.push('', `## ${trimmed}`, '');
      continue;
    }

    // Short standalone lines ending with "은" "는" "의" "금" (Korean topic markers - likely section titles)
    if (trimmed.length > 3 && trimmed.length < 60 && /[은는의금]$/.test(trimmed)
        && !trimmed.includes(',') && !trimmed.includes(';')) {
      // Check if surrounded by blank lines or at boundaries
      const prevBlank = formatted.length === 0 || formatted[formatted.length - 1] === '';
      const isLikelyHeading = prevBlank || trimmed.includes('장려금') || trimmed.includes('지원금');
      if (isLikelyHeading) {
        formatted.push('', `## ${trimmed}`, '');
        continue;
      }
    }

    // Circled number items: ① ② etc.
    if (/^[①②③④⑤⑥⑦⑧⑨⑩]/.test(trimmed)) {
      formatted.push('', trimmed);
      continue;
    }

    // Bullet-like markers
    if (/^[○●■□▶▷◆◇·•-]\s/.test(trimmed)) {
      formatted.push(`- ${trimmed.slice(2)}`);
      continue;
    }

    // Short standalone lines (likely titles/labels)
    if (trimmed.length < 30 && !trimmed.endsWith('.') && !trimmed.endsWith(',')
        && !/^\d/.test(trimmed) && formatted.length > 0
        && formatted[formatted.length - 1] === '') {
      formatted.push(`**${trimmed}**`);
      formatted.push('');
      continue;
    }

    formatted.push(trimmed);
  }

  return formatted.join('\n')
    .replace(/\n{4,}/g, '\n\n\n')
    .trim() + '\n';
}

/**
 * Post-process PDF markdown: add headings regardless of extraction method
 * Applied AFTER all PDF processing (markitdown, pdf-parse, etc.)
 */
function postProcessPdfMarkdown(md: string): string {
  const lines = md.split('\n');
  const result: string[] = [];

  for (let i = 0; i < lines.length; i++) {
    const trimmed = lines[i].trim();
    if (!trimmed) { result.push(''); continue; }

    // Skip lines already formatted as headings
    if (trimmed.startsWith('#')) { result.push(lines[i]); continue; }

    // Form feed characters → page breaks (remove)
    if (trimmed === '\f' || trimmed === '') { result.push(''); continue; }

    // Roman numeral headings: Ⅰ. Ⅱ. Ⅲ. etc.
    if (/[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ][.\s]/.test(trimmed) && trimmed.length < 120) {
      // Remove trailing dots (TOC separators)
      const clean = trimmed.replace(/\s*[·.]{5,}\s*\d*\s*$/, '').trim();
      result.push('', `# ${clean}`, '');
      continue;
    }

    // CONTENTS/목차
    if (/^(CONTENTS|목차|차례)/i.test(trimmed) && trimmed.length < 100) {
      result.push('', `# ${trimmed}`, '');
      continue;
    }

    // Numeric headings: "1. 제목" (short, no commas)
    if (/^\d+\.\s/.test(trimmed) && trimmed.length < 80 && !/[,;]/.test(trimmed)) {
      const clean = trimmed.replace(/\s*[·.]{5,}\s*\d*\s*$/, '').trim();
      result.push('', `## ${clean}`, '');
      continue;
    }

    // Korean chapter/section/article
    if (/^제\s*\d+\s*장/.test(trimmed) && trimmed.length < 80) {
      result.push('', `## ${trimmed}`, '');
      continue;
    }
    if (/^제\s*\d+\s*절/.test(trimmed) && trimmed.length < 80) {
      result.push('', `### ${trimmed}`, '');
      continue;
    }
    if (/^제\s*\d+\s*조[\s(]/.test(trimmed) && trimmed.length < 100) {
      result.push('', `#### ${trimmed}`, '');
      continue;
    }

    // Short standalone lines ending with topic markers (은/는/의/금)
    if (trimmed.length > 3 && trimmed.length < 60 && /[은는금]$/.test(trimmed)) {
      const prevBlank = i === 0 || lines[i - 1].trim() === '' || lines[i - 1].trim() === '\f';
      const nextBlank = i === lines.length - 1 || lines[i + 1].trim() === '';
      if (prevBlank && nextBlank) {
        result.push('', `## ${trimmed}`, '');
        continue;
      }
    }

    // ALL-CAPS short lines
    if (trimmed.length > 3 && trimmed.length < 50 && /^[A-Z\s]+$/.test(trimmed)) {
      result.push('', `## ${trimmed}`, '');
      continue;
    }

    // Bullet markers → list
    if (/^[○●■□▶▷◆◇·•]\s/.test(trimmed)) {
      result.push(`- ${trimmed.slice(2)}`);
      continue;
    }

    // Circled numbers
    if (/^[①②③④⑤⑥⑦⑧⑨⑩]/.test(trimmed)) {
      result.push('', trimmed);
      continue;
    }

    result.push(trimmed);
  }

  return result.join('\n').replace(/\n{4,}/g, '\n\n\n').replace(/\f/g, '').trim() + '\n';
}

/**
 * Post-process HTML markdown: add headings for TOPIC, bold sections, etc.
 * Applied AFTER turndown conversion
 */
function postProcessHtmlMarkdown(md: string): string {
  const lines = md.split('\n');
  const result: string[] = [];

  for (let i = 0; i < lines.length; i++) {
    const trimmed = lines[i].trim();

    // TOPIC XX → heading
    if (/^TOPIC\s+\d+/i.test(trimmed)) {
      result.push('', `## ${trimmed}`, '');
      continue;
    }

    // FROM THE CEO, NEWSLETTER etc.
    if (/^(FROM THE|NEWSLETTER|SPRING|SUMMER|FALL|WINTER)\s/i.test(trimmed) && trimmed.length < 60) {
      result.push('', `## ${trimmed}`, '');
      continue;
    }

    // Standalone bold lines → section headings
    if (/^\*\*[^*]+\*\*$/.test(trimmed) && trimmed.length < 100 && !trimmed.includes('http')) {
      const title = trimmed.replace(/\*\*/g, '');
      const prevBlank = i === 0 || lines[i - 1].trim() === '';
      if (prevBlank && title.length > 3 && title.length < 80) {
        result.push('', `## ${title}`, '');
        continue;
      }
    }

    result.push(lines[i]);
  }

  return result.join('\n').replace(/\n{4,}/g, '\n\n\n').trim() + '\n';
}

async function convertWithOfficeParser(filePath: string): Promise<string> {
  const officeparser = await import('officeparser');
  const result = await officeparser.parseOffice(filePath);
  if (result && typeof result === 'object' && 'content' in result) {
    return contentToMarkdown(result.content as ContentNode[], result.type as string);
  }
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const res = result as any;
  if (res && typeof res.toText === 'function') return res.toText();
  return typeof result === 'string' ? result : String(result);
}

interface ContentNode { type: string; text?: string; children?: ContentNode[]; metadata?: Record<string, unknown>; }

function contentToMarkdown(content: ContentNode[], docType?: string): string {
  const lines: string[] = [];
  for (const node of content) {
    switch (node.type) {
      case 'heading': { const level = (node.metadata?.level as number) || 1; lines.push(`${'#'.repeat(level)} ${node.text || ''}`); lines.push(''); break; }
      case 'paragraph': { const text = node.text || ''; if (text.trim()) { lines.push(text); lines.push(''); } break; }
      case 'table': {
        const tableRows = node.children || [];
        if (tableRows.length === 0) break;
        for (let i = 0; i < tableRows.length; i++) {
          const row = tableRows[i];
          const cells = (row.children || []).map(c => (c.text || '').replace(/\|/g, '\\|').replace(/\n/g, ' '));
          lines.push(`| ${cells.join(' | ')} |`);
          if (i === 0) lines.push(`| ${cells.map(() => '---').join(' | ')} |`);
        }
        lines.push(''); break;
      }
      case 'list': { const items = node.children || []; for (const item of items) lines.push(`- ${item.text || ''}`); lines.push(''); break; }
      case 'sheet': {
        const sheetName = (node.metadata?.name as string) || '';
        if (sheetName) lines.push(`## ${sheetName}`);
        const rows = node.children || [];
        if (rows.length > 0) {
          for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            const cells = (row.children || []).map(c => (c.text || '').replace(/\|/g, '\\|'));
            lines.push(`| ${cells.join(' | ')} |`);
            if (i === 0) lines.push(`| ${cells.map(() => '---').join(' | ')} |`);
          }
          lines.push('');
        }
        break;
      }
      case 'slide': { const slideNum = (node.metadata?.number as number) || ''; lines.push(`## 슬라이드 ${slideNum}`); if (node.children) lines.push(contentToMarkdown(node.children, docType)); break; }
      default: {
        if (node.text?.trim()) { lines.push(node.text); lines.push(''); }
        else if (node.children) lines.push(contentToMarkdown(node.children, docType));
      }
    }
  }
  return lines.join('\n').replace(/\n{3,}/g, '\n\n').trim() + '\n';
}

async function convertHtmlToMarkdown(htmlContent: string): Promise<string> {
  const TurndownService = (await import('turndown')).default;
  const turndown = new TurndownService({ headingStyle: 'atx', codeBlockStyle: 'fenced', bulletListMarker: '-' });

  let md = turndown.turndown(htmlContent);

  // Post-process: detect heading-like patterns in the output
  const lines = md.split('\n');
  const result: string[] = [];
  
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const trimmed = line.trim();

    // TOPIC XX pattern → heading
    if (/^TOPIC\s+\d+/i.test(trimmed)) {
      result.push('', `## ${trimmed}`, '');
      continue;
    }

    // Short bold lines that look like section titles (isolated, standalone)
    if (/^\*\*[^*]+\*\*$/.test(trimmed) && trimmed.length < 100
        && !trimmed.includes('http') && !trimmed.includes('@')) {
      const title = trimmed.replace(/\*\*/g, '');
      // Check if it's isolated (surrounded by blank lines or start/end)
      const prevBlank = i === 0 || lines[i - 1].trim() === '';
      const nextBlank = i === lines.length - 1 || lines[i + 1]?.trim() === '';
      if (prevBlank && title.length < 80 && title.length > 3) {
        result.push('', `## ${title}`, '');
        continue;
      }
    }

    // Korean section markers ending with 은/는/의/다 (standalone lines)
    if (trimmed.length > 3 && trimmed.length < 80 && /[은는의다]$/.test(trimmed)
        && !trimmed.includes(',') && !trimmed.includes('@')) {
      const prevBlank = i === 0 || lines[i - 1].trim() === '';
      if (prevBlank) {
        result.push('', `## ${trimmed}`, '');
        continue;
      }
    }

    // All-caps lines (likely headings)
    if (trimmed.length > 3 && trimmed.length < 50 && /^[A-Z\s]+$/.test(trimmed)) {
      result.push('', `## ${trimmed}`, '');
      continue;
    }

    result.push(line);
  }

  return result.join('\n')
    .replace(/\n{4,}/g, '\n\n\n')
    .trim() + '\n';
}

function convertCsvToMarkdown(csvContent: string): string {
  const lines = csvContent.trim().split('\n');
  if (lines.length === 0) return csvContent;
  const rows = lines.map(line => {
    const cells: string[] = []; let current = ''; let inQuotes = false;
    for (const ch of line) {
      if (ch === '"') inQuotes = !inQuotes;
      else if (ch === ',' && !inQuotes) { cells.push(current.trim()); current = ''; }
      else current += ch;
    }
    cells.push(current.trim()); return cells;
  });
  if (rows.length === 0) return csvContent;
  const header = `| ${rows[0].join(' | ')} |`;
  const separator = `| ${rows[0].map(() => '---').join(' | ')} |`;
  const body = rows.slice(1).map(r => `| ${r.join(' | ')} |`).join('\n');
  return [header, separator, body].join('\n') + '\n';
}

/**
 * Convert only layout tables to plain text, preserve data tables as markdown
 * Layout tables: 2-column with legal patterns (제X조/제X장 + [필수]/[선택]/착안사항/☞)
 * Data tables: 3+ columns, or non-legal 2-column → kept as markdown table
 * Used for HWP conversion via markitdown path
 */
function convertLayoutTablesToText(markdown: string): string {
  const lines = markdown.split('\n');
  const result: string[] = [];
  let tableBlock: string[] = [];
  let inTable = false;

  const flushTable = () => {
    if (tableBlock.length === 0) return;

    // Parse content rows (skip separators) for analysis
    const contentRows: string[][] = [];
    for (const tl of tableBlock) {
      if (!isTableSeparator(tl)) {
        contentRows.push(parseTableRow(tl));
      }
    }

    const maxCols = contentRows.length > 0 ? Math.max(...contentRows.map(r => r.length)) : 0;

    // Check if layout table: 2 columns with legal patterns
    let isLayout = false;
    if (maxCols === 2 && contentRows.length >= 2) {
      let score = 0;
      for (const row of contentRows) {
        const left = (row[0] || '').trim();
        const right = (row[1] || '').trim();
        if (/제\d+조|제\d+장|제\d+절/.test(left)) score++;
        if (/\[필수\]|\[선택\]|\[필수,\s*선택\]|\[선택,\s*필수\]|착안사항|☞|◈/.test(right)) score++;
      }
      isLayout = score >= 3;
    }

    if (isLayout) {
      // Layout table → text, mark right column as annotation
      for (const row of contentRows) {
        const left = (row[0] || '').trim();
        const right = (row[1] || '').trim();
        if (left) result.push(left);
        if (right) {
          const alreadyMarked = /^\[필수\]|^\[선택\]|^\[필수,|^\[선택,|^◈|^☞|^착안사항|^※/.test(right);
          result.push(alreadyMarked ? right : `◈ ${right}`);
        }
      }
      result.push('');
    } else {
      // Data/form table → keep as markdown table
      for (const tl of tableBlock) {
        result.push(tl);
      }
      result.push('');
    }

    tableBlock = [];
    inTable = false;
  };

  for (const line of lines) {
    const trimmed = line.trim();

    if (/^\|.*\|$/.test(trimmed)) {
      inTable = true;
      tableBlock.push(trimmed);
      continue;
    }

    // Not a table row - flush accumulated table
    if (inTable) {
      flushTable();
    }

    result.push(line);
  }

  // Flush any remaining table
  if (inTable) flushTable();

  return result.join('\n');
}

/**
 * Parse HTML table inner content into rows of cell strings
 */
function parseHtmlTableToRows(tableInner: string): string[][] {
  const rows: string[][] = [];
  const rowRegex = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
  let rowMatch;
  while ((rowMatch = rowRegex.exec(tableInner)) !== null) {
    const cells: string[] = [];
    const cellRegex = /<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/gi;
    let cellMatch;
    while ((cellMatch = cellRegex.exec(rowMatch[1])) !== null) {
      const cellText = cellMatch[1]
        .replace(/<br\s*\/?>/gi, ' ')
        .replace(/<[^>]+>/g, '')
        .replace(/&nbsp;/g, ' ')
        .replace(/&amp;/g, '&')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&#39;/g, "'")
        .replace(/\s+/g, ' ')
        .trim();
      cells.push(cellText);
    }
    if (cells.length > 0) rows.push(cells);
  }
  return rows;
}

/**
 * Convert an HTML table block to either markdown table (data) or plain text (layout)
 * Layout tables: 2-column legal document format → text with annotation markers
 * Data tables: 3+ columns or non-legal → markdown table syntax
 */
function htmlTableToMarkdownOrText(tableInner: string): string {
  const rows = parseHtmlTableToRows(tableInner);
  if (rows.length === 0) return '\n';

  const maxCols = Math.max(...rows.map(r => r.length));

  // Check if this is a layout table (2 cols with legal patterns)
  if (maxCols <= 2 && rows.length >= 3) {
    let score = 0;
    for (const row of rows) {
      const left = row[0] || '';
      const right = row[1] || '';
      if (/제\d+조|제\d+장|제\d+절/.test(left)) score++;
      if (/\[필수\]|\[선택\]|\[필수,\s*선택\]|\[선택,\s*필수\]|착안사항|☞|◈/.test(right)) score++;
    }
    if (score >= 3) {
      // Layout table → text extraction, mark right column as annotation
      const textLines: string[] = [];
      for (const row of rows) {
        const left = (row[0] || '').trim();
        const right = (row[1] || '').trim();
        if (left) textLines.push(left);
        if (right) {
          const alreadyMarked = /^\[필수\]|^\[선택\]|^\[필수,|^\[선택,|^◈|^☞|^착안사항|^※/.test(right);
          textLines.push(alreadyMarked ? right : `◈ ${right}`);
        }
      }
      return '\n' + textLines.join('\n') + '\n\n';
    }
  }

  // Data table (3+ cols or non-legal 2-col) → markdown table
  const normalized = rows.map(row => {
    const r = [...row];
    while (r.length < maxCols) r.push('');
    return r;
  });

  const mdLines: string[] = [''];
  mdLines.push('| ' + normalized[0].map(c => c.replace(/\|/g, '\\|')).join(' | ') + ' |');
  mdLines.push('| ' + normalized[0].map(() => '---').join(' | ') + ' |');
  for (let i = 1; i < normalized.length; i++) {
    mdLines.push('| ' + normalized[i].map(c => c.replace(/\|/g, '\\|')).join(' | ') + ' |');
  }
  mdLines.push('');

  return mdLines.join('\n');
}

/**
 * Extract clean text from HWP-generated HTML
 * Preserves data tables as markdown, converts layout tables to text
 * This avoids markitdown's faithful table conversion which makes layout tables unreadable
 */
function extractTextFromHwpHtml(html: string): string {
  let text = html;

  // Remove <head> section entirely
  text = text.replace(/<head[\s\S]*?<\/head>/gi, '');

  // Remove <style> and <script> tags
  text = text.replace(/<style[\s\S]*?<\/style>/gi, '');
  text = text.replace(/<script[\s\S]*?<\/script>/gi, '');

  // Process HTML tables: data tables → markdown, layout tables → text
  // Handle innermost tables first (no nested tables inside), iterate outward
  let tableIterations = 0;
  while (/<table[^>]*>((?:(?!<table)[\s\S])*?)<\/table>/i.test(text) && tableIterations < 20) {
    text = text.replace(/<table[^>]*>((?:(?!<table)[\s\S])*?)<\/table>/gi, (_match, inner) => {
      return htmlTableToMarkdownOrText(inner);
    });
    tableIterations++;
  }

  // Convert <br> to newline
  text = text.replace(/<br\s*\/?>/gi, '\n');

  // Add paragraph breaks at block boundaries
  text = text.replace(/<\/p>/gi, '\n\n');
  text = text.replace(/<\/div>/gi, '\n');
  text = text.replace(/<\/h[1-6]>/gi, '\n\n');
  text = text.replace(/<\/li>/gi, '\n');
  text = text.replace(/<\/blockquote>/gi, '\n\n');

  // Table cell boundaries → newline (key for HWP layout tables)
  text = text.replace(/<\/td>/gi, '\n');
  text = text.replace(/<\/th>/gi, '\n');

  // Table row boundaries → double newline (paragraph break)
  text = text.replace(/<\/tr>/gi, '\n\n');

  // Table start/end → paragraph break
  text = text.replace(/<\/?table[^>]*>/gi, '\n\n');
  text = text.replace(/<\/?tbody[^>]*>/gi, '');
  text = text.replace(/<\/?thead[^>]*>/gi, '');

  // Strip all remaining HTML tags
  text = text.replace(/<[^>]+>/g, '');

  // Decode HTML entities
  text = text
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&nbsp;/g, ' ')
    .replace(/&#8203;/g, '')          // Zero-width space
    .replace(/&#x200B;/g, '')         // Zero-width space (hex)
    .replace(/&#(\d+);/g, (_, code) => {
      const n = parseInt(code, 10);
      return n > 31 && n < 127 ? String.fromCharCode(n) : '';
    });

  // Clean up whitespace
  text = text
    .replace(/[ \t]+/g, ' ')           // Collapse horizontal whitespace
    .replace(/^ +| +$/gm, '')          // Trim each line
    .replace(/\n{3,}/g, '\n\n')        // Max 2 consecutive newlines
    .trim();

  return text;
}

/**
 * Parse a markdown table row into cells, handling escaped pipes
 */
function parseTableRow(row: string): string[] {
  const inner = row.replace(/^\|/, '').replace(/\|$/, '');
  return inner.split('|').map(c => c.trim());
}

/**
 * Check if a line is a markdown table separator (| --- | --- |)
 */
function isTableSeparator(line: string): boolean {
  return /^\|[\s\-:]+(\|[\s\-:]+)+\|?$/.test(line.trim());
}

/**
 * Detect if a table is a 2-column legal document layout table
 * (조문 on left, 착안사항/참고 on right)
 */
function isLegalLayoutTable(rows: string[][]): boolean {
  if (rows.length < 3) return false;
  const cols = rows[0]?.length || 0;
  if (cols !== 2) return false;
  // Check if content matches legal document patterns
  let legalPatterns = 0;
  for (const row of rows) {
    const left = row[0] || '';
    const right = row[1] || '';
    if (/제\d+조/.test(left) || /제\d+장/.test(left) || /제\d+절/.test(left)) legalPatterns++;
    if (/\[필수\]|\[선택\]|착안사항|참고\)|☞/.test(right)) legalPatterns++;
  }
  return legalPatterns >= 3;
}

/**
 * Convert a legal layout table into structured markdown
 */
function convertLegalTable(rows: string[][]): string {
  const output: string[] = [];

  for (const row of rows) {
    const left = (row[0] || '').trim();
    const right = (row[1] || '').trim();

    if (!left && !right) continue;

    // Skip header rows like "취업규칙(안) | (작성시 착안사항)"
    if (/^취업규칙/.test(left) && /착안사항/.test(right)) continue;
    if (/^취업규칙/.test(left) && /^취업규칙/.test(right)) continue;

    // Chapter heading: 제X장
    const chapterMatch = left.match(/^(제\d+장\s+.+?)$/);
    if (chapterMatch && left.length < 50 && !/제\d+조/.test(left)) {
      output.push('', `## ${chapterMatch[1].trim()}`, '');
      if (right && right.length > 5) {
        output.push(`> **착안사항**: ${right}`, '');
      }
      continue;
    }

    // Section heading: 제X절
    const sectionMatch = left.match(/^(제\d+절\s+.+?)$/);
    if (sectionMatch && left.length < 50) {
      output.push('', `### ${sectionMatch[1].trim()}`, '');
      if (right && right.length > 5) {
        output.push(`> ${right}`, '');
      }
      continue;
    }

    // Article: 제X조(제목) + body text
    const articleMatch = left.match(/^(제\d+조(?:의\d+)?\([^)]+\))\s*([\s\S]*)/);
    if (articleMatch) {
      const title = articleMatch[1].trim();
      let body = articleMatch[2].trim();

      output.push('', `#### ${title}`, '');

      if (body) {
        // Split body into paragraphs by clause markers
        body = formatLegalBody(body);
        output.push(body, '');
      }

      if (right && right.length > 5) {
        output.push(`> **착안사항**: ${right}`, '');
      }
      continue;
    }

    // TOC or other structured content - just output as text
    if (left) {
      // Check if it's a TOC block (contains multiple 제X조 references)
      const articleRefs = left.match(/제\d+조/g);
      if (articleRefs && articleRefs.length > 3) {
        // It's a TOC block - format as list
        const tocLines = left.split(/\s{2,}/).filter(l => l.trim());
        for (const tocLine of tocLines) {
          const tl = tocLine.trim();
          if (/^제\d+장/.test(tl)) output.push(`\n**${tl}**`);
          else if (/^제\d+절/.test(tl)) output.push(`  *${tl}*`);
          else if (/^제\d+조/.test(tl)) output.push(`  - ${tl}`);
          else output.push(`  ${tl}`);
        }
      } else {
        output.push(left);
      }
    }

    if (right && right.length > 5 && !left.includes(right)) {
      // Standalone right column content (착안사항 without left content)
      if (/^\[필수\]|\[선택\]|☞|참고/.test(right)) {
        output.push(`> ${right}`);
      } else {
        output.push(right);
      }
    }
  }

  return output.join('\n');
}

/**
 * Format legal body text with proper paragraph breaks
 */
function formatLegalBody(text: string): string {
  // Insert line breaks before clause markers
  let result = text
    .replace(/\s+(①|②|③|④|⑤|⑥|⑦|⑧|⑨|⑩|⑪|⑫|⑬|⑭|⑮)/g, '\n\n$1')
    .replace(/\s+(\d+\.)\s/g, '\n$1 ')
    .trim();

  // Format numbered items
  const lines = result.split('\n');
  const formatted: string[] = [];
  for (const line of lines) {
    const t = line.trim();
    if (!t) { formatted.push(''); continue; }
    formatted.push(t);
  }
  return formatted.join('\n');
}

/**
 * Enhanced post-processing for HWP → Markdown conversion
 * Handles legal document layout tables, chapter/article headings,
 * and general HWP formatting artifacts
 */
function postProcessHwpMarkdown(md: string): string {
  // Phase 1: Clean up common HWP artifacts
  let result = md
    .replace(/^-\s*\d+\s*-\s*$/gm, '')        // Page numbers like "- 7 -"
    .replace(/^#{1,6}\s*$/gm, '')               // Empty headings
    .replace(/\[autonumbering[^\]]*\]/g, '')    // Auto-numbering markers
    .replace(/[ \t]+$/gm, '');                  // Trailing whitespace

  // Phase 2: Detect and convert legal layout tables
  const lines = result.split('\n');
  const segments: { type: 'text' | 'table'; lines: string[] }[] = [];
  let currentSegment: { type: 'text' | 'table'; lines: string[] } = { type: 'text', lines: [] };

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    const isTableRow = /^\|.*\|$/.test(line);

    if (isTableRow || isTableSeparator(line)) {
      if (currentSegment.type !== 'table') {
        if (currentSegment.lines.length > 0) segments.push(currentSegment);
        currentSegment = { type: 'table', lines: [] };
      }
      if (!isTableSeparator(line)) {
        currentSegment.lines.push(line);
      }
    } else {
      if (currentSegment.type !== 'text') {
        if (currentSegment.lines.length > 0) segments.push(currentSegment);
        currentSegment = { type: 'text', lines: [] };
      }
      currentSegment.lines.push(lines[i]);
    }
  }
  if (currentSegment.lines.length > 0) segments.push(currentSegment);

  // Phase 3: Process each segment
  const output: string[] = [];

  for (const seg of segments) {
    if (seg.type === 'table') {
      const parsedRows = seg.lines.map(parseTableRow);

      if (isLegalLayoutTable(parsedRows)) {
        // Convert legal layout table to structured markdown
        output.push(convertLegalTable(parsedRows));
      } else {
        // Keep non-legal tables as-is (data tables, form tables, etc.)
        // But clean them up
        for (const line of seg.lines) {
          output.push(line);
        }
      }
    } else {
      // Process text lines
      for (const line of seg.lines) {
        const trimmed = line.trim();
        if (!trimmed) { output.push(''); continue; }

        // Chapter headings
        if (/^제\d+장\s+/.test(trimmed) && trimmed.length < 40) {
          output.push('', `## ${trimmed}`, '');
          continue;
        }
        // Section headings
        if (/^제\d+절\s+/.test(trimmed) && trimmed.length < 40) {
          output.push('', `### ${trimmed}`, '');
          continue;
        }
        // Article headings
        if (/^제\d+조(?:의\d+)?\(/.test(trimmed) && trimmed.length < 40) {
          output.push('', `#### ${trimmed}`, '');
          continue;
        }
        // Roman numeral headings
        if (/^[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ][\.\s]/.test(trimmed) && !trimmed.includes('|')) {
          output.push('', `# ${trimmed}`, '');
          continue;
        }

        output.push(line);
      }
    }
  }

  return output.join('\n')
    .replace(/\n{4,}/g, '\n\n\n')
    .trim() + '\n';
}

/**
 * Comprehensive HWP text → Markdown converter
 * Handles legal documents, general documents, and mixed content
 */
function formatHwpTextToMarkdown(text: string): string {
  // Phase 1: Clean up raw text
  let cleaned = text
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/^-\s*\d+\s*-\s*$/gm, '')              // Page numbers "- 7 -"
    .replace(/\[autonumbering[^\]]*\]/g, '')          // Auto-numbering
    .replace(/[ \t]+$/gm, '')                         // Trailing whitespace
    .replace(/^[ \t]{20,}/gm, '')                     // Over-indented lines (layout artifacts)
    .replace(/(\S)\s{4,}(\S)/g, '$1\n$2');            // Split text joined by large whitespace gaps

  // Fix spaced-out titles: "표 준 취 업 규 칙" → "표준취업규칙"
  cleaned = cleaned.replace(/^([가-힣])\s([가-힣])\s([가-힣])\s([가-힣])\s([가-힣])\s([가-힣])\s([가-힣])$/gm,
    '$1$2$3$4$5$6$7');
  cleaned = cleaned.replace(/^([가-힣])\s([가-힣])\s([가-힣])\s([가-힣])\s([가-힣])\s([가-힣])$/gm,
    '$1$2$3$4$5$6');
  cleaned = cleaned.replace(/^([가-힣])\s([가-힣])\s([가-힣])\s([가-힣])\s([가-힣])$/gm,
    '$1$2$3$4$5');
  // Generic: lines of single Korean chars separated by spaces (3+ chars)
  cleaned = cleaned.replace(/^(([가-힣])\s){2,}([가-힣])$/gm, (match) =>
    match.replace(/\s/g, ''));

  const lines = cleaned.split('\n');
  const result: string[] = [];
  let inAnnotation = false;    // Whether we're in a 착안사항/참고 block
  let lastWasHeading = false;
  let documentTitle = '';

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const trimmed = line.trim();

    // Skip empty lines (but preserve paragraph breaks)
    if (!trimmed) {
      if (inAnnotation) {
        inAnnotation = false;
      }
      if (result.length > 0 && result[result.length - 1] !== '') {
        result.push('');
      }
      lastWasHeading = false;
      continue;
    }

    // Skip pure page layout artifacts
    if (/^(조\s*문\s*순\s*서|취업규칙\s*\(안\)|작성시\s*착안사항)$/.test(trimmed)) continue;
    if (/^(일반\s*근로자용|고용노동부)$/.test(trimmed)) {
      if (!documentTitle) continue;
    }

    // Document title detection (first substantial text)
    if (!documentTitle && /^[가-힣]{2,}$/.test(trimmed) && trimmed.length >= 3 && trimmed.length <= 20) {
      if (/취업규칙|근로계약|인사규정|복무규정|보수규정|급여규정/.test(trimmed)) {
        documentTitle = trimmed;
        result.push(`# ${trimmed}`);
        result.push('');
        lastWasHeading = true;
        continue;
      }
    }

    // Roman numeral top-level headings
    if (/^[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ][\.\s]/.test(trimmed)) {
      result.push('', `# ${trimmed}`, '');
      lastWasHeading = true;
      continue;
    }

    // Chapter headings: 제1장 총 칙
    if (/^제\s*\d+\s*장\s+/.test(trimmed) && trimmed.length < 50) {
      const chapTitle = trimmed.replace(/\s+/g, ' ').trim();
      result.push('', `## ${chapTitle}`, '');
      lastWasHeading = true;
      continue;
    }

    // Section headings: 제1절 인사위원회
    if (/^제\s*\d+\s*절\s+/.test(trimmed) && trimmed.length < 50) {
      const secTitle = trimmed.replace(/\s+/g, ' ').trim();
      result.push('', `### ${secTitle}`, '');
      lastWasHeading = true;
      continue;
    }

    // TOC detection: line with many 제X조 references (moved up to prevent #### conversion)
    // TOC pattern: "제1조(목적) [필수] 7" (article + tag + PAGE NUMBER)
    // Annotation pattern: "[필수] 취업규칙을..." (tag + DESCRIPTION)
    // Key difference: TOC has "[필수/선택] \d+" (tag+page), annotations have "[필수/선택] text"
    const tocArticleRefs = trimmed.match(/제\d+조/g);
    const isTocPattern = tocArticleRefs && tocArticleRefs.length > 2
      && (trimmed.match(/\[필수[^\]]*\]\s*\d+|제\d+조\([^)]+\)\s*\d+/g) || []).length > 2;
    if (!inAnnotation && isTocPattern && !/착안사항|☞|◈/.test(trimmed)) {
      // Split concatenated TOC entries: "제1장 총칙제1조(목적) [필수] 7제2조..."
      // Insert newlines before each 제X장, 제X절, 제X조 pattern
      // Strip table separators first, then split on article patterns
      const tocClean = trimmed.replace(/\|/g, ' ').replace(/---/g, '').replace(/\s+/g, ' ');
      const tocFormatted = tocClean
        .replace(/(제\s*\d+\s*장\s*[^\d제]*)/g, '\n$1')
        .replace(/(제\s*\d+\s*절\s*[^\d제]*)/g, '\n$1')
        .replace(/(제\s*\d+\s*조(?:의\s*\d+)?\s*\([^)]+\)\s*\[(?:필수|선택)[^\]]*\]\s*\d*)/g, '\n$1')
        .replace(/(제\s*\d+\s*조(?:의\s*\d+)?\s*\([^)]+\)\s*(?!\[))/g, '\n$1')
        .replace(/(부\s*칙\s*\d*)/g, '\n$1')
        .trim();
      const tocLines = tocFormatted.split('\n').filter(l => l.trim());
      for (const tl of tocLines) {
        const p = tl.trim();
        if (/^제\s*\d+\s*장/.test(p)) result.push(`\n**${p}**`);
        else if (/^제\s*\d+\s*절/.test(p)) result.push(`  *${p}*`);
        else if (/^제\s*\d+\s*조/.test(p)) result.push(`  - ${p}`);
        else if (/^부\s*칙/.test(p)) result.push(`\n**${p}**`);
        else result.push(`  ${p}`);
      }
      continue;
    }

    // Article headings: 제1조(목적) ...body
    const articleMatch = trimmed.match(/^(제\s*\d+\s*조(?:의\s*\d+)?\s*\([^)]+\))\s*([\s\S]*)/);
    if (articleMatch) {
      const title = articleMatch[1].replace(/\s+/g, ' ').trim();
      const body = (articleMatch[2] || '').trim();
      result.push('', `#### ${title}`, '');
      if (body) {
        result.push(formatArticleBody(body));
      }
      lastWasHeading = true;
      continue;
    }

    // Standalone article reference without parenthetical title
    if (/^제\s*\d+\s*조\s/.test(trimmed) && trimmed.length < 50 && !/[①②③④⑤]/.test(trimmed)) {
      result.push('', `#### ${trimmed.replace(/\s+/g, ' ').trim()}`, '');
      lastWasHeading = true;
      continue;
    }

    // Continuation of annotation block (check BEFORE new annotation start)
    if (inAnnotation) {
      const breaksAnnotation = /^제\s*\d+\s*(조|장|절)/.test(trimmed)
        || /^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮]/.test(trimmed)
        || /^[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]/.test(trimmed)
        || /^부\s*칙/.test(trimmed)
        || /^\[?별지/.test(trimmed)
        || /^\|/.test(trimmed);

      // New major annotation markers start a fresh blockquote header
      const isNewAnnotation = /^◈/.test(trimmed)
        || /^\[필수\]/.test(trimmed) || /^\[선택\]/.test(trimmed)
        || /^\[필수,\s*선택\]/.test(trimmed) || /^\[선택,\s*필수\]/.test(trimmed)
        || /^착안사항/.test(trimmed) || /^※\s/.test(trimmed)
        || (/\[필수\]|\[선택\]|\[필수,\s*선택\]|\[선택,\s*필수\]/.test(trimmed) && !/^제\s*\d+/.test(trimmed));

      if (breaksAnnotation) {
        inAnnotation = false;
        // Fall through to process as normal element
      } else if (isNewAnnotation) {
        result.push(`> **착안사항**: ${trimmed}`);
        continue;
      } else {
        // ☞, *, -, · and other text continue in blockquote
        result.push(`> ${trimmed}`);
        continue;
      }
    }

    // 착안사항 / annotation markers - start new annotation block
    const isAnnotationStart = /^◈/.test(trimmed)
      || /^\[필수\]/.test(trimmed) || /^\[선택\]/.test(trimmed)
      || /^\[필수,\s*선택\]/.test(trimmed) || /^\[선택,\s*필수\]/.test(trimmed)
      || /^☞/.test(trimmed) || /^착안사항/.test(trimmed)
      || /^※\s/.test(trimmed) || /^\(참고\)/.test(trimmed)
      || (/\[필수\]|\[선택\]|\[필수,\s*선택\]|\[선택,\s*필수\]/.test(trimmed) && !/^제\s*\d+/.test(trimmed));
    if (isAnnotationStart) {
      result.push(`> **착안사항**: ${trimmed}`);
      inAnnotation = true;
      continue;
    }

    // "부 칙" or appendix
    if (/^부\s*칙/.test(trimmed)) {
      result.push('', `## 부칙`, '');
      lastWasHeading = true;
      continue;
    }

    // 별지/별첨 (appendix forms)
    if (/^\[?별지\s*\d+\]?/.test(trimmed) || /^\[?별첨\]?/.test(trimmed)) {
      result.push('', `### ${trimmed.replace(/[\[\]]/g, '')}`, '');
      lastWasHeading = true;
      continue;
    }

    // Clause markers ① ② etc. at start of line
    if (/^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮]/.test(trimmed)) {
      result.push('');
      result.push(trimmed);
      inAnnotation = false;
      continue;
    }

    // Numbered list items: 1. 2. etc.
    if (/^\d+\.\s/.test(trimmed)) {
      result.push(trimmed);
      continue;
    }

    // Bullet-like markers
    if (/^[○●■□▶▷◆◇]\s/.test(trimmed)) {
      result.push(`- ${trimmed.slice(2)}`);
      continue;
    }

    // Korean letter list items: 가. 나. 다.
    if (/^[가나다라마바사아자차카타파하]\.\s/.test(trimmed)) {
      result.push(`  - ${trimmed}`);
      continue;
    }

    // Date line (2026. 2.)
    if (/^\d{4}\.\s*\d{1,2}\.\s*$/.test(trimmed)) {
      result.push(`**${trimmed}**`);
      result.push('');
      continue;
    }

    // (TOC detection moved above article heading matching)

    // Default: regular text
    inAnnotation = false;
    result.push(trimmed);
  }

  const joined = result.join('\n')
    .replace(/\n{4,}/g, '\n\n\n')
    .trim() + '\n';

  return finalCleanup(joined);
}

/**
 * Format article body text (after the article title)
 * Handles clause markers, sub-items, and paragraph structure
 */
function formatArticleBody(body: string): string {
  // Split embedded articles: "...따른다.제3조(정의) ..." → separate heading
  // Only when preceded by sentence-ending period (NOT law references like "근로기준법 제17조")
  let result = body.replace(/([다함음임됨짐음])\.(\s*)(제\s*\d+\s*조(?:의\s*\d+)?\s*\([^)]+\))/g,
    '$1.\n\n#### $3\n');

  // Split on clause markers (①②③...)
  result = result
    .replace(/(①|②|③|④|⑤|⑥|⑦|⑧|⑨|⑩|⑪|⑫|⑬|⑭|⑮)/g, '\n\n$1')
    .replace(/\s(\d+\.)\s/g, '\n$1 ')
    .trim();

  // Korean letter list items (가. 나. 다.) — ONLY at start of segment after newline
  // NOT mid-sentence like "하여야 한다." or "있다."
  result = result.replace(/\n([가나다라마바사아자차카타파하]\.)\s/g, '\n  $1 ');

  const lines = result.split('\n');
  const formatted: string[] = [];
  for (const line of lines) {
    const t = line.trim();
    if (!t) continue;
    formatted.push(t);
  }
  return formatted.join('\n');
}

/**
 * Post-process final markdown to fix remaining issues
 */
function finalCleanup(md: string): string {
  let result = md;

  // 1. Split ☞ in annotations onto new lines (within blockquotes)
  result = result.replace(/(\s)(☞\s*\(참고\))/g, '\n> $2');

  // 2. Remove empty table rows: lines with only | and whitespace (but not annotation lines)
  result = result.replace(/^(?:\|\s*)+\|\s*$/gm, (match) => {
    // Keep if it's part of a table structure (has --- separator nearby)
    return match.includes('착안사항') ? match : '';
  });

  // 3. Remove orphan table separators with no data rows around them
  result = result.replace(/^(?:\|\s*---\s*)+\|\s*$/gm, (match, offset) => {
    // Check if there's actual table content nearby (within 200 chars before/after)
    const before = result.slice(Math.max(0, offset - 200), offset);
    const after = result.slice(offset + match.length, offset + match.length + 200);
    const hasContentBefore = /\|[^|\s\-][^|]*\|/.test(before);
    const hasContentAfter = /\|[^|\s\-][^|]*\|/.test(after);
    return (hasContentBefore || hasContentAfter) ? match : '';
  });

  // 4. Clean up consecutive empty lines
  result = result.replace(/\n{4,}/g, '\n\n\n');

  // 5. Fix annotation lines that have table fragments mixed in
  result = result.replace(/^>\s*\*\*착안사항\*\*:\s*\|\s*/gm, '> **착안사항**: ');

  // 6. Clean inline table fragments in blockquotes: " | ... | | --- |" at end of annotation lines
  result = result.replace(/(\s)\|\s*\|\s*---\s*\|\s*$/gm, '$1');
  result = result.replace(/\s*\|\s*\|\s*---\s*\|\s*$/gm, '');

  // 7. Remove standalone "| 취업규칙(안) | 취업규칙(안) |" header rows (TOC decorations)
  result = result.replace(/^\|\s*취업규칙\(안\)\s*\|\s*취업규칙\(안\)\s*\|\s*$/gm, '');

  // 8. Remove trailing " |" from TOC entries (residual table cell separators)
  result = result.replace(/^(\s+-\s+제\d+조.+?)\s*\|\s*$/gm, '$1');
  result = result.replace(/^(\*\*제\d+장.+?)\s*\|\s*\*\*\s*$/gm, '$1**');
  result = result.replace(/^(\*\*제\d+장\s+.+?)\s*\|\*\*\s*$/gm, '$1**');

  // 9. Remove orphan "| --- | --- |" lines not adjacent to table content
  result = result.replace(/^\|\s*---\s*\|\s*---\s*\|\s*$/gm, '');

  return result.trim() + '\n';
}

export async function OPTIONS() {
  return new NextResponse(null, { status: 204, headers: CORS_HEADERS });
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function jsonWithCors(data: any, status = 200) {
  return NextResponse.json(data, { status, headers: CORS_HEADERS });
}

export async function POST(request: NextRequest) {
  // === PROXY MODE (Vercel → Render) ===
  if (RENDER_API_URL) {
    try {
      const formData = await request.formData();
      const file = formData.get('file') as File | null;
      if (!file) return jsonWithCors({ error: '파일을 선택해 주세요.' }, 400);

      const renderFormData = new FormData();
      renderFormData.append('file', file);

      const res = await fetch(`${RENDER_API_URL}/api/convert`, {
        method: 'POST',
        body: renderFormData,
      });
      const data = await res.json();
      return jsonWithCors(data, res.status);
    } catch (error: unknown) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return jsonWithCors({ error: `서버 연결 오류: ${message}` }, 502);
    }
  }

  // === DIRECT MODE (Render / Mac mini) ===
  let tempPath = '';
  let tempHtmlPath = '';

  try {
    const formData = await request.formData();
    const file = formData.get('file') as File | null;
    if (!file) return jsonWithCors({ error: '파일을 선택해 주세요.' }, 400);

    const ext = getExtension(file.name);
    if (!SUPPORTED_EXTENSIONS.includes(ext)) {
      return jsonWithCors({ error: `지원하지 않는 파일 형식입니다: ${ext}` }, 400);
    }

    const tempDir = join(tmpdir(), 'md-converter');
    await mkdir(tempDir, { recursive: true });
    const tempId = randomUUID();
    tempPath = join(tempDir, `${tempId}${ext}`);

    const bytes = await file.arrayBuffer();
    await writeFile(tempPath, Buffer.from(bytes));
    const fileSizeMB = (file.size / (1024 * 1024)).toFixed(2);

    if (TEXT_EXTENSIONS.includes(ext)) {
      const textContent = await readFile(tempPath, 'utf-8');
      return jsonWithCors({ markdown: textContent, filename: file.name.replace(/\.[^.]+$/, '.md'), originalName: file.name, fileSize: `${fileSizeMB} MB`, lineCount: textContent.split('\n').length, charCount: textContent.length });
    }

    if (HWP_EXTENSIONS.includes(ext)) {
      const hwpPaths = [
        join(HOME, '.local', 'bin'),
        '/usr/local/bin',
        '/opt/homebrew/bin',
      ];
      let hwp5htmlBin = '';
      let hwp5txtBin = '';
      for (const dir of hwpPaths) {
        try {
          await execFileAsync('test', ['-f', join(dir, 'hwp5html')]);
          hwp5htmlBin = join(dir, 'hwp5html');
          hwp5txtBin = join(dir, 'hwp5txt');
          break;
        } catch { continue; }
      }
      tempHtmlPath = join(tempDir, `${tempId}.html`);
      const hwpToolAvailable = !!hwp5htmlBin;
      if (!hwpToolAvailable) {
        return proxyToRender(file, ext);
      }

      // Strategy: hwp5html → custom HTML text extraction → formatHwpTextToMarkdown
      let conversionMethod = '';
      let htmlSize = 0;
      let plainTextSize = 0;

      // Method 1: hwp5html → extractTextFromHwpHtml → formatHwpTextToMarkdown
      try {
        await execFileAsync(hwp5htmlBin, ['--html', tempPath, '--output', tempHtmlPath], { timeout: 120000, maxBuffer: 100 * 1024 * 1024, env: ENV });
        const htmlContent = await readFile(tempHtmlPath, 'utf-8');
        htmlSize = htmlContent.length;
        const plainText = extractTextFromHwpHtml(htmlContent);
        plainTextSize = plainText.length;
        if (plainText.trim().length > 500) {
          conversionMethod = 'hwp5html→extractText';
          const markdown = formatHwpTextToMarkdown(plainText);
          return jsonWithCors({ markdown, filename: file.name.replace(/\.[^.]+$/, '.md'), originalName: file.name, fileSize: `${fileSizeMB} MB`, lineCount: markdown.split('\n').length, charCount: markdown.length, _debug: { method: conversionMethod, htmlSize, plainTextSize } });
        }
        conversionMethod = 'hwp5html→extractText (too short, trying next)';
      } catch (e: unknown) {
        conversionMethod = `hwp5html failed: ${e instanceof Error ? e.message : String(e)}`;
      }

      // Method 2: hwp5html → markitdown → convertLayoutTablesToText → formatHwpTextToMarkdown
      try {
        if (!htmlSize) {
          await execFileAsync(hwp5htmlBin, ['--html', tempPath, '--output', tempHtmlPath], { timeout: 120000, maxBuffer: 100 * 1024 * 1024, env: ENV });
        }
        const markitdownBin = await checkMarkitdown();
        if (markitdownBin) {
          const { stdout } = await execFileAsync(markitdownBin, [tempHtmlPath], { timeout: 120000, maxBuffer: 100 * 1024 * 1024, env: ENV });
          if (stdout.trim().length > 500) {
            conversionMethod = 'hwp5html→markitdown→tableToText';
            // Convert markdown tables to plain text, then apply HWP formatter
            const textWithoutTables = convertLayoutTablesToText(stdout);
            const markdown = formatHwpTextToMarkdown(textWithoutTables);
            return jsonWithCors({ markdown, filename: file.name.replace(/\.[^.]+$/, '.md'), originalName: file.name, fileSize: `${fileSizeMB} MB`, lineCount: markdown.split('\n').length, charCount: markdown.length, _debug: { method: conversionMethod, htmlSize, markitdownSize: stdout.length, textSize: textWithoutTables.length } });
          }
        }
      } catch { /* markitdown failed */ }

      // Method 3: hwp5txt → formatHwpTextToMarkdown
      try {
        const { stdout: hwpText } = await execFileAsync(hwp5txtBin, [tempPath], { timeout: 60000, maxBuffer: 50 * 1024 * 1024, env: ENV });
        if (hwpText && hwpText.trim().length > 200) {
          conversionMethod = 'hwp5txt';
          const markdown = formatHwpTextToMarkdown(hwpText);
          return jsonWithCors({ markdown, filename: file.name.replace(/\.[^.]+$/, '.md'), originalName: file.name, fileSize: `${fileSizeMB} MB`, lineCount: markdown.split('\n').length, charCount: markdown.length, _debug: { method: conversionMethod, hwp5txtSize: hwpText.length, htmlSize, plainTextSize } });
        }
      } catch { /* hwp5txt also failed */ }

      return jsonWithCors({ error: 'HWP 변환 오류: 모든 변환 방법이 실패했습니다.', _debug: { conversionMethod, htmlSize, plainTextSize } }, 500);
    }

    const markitdownBin = await checkMarkitdown();
    // Skip markitdown for PDF and HTML — use our custom processors for better structure
    if (markitdownBin && !PDF_EXTENSIONS.includes(ext) && !HTML_EXTENSIONS.includes(ext)) {
      try {
        const { stdout, stderr } = await execFileAsync(markitdownBin, [tempPath], { timeout: 120000, maxBuffer: 50 * 1024 * 1024, env: ENV });
        if (stdout || !stderr) {
          return jsonWithCors({ markdown: stdout, filename: file.name.replace(/\.[^.]+$/, '.md'), originalName: file.name, fileSize: `${fileSizeMB} MB`, lineCount: stdout.split('\n').length, charCount: stdout.length });
        }
      } catch { /* fallback */ }
    }

    let markdown = '';
    let _pdfDebug = '';
    if (PDF_EXTENSIONS.includes(ext)) {
      markdown = await convertPdfToMarkdown(tempPath);
      const pdfHeadings = markdown.split('\n').filter(l => l.startsWith('#')).length;
      _pdfDebug = `path=convertPdfToMarkdown,headings=${pdfHeadings},lines=${markdown.split('\n').length}`;
      // PDF 로컬 변환 실패 시 Render 백엔드로 프록시
      if (!markdown || markdown.trim().length === 0) {
        return proxyToRender(file, ext);
      }
    } else if (HTML_EXTENSIONS.includes(ext)) {
      const htmlContent = await readFile(tempPath, 'utf-8');
      markdown = await convertHtmlToMarkdown(htmlContent);
    } else if (ext === '.csv') {
      markdown = convertCsvToMarkdown(await readFile(tempPath, 'utf-8'));
    } else if (ext === '.json') {
      const jsonContent = await readFile(tempPath, 'utf-8');
      try { markdown = '```json\n' + JSON.stringify(JSON.parse(jsonContent), null, 2) + '\n```\n'; }
      catch { markdown = '```\n' + jsonContent + '\n```\n'; }
    } else if (OFFICE_EXTENSIONS.includes(ext)) {
      markdown = await convertWithOfficeParser(tempPath);
    } else {
      // Images, audio, zip etc. - proxy to Render backend
      return proxyToRender(file, ext);
    }

    // Final post-processing: add headings for PDF and HTML regardless of code path
    if (PDF_EXTENSIONS.includes(ext)) {
      markdown = postProcessPdfMarkdown(markdown);
    } else if (HTML_EXTENSIONS.includes(ext)) {
      markdown = postProcessHtmlMarkdown(markdown);
    }

    return jsonWithCors({ markdown, filename: file.name.replace(/\.[^.]+$/, '.md'), originalName: file.name, fileSize: `${fileSizeMB} MB`, lineCount: markdown.split('\n').length, charCount: markdown.length });
  } catch (error: unknown) {
    const message = error instanceof Error ? error.message : 'Unknown error';
    if (message.includes('timeout')) return jsonWithCors({ error: '변환 시간이 초과되었습니다.' }, 504);
    return jsonWithCors({ error: `변환 오류: ${message}` }, 500);
  } finally {
    for (const p of [tempPath, tempHtmlPath]) {
      if (p) { try { await unlink(p); } catch { /* */ } }
    }
  }
}

export const config = { api: { bodyParser: false } };


