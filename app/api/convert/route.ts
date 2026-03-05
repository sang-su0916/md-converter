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
async function proxyToRender(file: File): Promise<Response> {
  const formData = new FormData();
  formData.append('file', file);

  const res = await fetch(`${RENDER_BACKEND}/api/convert`, {
    method: 'POST',
    body: formData,
    signal: AbortSignal.timeout(180000), // 3min timeout for cold start
  });

  const data = await res.json();
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
  const { extractText, getDocumentProxy } = await import('unpdf');
  const buffer = await readFile(filePath);
  const pdf = await getDocumentProxy(new Uint8Array(buffer));
  const { text } = await extractText(pdf, { mergePages: true });
  return text || '';
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
  return turndown.turndown(htmlContent);
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

    // 착안사항 / annotation markers
    if (/^◈\s/.test(trimmed) || /^\[필수\]/.test(trimmed) || /^\[선택\]/.test(trimmed)
        || /^☞\s*\(참고\)/.test(trimmed) || /^착안사항/.test(trimmed)) {
      result.push(`> ${trimmed}`);
      inAnnotation = true;
      continue;
    }

    // Continuation of annotation (indented or starts with specific patterns)
    if (inAnnotation && (/^\*\s/.test(trimmed) || /^☞/.test(trimmed) || /^-\s/.test(trimmed)
        || /^다만/.test(trimmed) || /^단,/.test(trimmed))) {
      result.push(`> ${trimmed}`);
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

    // TOC detection: line with many 제X조 references
    const articleRefs = trimmed.match(/제\d+조/g);
    if (articleRefs && articleRefs.length > 2) {
      // Format as structured list
      const parts = trimmed.split(/\s{2,}/);
      for (const part of parts) {
        const p = part.trim();
        if (!p) continue;
        if (/^제\d+장/.test(p)) result.push(`\n**${p}**`);
        else if (/^제\d+절/.test(p)) result.push(`  *${p}*`);
        else if (/^제\d+조/.test(p)) result.push(`  - ${p}`);
        else result.push(`  ${p}`);
      }
      continue;
    }

    // Default: regular text
    inAnnotation = false;
    result.push(trimmed);
  }

  return result.join('\n')
    .replace(/\n{4,}/g, '\n\n\n')
    .trim() + '\n';
}

/**
 * Format article body text (after the article title)
 * Handles clause markers, sub-items, and paragraph structure
 */
function formatArticleBody(body: string): string {
  // Split on clause markers (①②③...)
  let result = body
    .replace(/(①|②|③|④|⑤|⑥|⑦|⑧|⑨|⑩|⑪|⑫|⑬|⑭|⑮)/g, '\n\n$1')
    .replace(/\s{2,}(\d+\.)\s/g, '\n$1 ')
    .trim();

  const lines = result.split('\n');
  const formatted: string[] = [];
  for (const line of lines) {
    const t = line.trim();
    if (!t) continue;
    formatted.push(t);
  }
  return formatted.join('\n');
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
        // Proxy to Render backend (has hwp5html/hwp5txt in Docker)
        return proxyToRender(file);
      }

      // Strategy: hwp5txt FIRST (produces cleaner, table-free text)
      // then hwp5html + markitdown as fallback
      try {
        const { stdout: hwpText } = await execFileAsync(hwp5txtBin, [tempPath], { timeout: 60000, maxBuffer: 50 * 1024 * 1024, env: ENV });
        if (hwpText && hwpText.trim().length > 100) {
          const markdown = formatHwpTextToMarkdown(hwpText);
          return jsonWithCors({ markdown, filename: file.name.replace(/\.[^.]+$/, '.md'), originalName: file.name, fileSize: `${fileSizeMB} MB`, lineCount: markdown.split('\n').length, charCount: markdown.length });
        }
      } catch { /* hwp5txt failed, try hwp5html */ }

      // Fallback: hwp5html → markitdown/turndown → postProcess
      try {
        await execFileAsync(hwp5htmlBin, ['--html', tempPath, '--output', tempHtmlPath], { timeout: 120000, maxBuffer: 100 * 1024 * 1024, env: ENV });
        const markitdownBin = await checkMarkitdown();
        let markdown: string;
        if (markitdownBin) {
          const { stdout } = await execFileAsync(markitdownBin, [tempHtmlPath], { timeout: 120000, maxBuffer: 100 * 1024 * 1024, env: ENV });
          markdown = postProcessHwpMarkdown(stdout);
        } else {
          const htmlContent = await readFile(tempHtmlPath, 'utf-8');
          markdown = postProcessHwpMarkdown(await convertHtmlToMarkdown(htmlContent));
        }
        return jsonWithCors({ markdown, filename: file.name.replace(/\.[^.]+$/, '.md'), originalName: file.name, fileSize: `${fileSizeMB} MB`, lineCount: markdown.split('\n').length, charCount: markdown.length });
      } catch (hwpError: unknown) {
        const msg = hwpError instanceof Error ? hwpError.message : 'Unknown error';
        return jsonWithCors({ error: `HWP 변환 오류: ${msg}` }, 500);
      }
    }

    const markitdownBin = await checkMarkitdown();
    if (markitdownBin) {
      try {
        const { stdout, stderr } = await execFileAsync(markitdownBin, [tempPath], { timeout: 120000, maxBuffer: 50 * 1024 * 1024, env: ENV });
        if (stdout || !stderr) {
          return jsonWithCors({ markdown: stdout, filename: file.name.replace(/\.[^.]+$/, '.md'), originalName: file.name, fileSize: `${fileSizeMB} MB`, lineCount: stdout.split('\n').length, charCount: stdout.length });
        }
      } catch { /* fallback */ }
    }

    let markdown = '';
    if (PDF_EXTENSIONS.includes(ext)) {
      markdown = await convertPdfToMarkdown(tempPath);
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
      return proxyToRender(file);
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
