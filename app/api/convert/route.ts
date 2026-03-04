import { NextRequest, NextResponse } from 'next/server';
import { writeFile, readFile, unlink, mkdir } from 'fs/promises';
import { join } from 'path';
import { execFile } from 'child_process';
import { promisify } from 'util';
import { tmpdir } from 'os';
import { randomUUID } from 'crypto';

const execFileAsync = promisify(execFile);

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

function postProcessHwpMarkdown(md: string): string {
  let result = md.replace(/^-\s*\d+\s*-\s*$/gm, '').replace(/^#{1,6}\s*$/gm, '').replace(/\n{4,}/g, '\n\n\n').replace(/\[autonumbering[^\]]*\]/g, '').replace(/[ \t]+$/gm, '');
  const lines = result.split('\n'); const processed: string[] = [];
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) { processed.push(''); continue; }
    if (/^제\d+장\s+/.test(trimmed) && !trimmed.includes('|') && trimmed.length < 30) { processed.push(`## ${trimmed}`); continue; }
    if (/^제\d+조[\s(（]/.test(trimmed) && !trimmed.includes('|') && trimmed.length < 30) { processed.push(`### ${trimmed}`); continue; }
    if (/^[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ][\.\s]/.test(trimmed) && !trimmed.includes('|')) { processed.push(`# ${trimmed}`); continue; }
    processed.push(line);
  }
  return processed.join('\n').trim() + '\n';
}

function formatHwpTextToMarkdown(text: string): string {
  const lines = text.split('\n'); const result: string[] = [];
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) { result.push(''); continue; }
    if (/^[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ][\.\s]/.test(trimmed)) result.push(`# ${trimmed}`);
    else if (/^제\s*\d+\s*장/.test(trimmed)) result.push(`## ${trimmed}`);
    else if (/^제\s*\d+\s*조/.test(trimmed)) result.push(`### ${trimmed}`);
    else if (/^[①②③④⑤⑥⑦⑧⑨⑩]/.test(trimmed)) result.push(`- ${trimmed}`);
    else if (/^[○●■□▶▷◆◇]/.test(trimmed)) result.push(`- ${trimmed}`);
    else if (/^[가나다라마바사아자차카타파하][\.\)]\s/.test(trimmed)) result.push(`  - ${trimmed}`);
    else if (/^\d+[\.\)]\s/.test(trimmed)) result.push(`${trimmed}`);
    else result.push(trimmed);
  }
  return result.join('\n');
}

export async function POST(request: NextRequest) {
  // === PROXY MODE (Vercel → Render) ===
  if (RENDER_API_URL) {
    try {
      const formData = await request.formData();
      const file = formData.get('file') as File | null;
      if (!file) return NextResponse.json({ error: '파일을 선택해 주세요.' }, { status: 400 });

      const renderFormData = new FormData();
      renderFormData.append('file', file);

      const res = await fetch(`${RENDER_API_URL}/api/convert`, {
        method: 'POST',
        body: renderFormData,
      });
      const data = await res.json();
      return NextResponse.json(data, { status: res.status });
    } catch (error: unknown) {
      const message = error instanceof Error ? error.message : 'Unknown error';
      return NextResponse.json({ error: `서버 연결 오류: ${message}` }, { status: 502 });
    }
  }

  // === DIRECT MODE (Render / Mac mini) ===
  let tempPath = '';
  let tempHtmlPath = '';

  try {
    const formData = await request.formData();
    const file = formData.get('file') as File | null;
    if (!file) return NextResponse.json({ error: '파일을 선택해 주세요.' }, { status: 400 });

    const ext = getExtension(file.name);
    if (!SUPPORTED_EXTENSIONS.includes(ext)) {
      return NextResponse.json({ error: `지원하지 않는 파일 형식입니다: ${ext}` }, { status: 400 });
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
      return NextResponse.json({ markdown: textContent, filename: file.name.replace(/\.[^.]+$/, '.md'), originalName: file.name, fileSize: `${fileSizeMB} MB`, lineCount: textContent.split('\n').length, charCount: textContent.length });
    }

    if (HWP_EXTENSIONS.includes(ext)) {
      const hwp5htmlBin = join(HOME, '.local', 'bin', 'hwp5html');
      const hwp5txtBin = join(HOME, '.local', 'bin', 'hwp5txt');
      tempHtmlPath = join(tempDir, `${tempId}.html`);
      let hwpToolAvailable = false;
      try { await execFileAsync('test', ['-f', hwp5htmlBin]); hwpToolAvailable = true; } catch { /* */ }
      if (!hwpToolAvailable) {
        return NextResponse.json({ error: `HWP 파일은 Render 서버에서만 변환 가능합니다.\n\nRender: https://md-converter-ghdf.onrender.com` }, { status: 400 });
      }
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
        return NextResponse.json({ markdown, filename: file.name.replace(/\.[^.]+$/, '.md'), originalName: file.name, fileSize: `${fileSizeMB} MB`, lineCount: markdown.split('\n').length, charCount: markdown.length });
      } catch (hwpError: unknown) {
        try {
          const { stdout: hwpText } = await execFileAsync(hwp5txtBin, [tempPath], { timeout: 60000, maxBuffer: 50 * 1024 * 1024, env: ENV });
          const markdown = formatHwpTextToMarkdown(hwpText);
          return NextResponse.json({ markdown, filename: file.name.replace(/\.[^.]+$/, '.md'), originalName: file.name, fileSize: `${fileSizeMB} MB`, lineCount: markdown.split('\n').length, charCount: markdown.length });
        } catch {
          const msg = hwpError instanceof Error ? hwpError.message : 'Unknown error';
          return NextResponse.json({ error: `HWP 변환 오류: ${msg}` }, { status: 500 });
        }
      }
    }

    const markitdownBin = await checkMarkitdown();
    if (markitdownBin) {
      try {
        const { stdout, stderr } = await execFileAsync(markitdownBin, [tempPath], { timeout: 120000, maxBuffer: 50 * 1024 * 1024, env: ENV });
        if (stdout || !stderr) {
          return NextResponse.json({ markdown: stdout, filename: file.name.replace(/\.[^.]+$/, '.md'), originalName: file.name, fileSize: `${fileSizeMB} MB`, lineCount: stdout.split('\n').length, charCount: stdout.length });
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
      return NextResponse.json({ error: `이 파일 형식(${ext})은 Render 서버에서만 변환 가능합니다.` }, { status: 400 });
    }

    return NextResponse.json({ markdown, filename: file.name.replace(/\.[^.]+$/, '.md'), originalName: file.name, fileSize: `${fileSizeMB} MB`, lineCount: markdown.split('\n').length, charCount: markdown.length });
  } catch (error: unknown) {
    const message = error instanceof Error ? error.message : 'Unknown error';
    if (message.includes('timeout')) return NextResponse.json({ error: '변환 시간이 초과되었습니다.' }, { status: 504 });
    return NextResponse.json({ error: `변환 오류: ${message}` }, { status: 500 });
  } finally {
    for (const p of [tempPath, tempHtmlPath]) {
      if (p) { try { await unlink(p); } catch { /* */ } }
    }
  }
}

export const config = { api: { bodyParser: false } };
