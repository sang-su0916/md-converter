'use client';

import { useState, useCallback, useRef } from 'react';
import './globals.css';

interface ConvertResult {
  markdown: string;
  filename: string;
  originalName: string;
  fileSize: string;
  lineCount: number;
  charCount: number;
}

const ACCEPTED_TYPES: Record<string, string[]> = {
  'application/pdf': ['.pdf'],
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx'],
  'application/vnd.openxmlformats-officedocument.presentationml.presentation': ['.pptx'],
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
  'text/html': ['.html', '.htm'],
  'text/csv': ['.csv'],
  'application/json': ['.json'],
  'application/epub+zip': ['.epub'],
  'image/*': ['.jpg', '.jpeg', '.png', '.gif', '.webp'],
  'application/x-hwp': ['.hwp'],
  'application/haansofthwpx': ['.hwpx'],
};

type ViewMode = 'preview' | 'raw';

export default function Home() {
  const [file, setFile] = useState<File | null>(null);
  const [result, setResult] = useState<ConvertResult | null>(null);
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);
  const [progress, setProgress] = useState(0);
  const [viewMode, setViewMode] = useState<ViewMode>('preview');
  const [dragActive, setDragActive] = useState(false);
  const [copied, setCopied] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFile = useCallback((f: File) => {
    setFile(f);
    setResult(null);
    setError('');
  }, []);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setDragActive(false);
    if (e.dataTransfer.files?.[0]) {
      handleFile(e.dataTransfer.files[0]);
    }
  }, [handleFile]);

  const handleConvert = async () => {
    if (!file) return;
    setLoading(true);
    setError('');
    setResult(null);
    setProgress(10);

    const progressInterval = setInterval(() => {
      setProgress(p => Math.min(p + Math.random() * 15, 90));
    }, 500);

    try {
      const formData = new FormData();
      formData.append('file', file);

      const res = await fetch('/api/convert', { method: 'POST', body: formData });
      const data = await res.json();

      if (!res.ok) {
        setError(data.error || '변환에 실패했습니다.');
      } else {
        setResult(data);
      }
    } catch {
      setError('서버 연결에 실패했습니다. 잠시 후 다시 시도해 주세요.');
    } finally {
      clearInterval(progressInterval);
      setProgress(100);
      setTimeout(() => { setLoading(false); setProgress(0); }, 300);
    }
  };

  const handleDownload = () => {
    if (!result) return;
    const blob = new Blob([result.markdown], { type: 'text/markdown;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = result.filename;
    a.click();
    URL.revokeObjectURL(url);
  };

  const handleCopy = async () => {
    if (!result) return;
    await navigator.clipboard.writeText(result.markdown);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const handleReset = () => {
    setFile(null);
    setResult(null);
    setError('');
    setCopied(false);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const formatIcon = (name: string) => {
    const ext = name.split('.').pop()?.toLowerCase() || '';
    const icons: Record<string, string> = {
      pdf: '📄', docx: '📝', doc: '📝', hwp: '📃', hwpx: '📃', pptx: '📊', ppt: '📊',
      xlsx: '📈', xls: '📈', html: '🌐', htm: '🌐', csv: '📋',
      json: '⚙️', epub: '📚', jpg: '🖼️', jpeg: '🖼️', png: '🖼️',
      gif: '🖼️', webp: '🖼️', mp3: '🎵', wav: '🎵',
    };
    return icons[ext] || '📎';
  };

  return (
    <div style={{ minHeight: '100vh', display: 'flex', flexDirection: 'column' }}>
      {/* Header with Branding */}
      <header style={{
        padding: '16px 32px',
        borderBottom: '1px solid var(--border)',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'space-between',
        background: 'var(--surface)',
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 14 }}>
          <img
            src="/logo.png"
            alt="엘비즈파트너스"
            style={{ height: 56, width: 'auto', borderRadius: 8, filter: 'drop-shadow(0 2px 8px rgba(201,168,76,0.3))' }}
          />
          <div>
            <h1 style={{ fontSize: 24, fontWeight: 800, color: '#fff', letterSpacing: '-0.3px' }}>
              MD Converter
            </h1>
            <p style={{ fontSize: 13, color: 'var(--text-muted)', marginTop: 2 }}>
              파일을 마크다운으로 · <span style={{ color: '#C9A84C', fontWeight: 700 }}>엘비즈파트너스</span>
            </p>
          </div>
        </div>
        <div style={{ display: 'flex', gap: 6, fontSize: 11, color: 'var(--text-muted)', flexWrap: 'wrap', justifyContent: 'flex-end' }}>
          {[
            { label: 'PDF', color: '#6c63ff', bg: 'rgba(108,99,255,0.12)' },
            { label: 'DOCX', color: '#22c55e', bg: 'rgba(34,197,94,0.12)' },
            { label: 'HWP', color: '#a855f7', bg: 'rgba(168,85,247,0.12)' },
            { label: 'PPTX', color: '#f97316', bg: 'rgba(249,115,22,0.12)' },
            { label: '+12', color: '#ec4899', bg: 'rgba(236,72,153,0.12)' },
          ].map(t => (
            <span key={t.label} style={{
              background: t.bg, color: t.color,
              padding: '3px 10px', borderRadius: 12, fontWeight: 600, fontSize: 11,
            }}>
              {t.label}
            </span>
          ))}
        </div>
      </header>

      <main style={{ flex: 1, display: 'flex', flexDirection: 'column', padding: '28px 32px', gap: 20, maxWidth: 1200, width: '100%', margin: '0 auto' }}>
        {/* Brand Hero */}
        {!result && (
          <div style={{
            background: 'linear-gradient(135deg, rgba(201,168,76,0.08) 0%, rgba(108,99,255,0.06) 50%, rgba(201,168,76,0.04) 100%)',
            borderRadius: 16,
            padding: '32px 28px',
            textAlign: 'center',
            border: '1px solid rgba(201,168,76,0.15)',
            position: 'relative',
            overflow: 'hidden',
          }}>
            <div style={{
              position: 'absolute', top: 0, left: 0, right: 0, height: 3,
              background: 'linear-gradient(90deg, #C9A84C, #6c63ff, #C9A84C)',
            }} />
            <img
              src="/logo.png"
              alt="엘비즈파트너스"
              style={{ height: 72, width: 'auto', marginBottom: 14, filter: 'drop-shadow(0 4px 12px rgba(201,168,76,0.25))' }}
            />
            <h2 style={{ fontSize: 26, fontWeight: 800, color: '#fff', marginBottom: 6 }}>
              문서를 <span style={{ color: '#C9A84C' }}>마크다운</span>으로, 한 번에
            </h2>
            <p style={{ fontSize: 14, color: 'var(--text-muted)', lineHeight: 1.7, maxWidth: 520, margin: '0 auto' }}>
              PDF · DOCX · HWP · PPTX 등 <strong style={{ color: '#a855f7' }}>20종+</strong> 문서를 마크다운으로 변환하세요.
              <br />옵시디언 · 노션 · GitHub에 바로 사용할 수 있습니다.
            </p>
          </div>
        )}

        {/* Upload Area */}
        {!result && (
          <div
            onDragOver={e => { e.preventDefault(); setDragActive(true); }}
            onDragLeave={() => setDragActive(false)}
            onDrop={handleDrop}
            onClick={() => fileInputRef.current?.click()}
            style={{
              border: `2px dashed ${dragActive ? 'var(--accent)' : 'var(--border)'}`,
              borderRadius: 16,
              padding: file ? '28px' : '56px 32px',
              textAlign: 'center',
              cursor: 'pointer',
              background: dragActive ? 'var(--accent-glow)' : 'var(--surface)',
              transition: 'all 0.2s ease',
            }}
          >
            <input
              ref={fileInputRef}
              type="file"
              onChange={e => e.target.files?.[0] && handleFile(e.target.files[0])}
              style={{ display: 'none' }}
              accept={Object.entries(ACCEPTED_TYPES).flatMap(([k, v]) => [k, ...v]).join(',')}
            />
            {!file ? (
              <>
                <div style={{ fontSize: 44, marginBottom: 12 }}>📂</div>
                <p style={{ fontSize: 17, fontWeight: 700, color: '#fff', marginBottom: 6 }}>
                  파일을 드래그하거나 클릭하여 선택
                </p>
                <p style={{ fontSize: 13, color: 'var(--text-muted)' }}>
                  PDF · DOCX · HWP · PPTX · XLSX · HTML · 이미지 · EPUB 등 15+ 포맷
                </p>
              </>
            ) : (
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 16 }}>
                <span style={{ fontSize: 36 }}>{formatIcon(file.name)}</span>
                <div style={{ textAlign: 'left' }}>
                  <p style={{ fontSize: 15, fontWeight: 700, color: '#fff' }}>{file.name}</p>
                  <p style={{ fontSize: 12, color: 'var(--text-muted)' }}>
                    {(file.size / (1024 * 1024)).toFixed(2)} MB
                  </p>
                </div>
                <button
                  onClick={e => { e.stopPropagation(); handleReset(); }}
                  style={{
                    background: 'rgba(239,68,68,0.12)', color: '#ef4444', border: 'none',
                    padding: '7px 14px', borderRadius: 8, cursor: 'pointer', fontWeight: 600, fontSize: 12,
                  }}
                >
                  제거
                </button>
              </div>
            )}
          </div>
        )}

        {/* Convert Button */}
        {file && !result && (
          <button
            onClick={handleConvert}
            disabled={loading}
            style={{
              background: loading ? 'var(--border)' : 'linear-gradient(135deg, #6c63ff, #a855f7)',
              color: '#fff',
              border: 'none',
              padding: '14px 32px',
              borderRadius: 12,
              fontSize: 15,
              fontWeight: 700,
              cursor: loading ? 'not-allowed' : 'pointer',
              transition: 'all 0.2s',
              position: 'relative',
              overflow: 'hidden',
            }}
          >
            {loading && (
              <div style={{
                position: 'absolute', left: 0, top: 0, bottom: 0,
                width: `${progress}%`, background: 'linear-gradient(135deg, #6c63ff, #a855f7)',
                transition: 'width 0.3s ease',
              }} />
            )}
            <span style={{ position: 'relative', zIndex: 1 }}>
              {loading ? `변환 중... ${Math.round(progress)}%` : '✨ 마크다운으로 변환'}
            </span>
          </button>
        )}

        {/* Error */}
        {error && (
          <div style={{
            background: 'rgba(239,68,68,0.08)', border: '1px solid rgba(239,68,68,0.25)',
            borderRadius: 12, padding: 18, color: '#ef4444',
          }}>
            <p style={{ fontWeight: 700, marginBottom: 4 }}>⚠️ 오류</p>
            <p style={{ fontSize: 13, whiteSpace: 'pre-wrap' }}>{error}</p>
          </div>
        )}

        {/* Result */}
        {result && (
          <div style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: 14 }}>
            {/* Toolbar */}
            <div style={{
              display: 'flex', justifyContent: 'space-between', alignItems: 'center',
              background: 'var(--surface)', borderRadius: 12, padding: '10px 18px',
              border: '1px solid var(--border)', flexWrap: 'wrap', gap: 10,
            }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap' }}>
                <span style={{ fontSize: 13, color: 'var(--text-muted)' }}>
                  {formatIcon(result.originalName)} {result.originalName}
                </span>
                <span style={{ fontSize: 11, color: 'var(--text-muted)', background: 'var(--bg)', padding: '2px 8px', borderRadius: 6 }}>
                  {result.fileSize}
                </span>
                <span style={{ fontSize: 11, color: 'var(--text-muted)', background: 'var(--bg)', padding: '2px 8px', borderRadius: 6 }}>
                  {result.lineCount.toLocaleString()} lines · {result.charCount.toLocaleString()} chars
                </span>
              </div>

              <div style={{ display: 'flex', gap: 6 }}>
                <div style={{ display: 'flex', background: 'var(--bg)', borderRadius: 8, overflow: 'hidden' }}>
                  {(['preview', 'raw'] as ViewMode[]).map(mode => (
                    <button
                      key={mode}
                      onClick={() => setViewMode(mode)}
                      style={{
                        padding: '5px 12px', border: 'none', fontSize: 12, fontWeight: 600, cursor: 'pointer',
                        background: viewMode === mode ? 'var(--accent)' : 'transparent',
                        color: viewMode === mode ? '#fff' : 'var(--text-muted)',
                        transition: 'all 0.2s',
                      }}
                    >
                      {mode === 'preview' ? '미리보기' : '원본'}
                    </button>
                  ))}
                </div>

                <button onClick={handleCopy} style={{
                  background: copied ? 'rgba(34,197,94,0.15)' : 'var(--bg)',
                  color: copied ? '#22c55e' : 'var(--text)',
                  border: `1px solid ${copied ? 'rgba(34,197,94,0.3)' : 'var(--border)'}`,
                  padding: '5px 12px', borderRadius: 8, cursor: 'pointer', fontSize: 12, fontWeight: 600,
                  transition: 'all 0.2s',
                }}>
                  {copied ? '✅ 복사됨' : '📋 복사'}
                </button>
                <button onClick={handleDownload} style={{
                  background: 'linear-gradient(135deg, #6c63ff, #a855f7)', color: '#fff', border: 'none',
                  padding: '5px 14px', borderRadius: 8, cursor: 'pointer', fontSize: 12, fontWeight: 700,
                }}>
                  ⬇️ 다운로드
                </button>
                <button onClick={handleReset} style={{
                  background: 'transparent', color: 'var(--text-muted)', border: '1px solid var(--border)',
                  padding: '5px 12px', borderRadius: 8, cursor: 'pointer', fontSize: 12, fontWeight: 600,
                }}>
                  🔄 새 파일
                </button>
              </div>
            </div>

            {/* Content */}
            <div style={{
              flex: 1, background: 'var(--surface)', borderRadius: 12,
              border: '1px solid var(--border)', padding: 24, overflow: 'auto',
              minHeight: 400, maxHeight: 'calc(100vh - 300px)',
            }}>
              {viewMode === 'preview' ? (
                <MarkdownPreview content={result.markdown} />
              ) : (
                <pre style={{
                  fontFamily: '"SF Mono", "Fira Code", monospace',
                  fontSize: 12.5, lineHeight: 1.7, whiteSpace: 'pre-wrap', wordBreak: 'break-word',
                  color: 'var(--text)',
                }}>
                  {result.markdown}
                </pre>
              )}
            </div>
          </div>
        )}

        {/* 📘 설명서 & 사용법 */}
        <details style={{
          background: 'var(--surface)',
          border: '1px solid var(--border)',
          borderRadius: 14,
          overflow: 'hidden',
        }}>
          <summary style={{
            padding: '16px 22px',
            cursor: 'pointer',
            fontSize: 15,
            fontWeight: 700,
            color: 'var(--text)',
            display: 'flex',
            alignItems: 'center',
            gap: 10,
            userSelect: 'none',
            listStyle: 'none',
          }}>
            <span>📘</span> 설명서 & 사용법
            <span style={{ marginLeft: 'auto', fontSize: 11, color: 'var(--text-muted)', fontWeight: 400 }}>클릭하여 펼치기</span>
          </summary>
          <div style={{ padding: '4px 22px 24px', fontSize: 13, color: 'var(--text-muted)', lineHeight: 1.85 }}>

            {/* 앱 소개 */}
            <div style={{
              background: 'linear-gradient(135deg, rgba(108,99,255,0.08), rgba(168,85,247,0.08))',
              borderRadius: 12, padding: '18px 20px', marginBottom: 18,
              borderLeft: '3px solid #6c63ff',
            }}>
              <h3 style={{ fontSize: 15, fontWeight: 800, color: '#fff', marginBottom: 8 }}>
                MD Converter란?
              </h3>
              <p style={{ margin: 0, fontSize: 13, lineHeight: 1.8 }}>
                다양한 형식의 문서를 <strong style={{ color: '#a855f7' }}>마크다운(.md)</strong>으로 변환하는 도구입니다.
                변환된 파일은 <strong style={{ color: '#C9A84C' }}>옵시디언, 노션, GitHub</strong> 등 마크다운 기반 도구에 바로 사용할 수 있습니다.
                <br />
                Microsoft의 오픈소스 <code style={{ background: 'rgba(255,255,255,0.08)', padding: '1px 6px', borderRadius: 4, fontSize: 12 }}>MarkItDown</code> 엔진과
                한글(HWP) 전용 변환기를 탑재하여 한국어 문서도 높은 품질로 변환합니다.
              </p>
            </div>

            {/* 접속 안내 */}
            <div style={{
              display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12, marginBottom: 16,
            }}>
              <a href="https://md-converter-ghdf.onrender.com" style={{
                background: 'rgba(108,99,255,0.08)', borderRadius: 10, padding: '14px 16px',
                border: '1px solid rgba(108,99,255,0.2)', textDecoration: 'none', display: 'block',
              }}>
                <div style={{ fontSize: 12, fontWeight: 700, color: '#6c63ff', marginBottom: 4 }}>🌐 메인 서버 (Render)</div>
                <div style={{ fontSize: 11, color: 'var(--text-muted)' }}>PDF · DOCX · <strong style={{ color: '#a855f7' }}>HWP</strong> · PPTX · 이미지(OCR) · 오디오(STT) 등 20종+</div>
                <div style={{ fontSize: 10, color: 'var(--text-muted)', marginTop: 4, opacity: 0.7 }}>항상 사용 가능 · Docker 환경</div>
              </a>
              <a href="https://md-converter-drab.vercel.app" style={{
                background: 'rgba(34,197,94,0.08)', borderRadius: 10, padding: '14px 16px',
                border: '1px solid rgba(34,197,94,0.2)', textDecoration: 'none', display: 'block',
              }}>
                <div style={{ fontSize: 12, fontWeight: 700, color: '#22c55e', marginBottom: 4 }}>🚀 미러 서버 (Vercel)</div>
                <div style={{ fontSize: 11, color: 'var(--text-muted)' }}>전 포맷 지원 (특수 포맷은 백엔드 자동 연결)</div>
                <div style={{ fontSize: 10, color: 'var(--text-muted)', marginTop: 4, opacity: 0.7 }}>빠른 응답 · CDN 가속</div>
              </a>
            </div>

            {/* 사용법 + 지원 포맷 2컬럼 */}
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(320px, 1fr))', gap: 16, marginBottom: 16 }}>

              {/* 사용법 */}
              <div style={{ background: 'var(--bg)', borderRadius: 12, padding: '18px 20px' }}>
                <h4 style={{ fontSize: 14, fontWeight: 700, color: '#6c63ff', marginBottom: 12 }}>
                  🚀 사용 방법
                </h4>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
                  {[
                    { step: '1', title: '파일 업로드', desc: '파일을 드래그&드롭 하거나, 영역을 클릭하여 파일 선택' },
                    { step: '2', title: '변환 실행', desc: '"✨ 마크다운으로 변환" 버튼을 클릭하면 자동으로 변환 시작' },
                    { step: '3', title: '결과 확인', desc: '미리보기 탭에서 렌더링된 결과를, 원본 탭에서 마크다운 소스를 확인' },
                    { step: '4', title: '내보내기', desc: '📋 복사 버튼으로 클립보드에 복사하거나, ⬇️ 다운로드로 .md 파일 저장' },
                  ].map(s => (
                    <div key={s.step} style={{ display: 'flex', gap: 12, alignItems: 'flex-start' }}>
                      <span style={{
                        minWidth: 28, height: 28, borderRadius: '50%',
                        background: 'linear-gradient(135deg, #6c63ff, #a855f7)',
                        color: '#fff', fontSize: 13, fontWeight: 800,
                        display: 'flex', alignItems: 'center', justifyContent: 'center',
                      }}>{s.step}</span>
                      <div>
                        <strong style={{ color: 'var(--text)', fontSize: 13 }}>{s.title}</strong>
                        <p style={{ margin: '2px 0 0', fontSize: 12, lineHeight: 1.6 }}>{s.desc}</p>
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              {/* 지원 포맷 */}
              <div style={{ background: 'var(--bg)', borderRadius: 12, padding: '18px 20px' }}>
                <h4 style={{ fontSize: 14, fontWeight: 700, color: '#22c55e', marginBottom: 12 }}>
                  📁 지원 포맷 (20종+)
                </h4>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
                  {[
                    { icon: '📄', ext: 'PDF', note: '텍스트 추출 · 표 포함' },
                    { icon: '📝', ext: 'DOCX / DOC', note: 'MS Word 문서' },
                    { icon: '📃', ext: 'HWP / HWPX', note: '한글 문서 · 표 구조 보존' },
                    { icon: '📊', ext: 'PPTX / PPT', note: '슬라이드별 텍스트 추출' },
                    { icon: '📈', ext: 'XLSX / XLS', note: '시트별 마크다운 표 변환' },
                    { icon: '🌐', ext: 'HTML / HTM', note: '태그 → 마크다운 구조 변환' },
                    { icon: '📋', ext: 'CSV / JSON / XML', note: '데이터 → 표/구조 변환' },
                    { icon: '📚', ext: 'EPUB', note: '전자책 텍스트 추출' },
                    { icon: '🖼️', ext: 'JPG / PNG / GIF / WebP', note: 'OCR 텍스트 인식' },
                    { icon: '🎵', ext: 'MP3 / WAV / M4A', note: '음성→텍스트(STT)' },
                    { icon: '📋', ext: 'TXT / MD / RST / LOG', note: '그대로 통과' },
                    { icon: '📦', ext: 'ZIP', note: '압축 내부 파일 변환' },
                  ].map(f => (
                    <div key={f.ext} style={{
                      display: 'flex', alignItems: 'center', gap: 8,
                      padding: '5px 0', borderBottom: '1px solid rgba(255,255,255,0.03)',
                      fontSize: 12,
                    }}>
                      <span style={{ width: 20, textAlign: 'center' }}>{f.icon}</span>
                      <span style={{ fontWeight: 700, color: 'var(--text)', minWidth: 140 }}>{f.ext}</span>
                      <span style={{ color: 'var(--text-muted)', fontSize: 11 }}>{f.note}</span>
                    </div>
                  ))}
                </div>
              </div>
            </div>

            {/* 팁 + FAQ 2컬럼 */}
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(320px, 1fr))', gap: 16 }}>

              {/* 활용 팁 */}
              <div style={{ background: 'var(--bg)', borderRadius: 12, padding: '18px 20px' }}>
                <h4 style={{ fontSize: 14, fontWeight: 700, color: '#f97316', marginBottom: 12 }}>
                  💡 활용 팁
                </h4>
                <ul style={{ paddingLeft: 18, margin: 0, display: 'flex', flexDirection: 'column', gap: 8 }}>
                  <li><strong style={{ color: 'var(--text)' }}>옵시디언 연동</strong> — 변환 결과를 복사해 옵시디언에 바로 붙여넣기하면 서식이 유지됩니다</li>
                  <li><strong style={{ color: 'var(--text)' }}>표 변환</strong> — HWP, DOCX, XLSX의 표는 마크다운 테이블 문법으로 자동 변환됩니다</li>
                  <li><strong style={{ color: 'var(--text)' }}>원본 편집</strong> — 원본 탭에서 마크다운 소스를 확인하고, 복사 후 직접 수정할 수 있습니다</li>
                  <li><strong style={{ color: 'var(--text)' }}>대용량 파일</strong> — 100MB까지 업로드 가능하며, HWP는 최대 2분까지 변환 시간이 소요될 수 있습니다</li>
                  <li><strong style={{ color: 'var(--text)' }}>이미지 OCR</strong> — 사진이나 스캔 이미지를 업로드하면 텍스트를 자동 인식합니다</li>
                </ul>
              </div>

              {/* FAQ */}
              <div style={{ background: 'var(--bg)', borderRadius: 12, padding: '18px 20px' }}>
                <h4 style={{ fontSize: 14, fontWeight: 700, color: '#ec4899', marginBottom: 12 }}>
                  ❓ 자주 묻는 질문
                </h4>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
                  {[
                    { q: '모든 포맷이 어디서든 변환 가능한가요?', a: '네! HWP, 이미지(OCR), 오디오(STT) 등 특수 포맷도 자동으로 백엔드 서버를 통해 변환됩니다. 모든 접속 주소에서 20종+ 전 포맷을 지원합니다.' },
                    { q: 'HWP 변환이 느린 경우는?', a: '첫 요청 시 백엔드 서버가 대기 상태에서 깨어나는 데 30초~1분 정도 걸릴 수 있습니다. 이후 요청은 빠르게 처리됩니다.' },
                    { q: '변환된 파일은 서버에 저장되나요?', a: '아니요. 변환 후 즉시 삭제됩니다. 파일은 외부로 전송되지 않습니다.' },
                    { q: '마크다운이란 무엇인가요?', a: '텍스트 기반의 경량 문서 포맷입니다. 옵시디언, 노션, GitHub 등에서 바로 사용할 수 있습니다.' },
                  ].map((item, i) => (
                    <div key={i}>
                      <p style={{ margin: 0, fontSize: 12, fontWeight: 700, color: 'var(--text)' }}>Q. {item.q}</p>
                      <p style={{ margin: '3px 0 0', fontSize: 12, lineHeight: 1.6 }}>A. {item.a}</p>
                    </div>
                  ))}
                </div>
              </div>

            </div>
          </div>
        </details>
      </main>

      {/* Footer with Branding */}
      <footer style={{
        padding: '24px 32px 20px',
        borderTop: '1px solid var(--border)',
        background: 'linear-gradient(180deg, var(--surface) 0%, rgba(201,168,76,0.03) 100%)',
      }}>
        <div style={{
          maxWidth: 1200, margin: '0 auto',
          display: 'flex', justifyContent: 'space-between', alignItems: 'center',
          flexWrap: 'wrap', gap: 16,
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 14 }}>
            <img src="/logo.png" alt="엘비즈파트너스" style={{ height: 40, width: 'auto', filter: 'drop-shadow(0 2px 6px rgba(201,168,76,0.2))' }} />
            <div style={{ lineHeight: 1.6 }}>
              <div style={{ fontSize: 14, fontWeight: 700 }}>
                <span style={{ color: '#C9A84C' }}>엘비즈파트너스</span>
              </div>
              <div style={{ fontSize: 12, color: 'var(--text-muted)' }}>
                세무·노무·법무 컨설팅 & AI 활용 교육
              </div>
            </div>
          </div>
          <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-end', gap: 8 }}>
            <div style={{ fontSize: 12, color: 'var(--text-muted)', display: 'flex', gap: 16, alignItems: 'center' }}>
              <a href="tel:010-3709-5785" style={{ color: 'var(--text-muted)', textDecoration: 'none' }}>
                📞 010-3709-5785
              </a>
              <a href="mailto:sangsu0916@naver.com" style={{ color: 'var(--text-muted)', textDecoration: 'none' }}>
                ✉️ sangsu0916@naver.com
              </a>
            </div>
            <a href="https://lbiz-partners.com" target="_blank" rel="noopener noreferrer" style={{
              color: '#C9A84C', textDecoration: 'none', fontWeight: 700, fontSize: 13,
              padding: '6px 16px', borderRadius: 8,
              border: '1px solid rgba(201,168,76,0.3)',
              background: 'rgba(201,168,76,0.08)',
              transition: 'all 0.2s',
            }}>
              🌐 lbiz-partners.com →
            </a>
          </div>
        </div>
        <div style={{ maxWidth: 1200, margin: '14px auto 0', fontSize: 10, color: 'rgba(136,136,160,0.4)', textAlign: 'center', borderTop: '1px solid rgba(255,255,255,0.04)', paddingTop: 12 }}>
          © 2026 엘비즈파트너스. Powered by MarkItDown (Microsoft) · Built with Next.js
        </div>
      </footer>
    </div>
  );
}

function MarkdownPreview({ content }: { content: string }) {
  const html = simpleMarkdownToHtml(content);
  return <div className="md-preview" dangerouslySetInnerHTML={{ __html: html }} />;
}

function simpleMarkdownToHtml(md: string): string {
  let html = md
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    // Headers
    .replace(/^######\s+(.+)$/gm, '<h6>$1</h6>')
    .replace(/^#####\s+(.+)$/gm, '<h5>$1</h5>')
    .replace(/^####\s+(.+)$/gm, '<h4>$1</h4>')
    .replace(/^###\s+(.+)$/gm, '<h3>$1</h3>')
    .replace(/^##\s+(.+)$/gm, '<h2>$1</h2>')
    .replace(/^#\s+(.+)$/gm, '<h1>$1</h1>')
    // Bold & Italic
    .replace(/\*\*\*(.+?)\*\*\*/g, '<strong><em>$1</em></strong>')
    .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
    .replace(/\*(.+?)\*/g, '<em>$1</em>')
    // Inline code
    .replace(/`([^`]+)`/g, '<code>$1</code>')
    // Blockquotes
    .replace(/^&gt;\s+(.+)$/gm, '<blockquote>$1</blockquote>')
    // Horizontal rule
    .replace(/^---+$/gm, '<hr>')
    // Links
    .replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2" target="_blank">$1</a>')
    // Tables - basic support
    .replace(/^\|(.+)\|$/gm, (match) => {
      const cells = match.split('|').filter(c => c.trim() !== '');
      if (cells.every(c => /^[\s-:]+$/.test(c))) return ''; // separator row
      const tag = 'td';
      const row = cells.map(c => `<${tag}>${c.trim()}</${tag}>`).join('');
      return `<tr>${row}</tr>`;
    })
    // Unordered lists
    .replace(/^[\-\*]\s+(.+)$/gm, '<li>$1</li>')
    // Ordered lists
    .replace(/^\d+\.\s+(.+)$/gm, '<li>$1</li>')
    // Paragraphs
    .replace(/\n\n/g, '</p><p>')
    .replace(/\n/g, '<br>');

  // Wrap tr groups in table
  html = html.replace(/((<tr>.*?<\/tr>(<br>)?)+)/g, '<table>$1</table>');
  html = html.replace(/<table>(.*?)<\/table>/g, (match) => match.replace(/<br>/g, ''));

  // Wrap li groups in ul
  html = html.replace(/((<li>.*?<\/li>(<br>)?)+)/g, '<ul>$1</ul>');
  html = html.replace(/<ul>(.*?)<\/ul>/g, (match) => match.replace(/<br>/g, ''));

  html = '<p>' + html + '</p>';
  html = html.replace(/<p>\s*<\/p>/g, '');
  html = html.replace(/<p>\s*(<h[1-6]>)/g, '$1');
  html = html.replace(/(<\/h[1-6]>)\s*<\/p>/g, '$1');
  html = html.replace(/<p>\s*(<hr>)\s*<\/p>/g, '$1');
  html = html.replace(/<p>\s*(<ul>)/g, '$1');
  html = html.replace(/(<\/ul>)\s*<\/p>/g, '$1');
  html = html.replace(/<p>\s*(<table>)/g, '$1');
  html = html.replace(/(<\/table>)\s*<\/p>/g, '$1');
  html = html.replace(/<p>\s*(<blockquote>)/g, '$1');
  html = html.replace(/(<\/blockquote>)\s*<\/p>/g, '$1');

  return html;
}
