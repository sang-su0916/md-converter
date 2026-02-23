import type { Metadata } from 'next';

export const metadata: Metadata = {
  title: 'MD Converter — 엘비즈파트너스',
  description: 'PDF, DOCX, HWP, PPTX 등 15+ 포맷을 마크다운으로 변환합니다. 무료 온라인 도구 by 엘비즈파트너스',
  icons: { icon: '/logo-sm.png' },
  openGraph: {
    title: 'MD Converter — 파일을 마크다운으로',
    description: 'PDF, DOCX, HWP 등 다양한 문서를 마크다운으로 변환하세요. 무료 · 빠름 · 안전',
    siteName: '엘비즈파트너스',
    locale: 'ko_KR',
    type: 'website',
  },
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="ko">
      <body style={{ margin: 0, fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Noto Sans KR", sans-serif' }}>
        {children}
      </body>
    </html>
  );
}
