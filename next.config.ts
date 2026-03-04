import type { NextConfig } from 'next';

const nextConfig: NextConfig = {
  output: 'standalone',
  serverExternalPackages: ['officeparser', 'pdfjs-dist', 'pdf-parse', 'turndown'],
};

export default nextConfig;
