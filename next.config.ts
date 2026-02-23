import type { NextConfig } from 'next';

const nextConfig: NextConfig = {
  serverExternalPackages: ['officeparser', 'pdfjs-dist', 'pdf-parse', 'turndown'],
};

export default nextConfig;
