import type { NextConfig } from 'next';

const nextConfig: NextConfig = {
  // standalone only for Docker (Render), not needed for Vercel
  ...(process.env.DOCKER_BUILD === 'true' ? { output: 'standalone' } : {}),
  serverExternalPackages: ['officeparser', 'pdfjs-dist', 'pdf-parse', 'turndown'],
};

export default nextConfig;
