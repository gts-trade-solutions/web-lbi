/** @type {import('next').NextConfig} */
const nextConfig = {
  eslint: {
    // Allows production builds to successfully complete even if ESLint errors are present.
    ignoreDuringBuilds: true,
  },
  typescript: {
    // Allows production builds to successfully complete even if TypeScript errors are present.
    ignoreBuildErrors: true,
  },
  experimental: {
    // Load pdfjs-dist at runtime from node_modules on the server instead of
    // bundling it (it pulls in optional native deps like canvas) — the PDF
    // report importer dynamically imports it in a nodejs-runtime route.
    serverComponentsExternalPackages: ["pdfjs-dist"],
  },
};

export default nextConfig;
