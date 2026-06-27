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
};

export default nextConfig;
