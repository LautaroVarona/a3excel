import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  serverExternalPackages: ["officecrypto-tool", "xlsx", "cfb"],
  webpack: (config, { dev }) => {
    if (dev) {
      config.cache = false;
    }
    return config;
  },
};

export default nextConfig;
