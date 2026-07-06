import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  // xlsx es pesado; officecrypto-tool y cfb se empaquetan en la función serverless.
  serverExternalPackages: ["xlsx"],
  webpack: (config, { dev }) => {
    if (dev) {
      config.cache = false;
    }
    return config;
  },
};

export default nextConfig;
