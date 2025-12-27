import { defineConfig } from "vite";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";

export default defineConfig(() => {
  const devCertDir = path.join(os.homedir(), ".office-addin-dev-certs");
  const keyPath = path.join(devCertDir, "localhost.key");
  const certPath = path.join(devCertDir, "localhost.crt");

  const https = {
    key: fs.readFileSync(keyPath),
    cert: fs.readFileSync(certPath),
  };

  return {
    server: {
      https,
      port: 3000,
      strictPort: true,
    },
  };
});
