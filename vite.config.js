import { defineConfig } from 'vite';
import fs from 'fs';
import path from 'path';
import os from 'os';

const homedir = os.homedir();
const certDir = path.resolve(homedir, '.office-addin-dev-certs');

export default defineConfig({
  server: {
    port: 3000,
    https: {
      key: fs.readFileSync(path.resolve(certDir, 'localhost.key')),
      cert: fs.readFileSync(path.resolve(certDir, 'localhost.crt')),
      ca: fs.readFileSync(path.resolve(certDir, 'ca.crt')),
    },
    strictPort: true,
  },
});
