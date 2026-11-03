/* Local test server for sideloading the add-in during development.
 *
 * Serves the project root at https://localhost:3000 when the Office
 * add-in dev certificates are present (~/.office-addin-dev-certs/,
 * created by `npx office-addin-dev-certs install`); otherwise serves
 * plain HTTP and prints a warning — Excel requires HTTPS for the task
 * pane, so install the dev certs before sideloading manifest.dev.xml.
 *
 * Usage: npm start
 */
import fs from 'node:fs';
import http from 'node:http';
import https from 'node:https';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const PORT = 3000;
const TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.xml': 'text/xml; charset=utf-8',
  '.png': 'image/png',
  '.json': 'application/json',
  '.txt': 'text/plain; charset=utf-8',
  '.md': 'text/plain; charset=utf-8',
};

function handler(req, res) {
  const urlPath = decodeURIComponent(new URL(req.url, 'https://localhost').pathname);
  const file = path.join(root, urlPath === '/' ? 'index.html' : urlPath);
  if (!file.startsWith(root)) { res.writeHead(403); res.end(); return; }
  fs.readFile(file, (err, data) => {
    if (err) { res.writeHead(404); res.end('not found'); return; }
    res.writeHead(200, {
      'Content-Type': TYPES[path.extname(file)] || 'application/octet-stream',
      'Access-Control-Allow-Origin': '*',
      'Cache-Control': 'no-store',
    });
    res.end(data);
  });
}

const certDir = path.join(os.homedir(), '.office-addin-dev-certs');
const keyFile = path.join(certDir, 'localhost.key');
const certFile = path.join(certDir, 'localhost.crt');
if (fs.existsSync(keyFile) && fs.existsSync(certFile)) {
  https.createServer({ key: fs.readFileSync(keyFile), cert: fs.readFileSync(certFile) }, handler)
    .listen(PORT, () => console.log(`Serving ${root} at https://localhost:${PORT}`));
} else {
  http.createServer(handler)
    .listen(PORT, () => console.log(
      `Serving ${root} at http://localhost:${PORT} (no dev certs found; ` +
      'Excel needs HTTPS — run: npx office-addin-dev-certs install)'));
}
