// Minimal HTTPS static server for sideloading.
//
// Office will not load a task pane over plain HTTP, and will not trust a
// self-signed certificate it does not know about. office-addin-dev-certs
// installs a CA into your local trust store and hands us a matching cert.

const fs = require('fs');
const http = require('http');
const https = require('https');
const path = require('path');

const PORT = Number(process.env.PORT || 3000);
const ROOT = __dirname;

const TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.xml': 'text/xml; charset=utf-8'
};

function handler(req, res) {
  const url = decodeURIComponent(req.url.split('?')[0]);
  const rel = url === '/' ? '/src/taskpane/taskpane.html' : url;
  const file = path.join(ROOT, path.normalize(rel));

  // Never serve outside the project directory.
  if (!file.startsWith(ROOT)) {
    res.writeHead(403).end('Forbidden');
    return;
  }

  fs.readFile(file, (err, body) => {
    if (err) {
      res.writeHead(404, { 'Content-Type': 'text/plain' }).end('Not found');
      return;
    }
    res.writeHead(200, {
      'Content-Type': TYPES[path.extname(file)] || 'application/octet-stream',
      'Cache-Control': 'no-store',
      'Access-Control-Allow-Origin': '*'
    });
    res.end(body);
  });
}

async function main() {
  let options;
  try {
    const devCerts = require('office-addin-dev-certs');
    options = await devCerts.getHttpsServerOptions();
  } catch (e) {
    console.error('Could not get development certificates.');
    console.error('Run:  npx office-addin-dev-certs install');
    console.error(e.message);
    process.exit(1);
  }

  https.createServer(options, handler).listen(PORT, () => {
    console.log(`Cell Images serving on https://localhost:${PORT}`);
    console.log(`Task pane:  https://localhost:${PORT}/src/taskpane/taskpane.html`);
    console.log('Leave this running while you use the add-in.');
  });
}

main();
