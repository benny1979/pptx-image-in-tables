#!/usr/bin/env node
// Stamp every relative import and asset link with the current commit, so a
// deploy actually reaches the task pane.
//
// GitHub Pages sends Cache-Control: max-age=600 and the Office task pane's
// WebView caches on top of that. On 2026-09-26 a fix was live on Pages for
// twenty minutes while the pane kept running the previous ppt.js - the HTML
// had refreshed but its imports had not, so the two disagreed and the bug
// looked unfixed. A query string is the only lever available without a build
// step or control over headers.
//
// Run it before committing:  npm run bust
//
// ES modules are cached per resolved URL, so EVERY link in the graph needs
// the stamp, not just the entry point - that is exactly what went wrong.

const { execFileSync } = require('child_process');
const fs = require('fs');
const path = require('path');

// execFileSync, not execSync: no shell, nothing to interpolate into.
const stamp = execFileSync('git', ['rev-parse', '--short', 'HEAD']).toString().trim();
const root = path.join(__dirname, '..', 'src');

const files = [];
(function walk(dir) {
  for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, e.name);
    if (e.isDirectory()) walk(full);
    else if (/\.(js|html)$/.test(e.name)) files.push(full);
  }
})(root);

let changed = 0;
for (const file of files) {
  const before = fs.readFileSync(file, 'utf8');
  const after = before
    // import ... from './x.js'  /  import('./x.js')
    .replace(/(from\s+|import\()(['"])(\.[^'"?]+\.js)(\?v=[^'"]*)?\2/g,
             (_m, kw, q, p2) => `${kw}${q}${p2}?v=${stamp}${q}`)
    // <script src="app.js">, <link href="taskpane.css">
    .replace(/(src|href)="([^":?]+\.(?:js|css))(\?v=[^"]*)?"/g,
             (_m, attr, p2) => `${attr}="${p2}?v=${stamp}"`)
    // the build id the diagnostics panel reports
    .replace(/^const BUILD = '[^']*';$/m, `const BUILD = '${stamp}';`);

  if (after !== before) {
    fs.writeFileSync(file, after);
    changed++;
  }
}

console.log(`bust-cache: stamped ${changed} file(s) with ${stamp}`);
