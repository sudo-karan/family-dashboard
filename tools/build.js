#!/usr/bin/env node
/* Compiles the app's JSX (src/app.jsx) into plain JavaScript inside index.html,
 * between the GENERATED APP markers. Phones then run that script directly,
 * instead of downloading Babel (~3 MB) and compiling the app on every launch.
 *
 * It uses the very compiler and settings the page used to run in the browser —
 * @babel/standalone 7.26.4 with <script type="text/babel" data-presets="react">,
 * i.e. the "react" preset plus babel-standalone's default plugins — so the
 * code that runs is the same; only where it gets compiled has changed.
 *
 *   npm run build    rewrite the generated block in index.html
 *   npm run check    exit 1 if index.html is out of date with src/app.jsx (CI)
 */
'use strict';
const fs = require('fs');
const path = require('path');
const Babel = require('@babel/standalone');

const ROOT = path.join(__dirname, '..');
const SRC = path.join(ROOT, 'src', 'app.jsx');
const HTML = path.join(ROOT, 'index.html');
const BEGIN = '<!-- BEGIN GENERATED APP: compiled from src/app.jsx by `npm run build`. Edit that file, not this block. -->';
const END = '<!-- END GENERATED APP -->';

/** JSX source -> plain JS, exactly as babel-standalone did it in the browser. */
function compile(source) {
  const code = Babel.transform(source, {
    filename: 'src/app.jsx',
    presets: ['react'],
    plugins: ['transform-class-properties', 'transform-object-rest-spread', 'transform-flow-strip-types'],
  }).code;
  // The result is inlined in a <script>: these would end it early or confuse the HTML parser.
  const bad = /<\/script|<!--/i.exec(code);
  if (bad) {
    throw new Error('the compiled app contains "' + bad[0] + '", which cannot sit inside an inline <script>. ' +
      'Change that text in src/app.jsx (e.g. write "<\\/script" or split the string).');
  }
  return code;
}

/** The whole generated block, markers included. */
function block(code) {
  return BEGIN + '\n<script>\n' + code + '\n</script>\n' + END;
}

/** index.html with its generated block replaced by `code`. */
function withApp(html, code) {
  const a = html.indexOf(BEGIN);
  const b = html.indexOf(END);
  if (a < 0 || b < a || html.indexOf(BEGIN, a + 1) >= 0 || html.indexOf(END, b + 1) >= 0) {
    throw new Error('index.html must contain the GENERATED APP markers exactly once each, BEGIN before END.');
  }
  return html.slice(0, a) + block(code) + html.slice(b + END.length);
}

const lf = s => s.replace(/\r\n/g, '\n');

function main() {
  const check = process.argv.includes('--check');
  const html = fs.readFileSync(HTML, 'utf8');
  const next = withApp(lf(html), compile(lf(fs.readFileSync(SRC, 'utf8'))));
  if (check) {
    if (next !== lf(html)) {
      console.error('index.html is out of date with src/app.jsx. Run `npm run build` and commit index.html.');
      process.exit(1);
    }
    console.log('ok: index.html matches src/app.jsx');
  } else if (next === lf(html)) {
    console.log('index.html is already up to date');
  } else {
    fs.writeFileSync(HTML, next);
    console.log('index.html rebuilt from src/app.jsx');
  }
}

if (require.main === module) {
  try { main(); } catch (e) { console.error('build failed: ' + e.message); process.exit(1); }
}
module.exports = { compile, block, withApp, BEGIN, END };
