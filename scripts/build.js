#!/usr/bin/env node
const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const SRC = path.join(ROOT, 'src', 'app.jsx');
const DIST = path.join(ROOT, 'dist');
const OUT_BUNDLE = path.join(DIST, 'bundle.js');
const OUT_HTML = path.join(DIST, 'index.html');

// Ensure dist directory exists
fs.mkdirSync(DIST, { recursive: true });

// Run esbuild
console.log('Bundling with esbuild...');
execSync(
  'npx esbuild "' + SRC + '" --bundle --minify --format=iife --outfile="' + OUT_BUNDLE + '" --loader:.jsx=jsx --jsx=automatic',
  { cwd: ROOT, stdio: 'inherit' }
);

const bundleJS = fs.readFileSync(OUT_BUNDLE, 'utf-8');
const bundleSize = (Buffer.byteLength(bundleJS) / 1024).toFixed(0);
console.log('Bundle size: ' + bundleSize + ' KB');

// HTML template parts
const FAVICON = 'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCA0MCA0MCI+PHJlY3QgeD0iMiIgeT0iMiIgd2lkdGg9IjM2IiBoZWlnaHQ9IjM2IiByeD0iOCIgZmlsbD0iIzBhMGExNCIgc3Ryb2tlPSIjNkM5Q0ZGIiBzdHJva2Utd2lkdGg9IjIiLz48cmVjdCB4PSI4IiB5PSI4IiB3aWR0aD0iMTIiIGhlaWdodD0iMTIiIHJ4PSIyIiBmaWxsPSIjNkM5Q0ZGIi8+PHJlY3QgeD0iMjIiIHk9IjgiIHdpZHRoPSIxMCIgaGVpZ2h0PSI1IiByeD0iMS41IiBmaWxsPSIjNEFERTgwIi8+PHJlY3QgeD0iMjIiIHk9IjE1IiB3aWR0aD0iMTAiIGhlaWdodD0iNSIgcng9IjEuNSIgZmlsbD0iIzRBREU4MCIgZmlsbC1vcGFjaXR5PSIuNSIvPjxyZWN0IHg9IjgiIHk9IjIyIiB3aWR0aD0iMjQiIGhlaWdodD0iNCIgcng9IjEuNSIgZmlsbD0iIzZDOUNGRiIgZmlsbC1vcGFjaXR5PSIuNCIvPjxyZWN0IHg9IjgiIHk9IjI4IiB3aWR0aD0iMTgiIGhlaWdodD0iNCIgcng9IjEuNSIgZmlsbD0iIzZDOUNGRiIgZmlsbC1vcGFjaXR5PSIuMjUiLz48L3N2Zz4=';

const CSS = '*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}body{font-family:\'Inter\',system-ui,-apple-system,sans-serif;background:#0a0a14;color:#e8e8e8;min-height:100vh}#root{max-width:960px;margin:0 auto;padding:28px 16px}input[type=number]{-moz-appearance:textfield}input::-webkit-outer-spin-button,input::-webkit-inner-spin-button{-webkit-appearance:none}::-webkit-scrollbar{width:6px;height:6px}::-webkit-scrollbar-track{background:transparent}::-webkit-scrollbar-thumb{background:#333;border-radius:3px}table{border-spacing:0}@media(max-width:640px){.grid-2{grid-template-columns:1fr!important}.grid-3{grid-template-columns:1fr!important}}';

// Build self-contained HTML
var html = [
  '<!DOCTYPE html>',
  '<html lang="es">',
  '<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">',
  '<title>FiscalBox — Gestión Impositiva para Estudios Contables</title>',
  '<meta name="theme-color" content="#0a0a14">',
  '<link rel="icon" type="image/svg+xml" href="' + FAVICON + '">',
  '<style>' + CSS + '</style>',
  '</head>',
  '<body><div id="root"></div>',
  '<script>' + bundleJS + '</script>',
  '</body></html>'
].join('\n');

fs.writeFileSync(OUT_HTML, html, 'utf-8');
var htmlSize = (Buffer.byteLength(html) / 1024).toFixed(0);
console.log('Output: dist/index.html (' + htmlSize + ' KB)');
console.log('Build complete!');
