#!/usr/bin/env node
// Development server: builds and serves with live reload
const { execSync } = require('child_process');
const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = process.env.PORT || 3000;
const DIST = path.resolve(__dirname, '..', 'dist');

// Build first
console.log('Building...');
execSync('node ' + path.join(__dirname, 'build.js'), { stdio: 'inherit' });

// Simple static server
const server = http.createServer(function(req, res) {
  var filePath = path.join(DIST, req.url === '/' ? 'index.html' : req.url);
  if (!fs.existsSync(filePath)) {
    res.writeHead(404);
    res.end('Not found');
    return;
  }
  var ext = path.extname(filePath);
  var contentType = ext === '.html' ? 'text/html' :
                    ext === '.js' ? 'application/javascript' :
                    ext === '.css' ? 'text/css' : 'application/octet-stream';
  res.writeHead(200, { 'Content-Type': contentType + '; charset=utf-8' });
  res.end(fs.readFileSync(filePath));
});

server.listen(PORT, function() {
  console.log('FiscalBox dev server running at http://localhost:' + PORT);
  console.log('Press Ctrl+C to stop');
});
