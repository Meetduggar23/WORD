import http from 'http';
import fs from 'fs';
import path from 'path';
import { exec } from 'child_process';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const PORT = 8999;
const PUBLIC_DIR = path.join(__dirname, '../public');

// Ensure public directory exists
if (!fs.existsSync(PUBLIC_DIR)) {
  fs.mkdirSync(PUBLIC_DIR, { recursive: true });
}

const htmlPage = `
<!DOCTYPE html>
<html>
<head>
  <title>Icon Generator</title>
  <style>
    body { font-family: sans-serif; text-align: center; padding: 50px; background: #f8fafc; }
    h1 { color: #1e293b; }
    .status { font-weight: bold; color: #2563eb; margin: 20px; }
  </style>
</head>
<body>
  <h1>Generating Word Doc Icons...</h1>
  <div class="status" id="status">Starting...</div>
  <div style="display:none;">
    <img id="logo" src="/logo.png" alt="">
  </div>
  <canvas id="canvas" style="display:none;"></canvas>

  <script>
    const sizes = [
      { name: 'favicon-16x16.png', size: 16 },
      { name: 'favicon-32x32.png', size: 32 },
      { name: 'favicon.ico', size: 32 },
      { name: 'apple-touch-icon.png', size: 180 },
      { name: 'android-chrome-192.png', size: 192 },
      { name: 'android-chrome-512.png', size: 512 }
    ];

    async function generate() {
      const statusEl = document.getElementById('status');
      const canvas = document.getElementById('canvas');
      const ctx = canvas.getContext('2d');
      const img = document.getElementById('logo');

      await new Promise((resolve) => {
        if (img.complete) return resolve();
        img.onload = resolve;
      });

      for (const item of sizes) {
        canvas.width = item.size;
        canvas.height = item.size;
        ctx.clearRect(0, 0, item.size, item.size);
        ctx.drawImage(img, 0, 0, item.size, item.size);
        const dataUrl = canvas.toDataURL('image/png');

        statusEl.textContent = 'Saving ' + item.name + '...';
        await fetch('/save', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ filename: item.name, dataUrl })
        });
      }

      statusEl.textContent = 'All icons generated successfully! Closing...';
      setTimeout(() => {
        fetch('/done', { method: 'POST' });
        window.close();
      }, 1000);
    }

    window.onload = generate;
  </script>
</body>
</html>
`;

const server = http.createServer((req, res) => {
  if (req.method === 'GET' && req.url === '/') {
    res.writeHead(200, { 'Content-Type': 'text/html' });
    res.end(htmlPage);
  } else if (req.method === 'GET' && req.url === '/logo.png') {
    const filePath = path.join(PUBLIC_DIR, 'logo.png');
    if (fs.existsSync(filePath)) {
      res.writeHead(200, { 'Content-Type': 'image/png' });
      fs.createReadStream(filePath).pipe(res);
    } else {
      res.writeHead(404);
      res.end();
    }
  } else if (req.method === 'POST' && req.url === '/save') {
    let body = '';
    req.on('data', chunk => { body += chunk; });
    req.on('end', () => {
      const { filename, dataUrl } = JSON.parse(body);
      const base64Data = dataUrl.replace(/^data:image\/png;base64,/, '');
      const filePath = path.join(PUBLIC_DIR, filename);
      fs.writeFileSync(filePath, base64Data, 'base64');
      console.log(`Saved: ${filename}`);
      res.writeHead(200, { 'Content-Type': 'text/plain' });
      res.end('OK');
    });
  } else if (req.method === 'POST' && req.url === '/done') {
    res.writeHead(200, { 'Content-Type': 'text/plain' });
    res.end('OK');
    console.log('Icon generation finished. Shutting down server.');
    process.exit(0);
  } else {
    res.writeHead(404);
    res.end();
  }
});

server.listen(PORT, () => {
  console.log(`Temp server listening on http://localhost:${PORT}`);
  // Launch browser
  const url = `http://localhost:${PORT}`;
  if (process.platform === 'win32') {
    exec(`start ${url}`);
  } else if (process.platform === 'darwin') {
    exec(`open ${url}`);
  } else {
    exec(`xdg-open ${url}`);
  }
});