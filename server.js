// Paprastas HTTPS serveris add-in kūrimui
// Outlook reikalauja HTTPS net lokaliai

const https = require('https');
const fs    = require('fs');
const path  = require('path');

// Jei neturi sertifikatų, naudok office-addin-dev-certs (žr. README)
let sslOptions;
try {
  sslOptions = {
    key:  fs.readFileSync(path.join(__dirname, 'certs', 'server.key')),
    cert: fs.readFileSync(path.join(__dirname, 'certs', 'server.crt'))
  };
} catch {
  console.error('❌ SSL sertifikatai nerasti!');
  console.error('Paleisk: npx office-addin-dev-certs install');
  console.error('Tada nukopijuok ~/.office-addin-dev-certs/ → ./certs/');
  process.exit(1);
}

const MIME = {
  '.html': 'text/html',
  '.js':   'application/javascript',
  '.css':  'text/css',
  '.png':  'image/png',
  '.ico':  'image/x-icon',
  '.xml':  'application/xml',
};

const server = https.createServer(sslOptions, (req, res) => {
  let filePath = req.url.split('?')[0];

  if (filePath === '/' || filePath === '/taskpane.html') {
    filePath = '/src/taskpane.html';
  }

  try {
    filePath = decodeURIComponent(filePath);
  } catch (e) {
    res.writeHead(400); res.end('Bad Request'); return;
  }

  const root = __dirname;
  const fullPath = path.resolve(root, '.' + filePath);

  if (fullPath !== root && !fullPath.startsWith(root + path.sep)) {
    res.writeHead(403);
    res.end('Forbidden');
    return;
  }

  // Strict whitelist for allowed files
  const isAllowed = 
    filePath === '/src/taskpane.html' || 
    filePath === '/manifest.xml' || 
    filePath.startsWith('/assets/');

  if (!isAllowed) {
    res.writeHead(403);
    res.end('Forbidden');
    return;
  }

  fs.readFile(fullPath, (err, data) => {
    if (err) {
      res.writeHead(404);
      res.end('Not found: ' + filePath);
      return;
    }
    const ext  = path.extname(fullPath);
    const mime = MIME[ext] || 'application/octet-stream';
    res.writeHead(200, { 'Content-Type': mime });
    res.end(data);
  });
});

const PORT = 3000;
server.listen(PORT, () => {
  console.log(`✅ Add-in serveris veikia: https://localhost:${PORT}`);
  console.log(`📋 Manifest: https://localhost:${PORT}/manifest.xml`);
  console.log(`🔧 Taskpane: https://localhost:${PORT}/src/taskpane.html`);
});
