// Paprastas HTTPS serveris add-in kūrimui.
// Outlook reikalauja HTTPS net lokaliai.

const https = require("https");
const fs = require("fs");
const path = require("path");

const CONTENT_SECURITY_POLICY = [
  "default-src 'self'",
  "base-uri 'self'",
  "object-src 'none'",
  "script-src 'self' https://appsforoffice.microsoft.com https://alcdn.msauth.net",
  "style-src 'self'",
  "img-src 'self' data:",
  "connect-src 'self' https://graph.microsoft.com https://login.microsoftonline.com",
  "frame-src https://login.microsoftonline.com",
  "form-action 'none'",
  "upgrade-insecure-requests",
].join("; ");

const ALLOWED_FILES = new Set([
  "/manifest.xml",
  "/manifest.local.xml",
  "/src/taskpane.html",
  "/src/taskpane.css",
  "/src/taskpane.generated.css",
  "/src/taskpane.js",
]);

let sslOptions;
try {
  sslOptions = {
    key: fs.readFileSync(path.join(__dirname, "certs", "server.key")),
    cert: fs.readFileSync(path.join(__dirname, "certs", "server.crt")),
  };
} catch {
  console.error("❌ SSL sertifikatai nerasti!");
  console.error("Paleisk: npx office-addin-dev-certs install");
  console.error("Tada nukopijuok ~/.office-addin-dev-certs/ → ./certs/");
  process.exit(1);
}

const MIME = {
  ".css": "text/css",
  ".html": "text/html",
  ".ico": "image/x-icon",
  ".js": "application/javascript",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".xml": "application/xml",
};

const server = https.createServer(sslOptions, (req, res) => {
  if (req.method !== "GET" && req.method !== "HEAD") {
    writeError(res, 405, "Method Not Allowed", { Allow: "GET, HEAD" });
    return;
  }

  let filePath = req.url.split("?")[0];
  if (filePath === "/" || filePath === "/taskpane.html") {
    filePath = "/src/taskpane.html";
  }

  try {
    filePath = decodeURIComponent(filePath);
  } catch {
    writeError(res, 400, "Bad Request");
    return;
  }

  const root = __dirname;
  const fullPath = path.resolve(root, "." + filePath);

  if (fullPath !== root && !fullPath.startsWith(root + path.sep)) {
    writeError(res, 403, "Forbidden");
    return;
  }

  const isAllowed = ALLOWED_FILES.has(filePath) || filePath.startsWith("/assets/");
  if (!isAllowed) {
    writeError(res, 403, "Forbidden");
    return;
  }

  fs.readFile(fullPath, (err, data) => {
    if (err) {
      writeError(res, 404, "Not found: " + filePath);
      return;
    }

    res.writeHead(200, getSecurityHeaders(filePath, fullPath));
    if (req.method === "HEAD") {
      res.end();
      return;
    }

    res.end(data);
  });
});

function getSecurityHeaders(filePath, fullPath) {
  const mime = MIME[path.extname(fullPath)] || "application/octet-stream";
  const isUtf8 =
    mime.startsWith("text/") ||
    mime === "application/javascript" ||
    mime === "application/xml";

  return {
    "Cache-Control":
      filePath.endsWith(".html") || filePath.endsWith(".xml")
        ? "no-store"
        : "public, max-age=300",
    "Content-Security-Policy": CONTENT_SECURITY_POLICY,
    "Content-Type": mime + (isUtf8 ? "; charset=utf-8" : ""),
    "Permissions-Policy":
      "accelerometer=(), camera=(), geolocation=(), gyroscope=(), microphone=(), payment=(), usb=()",
    "Referrer-Policy": "no-referrer",
    "X-Content-Type-Options": "nosniff",
  };
}

function writeError(res, statusCode, message, extraHeaders = {}) {
  res.writeHead(statusCode, {
    "Cache-Control": "no-store",
    "Content-Type": "text/plain; charset=utf-8",
    "Referrer-Policy": "no-referrer",
    "X-Content-Type-Options": "nosniff",
    ...extraHeaders,
  });
  res.end(message);
}

const PORT = Number(process.env.PORT || 3000);
server.listen(PORT, () => {
  console.log(`✅ Add-in serveris veikia: https://localhost:${PORT}`);
  console.log(`📋 Local manifest: https://localhost:${PORT}/manifest.local.xml`);
  console.log(`🔧 Taskpane: https://localhost:${PORT}/src/taskpane.html`);
});
