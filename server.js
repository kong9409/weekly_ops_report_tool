const http = require("http");
const fs = require("fs");
const path = require("path");

const port = process.env.PORT || 3000;
const root = path.resolve(__dirname);
const types = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".md": "text/markdown; charset=utf-8",
};

http.createServer((request, response) => {
  const pathname = decodeURIComponent(request.url.split("?")[0]);
  const relativePath = pathname === "/" ? "index.html" : pathname.replace(/^[/\\]+/, "");
  const filePath = path.resolve(root, relativePath);

  if (filePath !== root && !filePath.startsWith(`${root}${path.sep}`)) {
    response.writeHead(403);
    response.end("Forbidden");
    return;
  }

  fs.readFile(filePath, (error, data) => {
    if (error) {
      response.writeHead(404);
      response.end("Not found");
      return;
    }
    response.writeHead(200, {
      "Content-Type": types[path.extname(filePath).toLowerCase()] || "application/octet-stream",
      "Cache-Control": "no-cache",
    });
    response.end(data);
  });
}).listen(port, "0.0.0.0", () => {
  console.log(`weekly tool on ${port}`);
});
