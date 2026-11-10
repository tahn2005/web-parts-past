const http = require('http'), fs = require('fs'), path = require('path');
const roots = { '/dist/': '/home/user/Manager/link-cards/dist/', '/pnploc/': '/home/user/Manager/link-cards/node_modules/@pnp/spfx-property-controls/lib/loc/', '/': __dirname + '/' };
http.createServer((req, res) => {
  const url = decodeURIComponent(req.url.split('?')[0]);
  const prefix = Object.keys(roots).find((p) => url.startsWith(p));
  const file = path.join(roots[prefix], url.slice(prefix.length) || 'index.html');
  fs.readFile(file, (err, data) => {
    if (err) { res.writeHead(404); return res.end('nf'); }
    res.writeHead(200, { 'Content-Type': file.endsWith('.js') ? 'text/javascript' : 'text/html' });
    res.end(data);
  });
}).listen(4380, () => console.log('listening'));
