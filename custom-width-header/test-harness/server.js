const http = require('http'), fs = require('fs'), path = require('path');
const roots = { '/dist/': '/home/user/Manager/custom-width-header/dist/', '/pnploc/': '/home/user/Manager/custom-width-header/node_modules/@pnp/spfx-property-controls/lib/loc/', '/': __dirname + '/' };
http.createServer((req, res) => {
  let url = decodeURIComponent(req.url.split('?')[0]);
  if (url.startsWith('/sites/hr/SiteAssets/')) url = '/logo.png'; // stands in for the uploaded file
  const prefix = Object.keys(roots).find((p) => url.startsWith(p));
  const file = path.join(roots[prefix], url.slice(prefix.length) || 'index.html');
  fs.readFile(file, (err, data) => {
    if (err) { res.writeHead(404); return res.end('nf'); }
    const type = file.endsWith('.js') ? 'text/javascript' : file.endsWith('.html') ? 'text/html' : file.endsWith('.png') ? 'image/png' : 'application/octet-stream';
    res.writeHead(200, { 'Content-Type': type }); res.end(data);
  });
}).listen(4378, () => console.log('listening'));
