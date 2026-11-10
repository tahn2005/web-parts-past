const http = require('http'), fs = require('fs'), path = require('path');
const roots = {
  '/dist-acc/': '/home/user/Manager/accordion-cards/dist/',
  '/dist-hdr/': '/home/user/Manager/custom-width-header/dist/',
  '/dist-lnk/': '/home/user/Manager/link-cards/dist/',
  '/pnploc/': '/home/user/Manager/accordion-cards/node_modules/@pnp/spfx-property-controls/lib/loc/',
  '/': __dirname + '/'
};
http.createServer((req, res) => {
  let url = decodeURIComponent(req.url.split('?')[0]);
  if (url.startsWith('/sites/hr/SiteAssets/')) url = '/logo.png';
  const prefix = Object.keys(roots).find((p) => url.startsWith(p));
  const file = path.join(roots[prefix], url.slice(prefix.length) || 'index.html');
  fs.readFile(file, (err, data) => {
    if (err) { res.writeHead(404); return res.end('nf'); }
    res.writeHead(200, { 'Content-Type': file.endsWith('.js') ? 'text/javascript' : file.endsWith('.html') ? 'text/html' : 'image/png' });
    res.end(data);
  });
}).listen(4379, () => console.log('listening'));
