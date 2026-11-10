// End-to-end tests of the PRODUCTION bundle (release/assets) running in a
// simulated SharePoint page: see README.md in this folder.
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const http = require('http');
const axeSrc = fs.readFileSync(require.resolve('axe-core/axe.min.js'), 'utf8');

// ---- Stage the production build next to the simulator and serve it ----------
const ROOT = path.resolve(__dirname, '..', '..');
const SERVE = path.join(ROOT, 'temp', 'e2e');
function stage() {
  const release = path.join(ROOT, 'release');
  if (!fs.existsSync(path.join(release, 'assets'))) throw new Error('Run "npm run build" first.');
  fs.rmSync(SERVE, { recursive: true, force: true });
  fs.mkdirSync(path.join(SERVE, 'assets'), { recursive: true });
  fs.mkdirSync(path.join(SERVE, 'manifests'), { recursive: true });
  fs.mkdirSync(path.join(SERVE, 'vendor'), { recursive: true });
  ['index.html', 'host.js'].forEach((f) => fs.copyFileSync(path.join(__dirname, f), path.join(SERVE, f)));
  const assets = fs.readdirSync(path.join(release, 'assets'));
  assets.forEach((f) => fs.copyFileSync(path.join(release, 'assets', f), path.join(SERVE, 'assets', f)));
  const manifest = fs.readdirSync(path.join(release, 'manifests')).filter((f) => f.endsWith('.manifest.json'))[0];
  fs.copyFileSync(path.join(release, 'manifests', manifest), path.join(SERVE, 'manifests', 'manifest.json'));
  const bundle = assets.filter((f) => /^resource-cards-web-part_.*\.js$/.test(f))[0];
  const strings = assets.filter((f) => /^ResourceCardsWebPartStrings_.*\.js$/.test(f))[0];
  fs.writeFileSync(path.join(SERVE, 'files.js'), `window.__files=${JSON.stringify({ bundle, strings })};`);
  const nm = path.join(ROOT, 'node_modules');
  [['react/umd/react.production.min.js'], ['react/umd/react.development.js'], ['react-dom/umd/react-dom.production.min.js'],
    ['react-dom/umd/react-dom.development.js'], ['lodash/lodash.min.js']].forEach(([f]) => fs.copyFileSync(path.join(nm, f), path.join(SERVE, 'vendor', path.basename(f))));
}
const siteAssets = new Map();
const TYPES = { '.html': 'text/html', '.js': 'text/javascript', '.json': 'application/json', '.woff2': 'font/woff2', '.txt': 'text/plain' };
function serve() {
  return new Promise((resolve) => {
    const server = http.createServer((req, res) => {
      // In-memory "Site Assets" library: the mock REST client PUTs uploads here.
      const urlPath = decodeURIComponent(new URL(req.url, 'http://x').pathname);
      if (urlPath.startsWith('/sites/hr/SiteAssets/')) {
        if (req.method === 'PUT') {
          const chunks = [];
          req.on('data', (c) => chunks.push(c));
          req.on('end', () => { siteAssets.set(urlPath, Buffer.concat(chunks)); res.writeHead(201); res.end(); });
          return;
        }
        const data = siteAssets.get(urlPath);
        if (!data) { res.writeHead(404); res.end(); return; }
        const ext = path.extname(urlPath).toLowerCase();
        res.writeHead(200, { 'Content-Type': { '.png': 'image/png', '.svg': 'image/svg+xml', '.jpg': 'image/jpeg' }[ext] || 'application/octet-stream' });
        res.end(data);
        return;
      }
      const file = path.join(SERVE, decodeURIComponent(new URL(req.url, 'http://x').pathname));
      if (!file.startsWith(SERVE) || !fs.existsSync(file) || fs.statSync(file).isDirectory()) { res.writeHead(404); res.end(); return; }
      res.writeHead(200, { 'Content-Type': TYPES[path.extname(file)] || 'application/octet-stream' });
      fs.createReadStream(file).pipe(res);
    });
    server.listen(0, '127.0.0.1', () => resolve(server));
  });
}
let BASE;
const results = [];
let page;

function check(name, ok, detail) {
  results.push({ name, ok: !!ok, detail });
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${detail !== undefined ? '  — ' + (typeof detail === 'string' ? detail : JSON.stringify(detail)) : ''}`);
}

async function boot(query = '', viewport = { width: 1400, height: 900 }, extra = {}) {
  if (page) await page.context().close();
  const ctx = await browser.newContext({ viewport, ...extra });
  // Saved properties go in via an init script rather than the URL, which
  // would exceed request-size limits for large card sets.
  const m = /(?:^|&)props=([^&]*)/.exec(query);
  if (m) {
    await ctx.addInitScript((text) => { window.__props = text; }, decodeURIComponent(m[1]));
    query = query.replace(m[0], '');
  }
  page = await ctx.newPage();
  page.on('pageerror', (e) => results.push({ name: 'pageerror', ok: false, detail: e.message }));
  page.on('dialog', async (d) => { check('no dialogs (script injection)', false, d.message()); await d.dismiss(); });
  await page.goto(`${BASE}?${query}`);
  await page.waitForFunction(() => window.__ready, null, { timeout: 15000 });
  const ready = await page.evaluate(() => window.__ready);
  if (ready !== true) throw new Error('boot failed: ' + JSON.stringify(await page.evaluate(() => window.__log)));
  return page;
}
const props = () => page.evaluate(() => JSON.parse(window.host.serialize()));
const logs = () => page.evaluate(() => window.__log);
const titles = () => page.locator('#wp li h2, #wp li h3, #wp li h4, #wp li h5, #wp li h6').allTextContents();
async function axe(label, selector) {
  await page.addScriptTag({ content: axeSrc });
  const v = await page.evaluate(async (sel) => {
    const r = await axe.run({ include: [sel] }, { runOnly: ['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa', 'wcag22aa', 'best-practice'] });
    return r.violations.map((x) => `${x.id} (${x.nodes.length}): ${x.nodes[0].target.join(' ')}`);
  }, selector);
  check(`axe: ${label}`, v.length === 0, v.length ? v : 'no violations');
}

let browser;
let server;
(async () => {
  stage();
  server = await serve();
  BASE = `http://127.0.0.1:${server.address().port}/index.html`;
  browser = await chromium.launch();

  // ---------------------------------------------------------------- view mode
  for (const dev of ['0', '1']) {
    await boot(`dev=${dev}`);
    check(`boots with no console errors/warnings (React ${dev === '1' ? 'development' : 'production'} build)`, (await logs()).length === 0, await logs());
  }
  const manifest = await page.evaluate(() => window.host.manifest);
  check('manifest: not available in full-width sections', manifest.supportsFullBleed === false, manifest.supportsFullBleed);
  check('manifest: toolbox name is "Resource Cards"', manifest.preconfiguredEntries[0].title.default === 'Resource Cards');
  check('manifest: toolbox group is "Resource Cards" and no other product name appears', manifest.preconfiguredEntries[0].group.default === 'Resource Cards' && !/manager.?hub/i.test(JSON.stringify(manifest)));
  check('renders the 3 preconfigured cards', (await titles()).length === 3, await titles());
  check('sample cards have no URL, so no link renders in view mode', (await page.locator('#wp a').count()) === 0);

  const styles = await page.evaluate(() => {
    const h = document.querySelector('#wp li h3'); const p = document.querySelector('#wp li p');
    const s = (e) => getComputedStyle(e);
    return { hSize: s(h).fontSize, hMargin: s(h).margin, hColor: s(h).color, hFont: s(h).fontFamily, pMargin: s(p).margin, pSize: s(p).fontSize };
  });
  check('host page heading/paragraph styles do not leak in', styles.hSize === '14px' && styles.hMargin === '0px 0px 6px' && styles.hColor === 'rgb(15, 28, 54)' && styles.pMargin === '0px 0px 12px' && styles.pSize === '12.48px', styles);

  // The web part is the cards only: nothing of its own around them.
  const shell = await page.evaluate(() => {
    const root = document.querySelector('#wp > div'); const ul = root.querySelector('ul');
    const s = (e) => getComputedStyle(e);
    const wp = document.getElementById('wp').getBoundingClientRect(); const first = ul.querySelector('li').getBoundingClientRect();
    const last = [...ul.querySelectorAll('li')].pop().getBoundingClientRect();
    return {
      rootBg: s(root).backgroundColor, rootBorder: s(root).borderTopWidth + ' ' + s(root).borderLeftWidth, rootPad: s(root).padding, rootShadow: s(root).boxShadow,
      ulBg: s(ul).backgroundColor, ulBorder: s(ul).borderTopWidth, ulPad: s(ul).padding,
      wrappers: root.children.length,
      leftGap: first.left - wp.left, rightGap: wp.right - last.right, topGap: first.top - wp.top
    };
  });
  check('no panel, border, background or padding around the cards', shell.rootBg === 'rgba(0, 0, 0, 0)' && shell.rootBorder === '0px 0px' && shell.rootPad === '0px' && shell.rootShadow === 'none' &&
    shell.ulBg === 'rgba(0, 0, 0, 0)' && shell.ulBorder === '0px' && shell.ulPad === '0px', shell);
  check('cards reach the edges of the web part area (flush with other web parts)', shell.leftGap === 0 && shell.rightGap === 0 && shell.topGap === 0, shell);
  check('each card keeps its own border', (await page.locator('#wp li').first().evaluate((e) => getComputedStyle(e).borderTopWidth + ' ' + getComputedStyle(e).borderTopColor)) === '1px rgba(26, 52, 97, 0.08)');

  const font = await page.evaluate(async () => { await document.fonts.ready; return { r: document.fonts.check("400 12px 'Resource Cards Inter'"), b: document.fonts.check("600 14px 'Resource Cards Inter'"), faces: [...document.fonts].filter((f) => f.family.includes('Resource Cards Inter')).map((f) => f.weight + ':' + f.status) }; });
  check('bundled Inter font (400 + 600) loads from the package', font.r && font.b && font.faces.every((f) => f.endsWith('loaded')), font);
  check('the font style tag is added only once', (await page.locator('style#resource-cards-inter-font').count()) === 1);

  // ---------------------------------------------------------------- edit mode
  await page.evaluate(() => window.host.setMode(2));
  check('entering edit mode shows "Add card" and a toolbar on each card', (await page.getByRole('button', { name: 'Add card' }).count()) === 1 && (await page.locator('#wp [role=group]').count()) === 3);
  check('edit mode shows a "no link" reminder on cards without a URL', (await page.locator('#wp li', { hasText: 'No link yet' }).count()) === 3);
  await axe('edit mode (canvas)', '#wp');

  await page.getByRole('button', { name: 'Add card' }).click();
  check('Add card appends a 4th card', (await titles()).length === 4, await titles());
  check('…and opens the property pane on that card', await page.evaluate(() => window.host.paneOpen) && (await page.locator('#pane .pane-header').textContent()).includes('Editing card 4 of 4'));
  check('…and moves keyboard focus to the new card\'s Edit button', (await page.evaluate(() => document.activeElement.getAttribute('aria-label'))) === 'Edit card: New card');
  check('…and announces it to screen readers', (await page.locator('#wp [role=status]').textContent()).includes('Card added'));
  const selectedCount = () => page.locator('#wp li').evaluateAll((els) => els.filter((e) => /(^|\s)selected_/.test(e.className)).length);
  check('the new card (only) is highlighted as selected', /(^|\s)selected_/.test(await page.locator('#wp li').nth(3).getAttribute('class')) && (await selectedCount()) === 1);

  const field = (label) => page.locator('#pane .field', { has: page.locator(`label:text-is("${label}")`) }).locator('input,textarea');
  await field('Title').fill('Benefits Portal');
  check('typing a title in the pane updates the card live', (await titles())[3] === 'Benefits Portal');
  check('…and stores it at cards[3].title', (await props()).cards[3].title === 'Benefits Portal');
  await field('Title').fill('');
  check('an empty title shows a validation message', (await page.locator('#pane .field', { has: page.locator('label:text-is("Title")') }).locator('.err').textContent()).length > 0);
  await field('Title').fill('Benefits Portal');
  await field('Description').fill('Enroll, compare plans <b>and</b> manage dependents.');
  await field('Link text').fill('Open portal');

  await field('Link URL').fill('javascript:alert(document.cookie)');
  check('a javascript: URL is rejected with a message', (await page.locator('#pane .field', { has: page.locator('label:text-is("Link URL")') }).locator('.err').textContent()).includes('aren\'t allowed'));
  await field('Link URL').fill('https://benefits.example.com/portal?a=1&b=2');
  check('a valid URL clears the message', (await page.locator('#pane .field', { has: page.locator('label:text-is("Link URL")') }).locator('.err').textContent()) === '');
  await page.locator('#pane label', { hasText: 'Open link in a new tab' }).locator('input').check();
  check('new-tab toggle is stored', (await props()).cards[3].openInNewTab === true);

  await page.locator('#pane input[aria-label], #pane').first(); // no-op anchor
  await page.locator('#pane').getByLabel('Search icons').fill('piggy');
  check('icon search filters the grid', (await page.locator('#pane').getByRole('radio').filter({ hasNot: page.locator('xx') }).count()) >= 1);
  await page.locator('#pane label[title="Piggy bank"]').click();
  check('choosing an icon stores it and updates the card', (await props()).cards[3].icon === 'PiggyBank' && (await page.locator('#wp li').nth(3).locator('svg.lucide-piggy-bank').count()) === 1);

  const cardAccent = page.locator('#pane fieldset', { has: page.locator('legend:text-is("Card accent colour")') });
  await cardAccent.locator('label[title="Green (Support & Retain)"]').click();
  check('a per-card accent is stored and applied to that card only', (await props()).cards[3].accentColor === '#059669' &&
    (await page.locator('#wp li').nth(3).evaluate((e) => e.style.getPropertyValue('--rc-accent'))) === '#059669' &&
    (await page.locator('#wp li').nth(0).evaluate((e) => e.style.getPropertyValue('--rc-accent'))) === '#2563eb');

  await page.locator('#pane button', { hasText: 'Move earlier' }).click();
  check('pane "Move earlier" moves the card', (await titles())[2] === 'Benefits Portal', await titles());
  check('…and the pane follows it ("card 3 of 4")', (await page.locator('#pane .pane-header').textContent()).includes('card 3 of 4'));
  await field('Title').fill('Benefits Portal ✓');
  check('after a move, pane fields edit the right card', (await props()).cards[2].title === 'Benefits Portal ✓' && (await props()).cards[3].title === 'Interviewing');

  await page.locator('#pane button', { hasText: 'Done: back to all settings' }).click();
  check('Done returns the pane to the general settings', (await page.locator('#pane h2').allTextContents()).join('|') === 'Cards|Appearance');
  check('settings page lists every card as an edit button', (await page.locator('#pane button[aria-label^="Edit card:"]').count()) === 4);

  const wpAccent = page.locator('#pane fieldset', { has: page.locator('legend:text-is("Accent colour")') });
  await wpAccent.locator('label[title="Indigo (Lead & Develop)"]').click();
  const accents = await page.locator('#wp li').evaluateAll((els) => els.map((e) => e.style.getPropertyValue('--rc-accent')));
  check('web part accent recolours cards that have no override', accents.join() === '#4f46e5,#4f46e5,#059669,#4f46e5', accents);
  await wpAccent.getByLabel('Custom colour (hex)').fill('#fde047');
  const yellow = await page.locator('#wp li').first().evaluate((e) => ({ a: e.style.getPropertyValue('--rc-accent'), t: e.style.getPropertyValue('--rc-text') }));
  check('custom hex accent is applied; too-light colour gets darker link text', (await props()).accentColor === '#fde047' && yellow.a === '#fde047' && yellow.t !== '#fde047', yellow);
  check('…and the pane explains it', (await wpAccent.textContent()).includes('too light'));
  await wpAccent.getByLabel('Custom colour (hex)').fill('#zz');
  check('an invalid hex shows an error and is not stored', (await props()).accentColor === '#fde047' && (await wpAccent.textContent()).includes('Enter a hex colour'));

  check('there is no background-panel setting any more', (await page.locator('#pane', { hasText: 'background panel' }).count()) === 0);
  await page.locator('#pane label', { hasText: 'Font' }).locator('input').uncheck();
  await page.locator('#pane select').selectOption('2');
  const look = await page.evaluate(() => ({ h: document.querySelectorAll('#wp li h2').length, font: getComputedStyle(document.querySelector('#wp li h2')).fontFamily }));
  check('site font and heading level 2 apply', look.h === 4 && !look.font.includes('Resource Cards Inter'), look);
  check('heading level is stored as a number', (await props()).headingLevel === 2);
  await axe('property pane (settings page)', '#pane');

  // Canvas toolbar while the pane is open
  await page.locator('#wp').getByRole('button', { name: 'Delete card: Greenhouse - Our ATS' }).click();
  check('canvas Delete removes the card', (await titles()).join('|') === 'Recruiting Process|Benefits Portal ✓|Interviewing', await titles());
  check('…pane refreshes to the new count', (await page.locator('#pane .pane-label').first().textContent()).startsWith('3 cards'));
  check('…focus moves to the next card', (await page.evaluate(() => document.activeElement.getAttribute('aria-label'))) === 'Edit card: Benefits Portal ✓');
  await page.locator('#wp').getByRole('button', { name: 'Move later: Recruiting Process' }).click();
  check('canvas Move later reorders', (await titles())[1] === 'Recruiting Process');
  await page.locator('#wp').getByRole('button', { name: 'Edit card: Interviewing' }).click();
  check('canvas Edit switches the pane to that card', (await page.locator('#pane .pane-header').textContent()).includes('card 3 of 3'));
  await axe('property pane (card page)', '#pane');
  await page.locator('#pane button', { hasText: 'Delete card' }).click();
  check('pane Delete removes the card and returns to settings', (await titles()).length === 2 && (await page.locator('#pane h2').allTextContents())[0] === 'Cards');

  check('selection highlight follows the card being edited', (await selectedCount()) === 0); // pane Delete cleared it
  await page.locator('#wp').getByRole('button', { name: 'Edit card: Recruiting Process' }).click();
  check('…canvas Edit highlights that card', (await selectedCount()) === 1);
  await page.evaluate(() => window.host.closePane());
  check('closing the pane clears the selection highlight', (await selectedCount()) === 0);

  // ---------------------------------------------------------------- persistence
  const saved = await props();
  check('saved properties are plain JSON with the expected shape', Array.isArray(saved.cards) && saved.cards.every((c) => ['id', 'icon', 'title', 'description', 'linkText', 'url', 'openInNewTab', 'accentColor'].every((k) => k in c)), saved);
  await boot('props=' + encodeURIComponent(JSON.stringify(saved)));
  check('re-opening the saved page (view mode) restores every card in order', (await titles()).join('|') === 'Benefits Portal ✓|Recruiting Process');
  const link = page.locator('#wp a');
  check('only the card with a URL has a link', (await link.count()) === 1);
  check('link has href, target=_blank, rel=noopener noreferrer, SharePoint interception off',
    (await link.getAttribute('href')) === 'https://benefits.example.com/portal?a=1&b=2' && (await link.getAttribute('target')) === '_blank' && (await link.getAttribute('rel')) === 'noopener noreferrer' && (await link.getAttribute('data-interception')) === 'off');
  check('link accessible name = visible text + title + new-tab warning', (await link.getAttribute('aria-label')) === 'Open portal: Benefits Portal ✓ (opens in a new tab)');
  check('HTML typed into a description is shown as text, not markup', (await page.locator('#wp li p').first().textContent()).includes('<b>and</b>') && (await page.locator('#wp li p b').count()) === 0);
  check('read mode has no edit controls', (await page.locator('#wp button').count()) === 0);
  await axe('view mode after editing (yellow accent, site font)', '#wp');

  // Stored javascript: URL (e.g. set before validation, or by script) must not render
  const evil = { ...saved, cards: [{ ...saved.cards[0], url: 'java\tscript:alert(1)', linkText: 'Click' }, { ...saved.cards[0], id: 'x2', url: 'data:text/html,<script>alert(1)</script>' }] };
  await boot('props=' + encodeURIComponent(JSON.stringify(evil)));
  check('stored javascript:/data: URLs never render as links', (await page.locator('#wp a').count()) === 0);

  // ---------------------------------------------------------------- robustness
  const junk = { cards: [null, 'x', 42, { title: 5, icon: 'NoSuchIcon', url: 7 }, { title: 'Ok' }], accentColor: 'red', headingLevel: 9, showPanel: true };
  await boot('dev=1&props=' + encodeURIComponent(JSON.stringify(junk)));
  check('malformed saved properties still render (2 valid cards, defaults applied)', (await page.locator('#wp li').count()) === 2 && (await logs()).length === 0, await logs());
  const fixed = await props();
  check('…and are repaired (accent, heading level, ids, fallback icon)', fixed.accentColor === '#2563eb' && fixed.headingLevel === 3 && !('showPanel' in fixed) && fixed.cards.every((c) => c.id) && fixed.cards[0].icon === 'NoSuchIcon' && (await page.locator('#wp li svg.lucide-file-text').count()) >= 1, fixed);
  // Pages saved by earlier versions used "useHubFont" for the font setting.
  await boot('props=' + encodeURIComponent(JSON.stringify({ useHubFont: false, cards: [{ id: 'o', title: 'Old page card' }] })));
  const migrated = await props();
  check('an earlier saved font setting ("useHubFont") is carried over and renamed', migrated.useInterFont === false && !('useHubFont' in migrated) &&
    !(await page.locator('#wp li h3').evaluate((h) => getComputedStyle(h).fontFamily)).includes('Resource Cards Inter'), migrated);
  const visibleText = async () => (await page.locator('body').innerText()) + (await page.evaluate(() => [...document.querySelectorAll('[aria-label],[title]')].map((e) => (e.getAttribute('aria-label') || '') + (e.getAttribute('title') || '')).join(' ')));
  await page.evaluate(() => window.host.setMode(2));
  await page.locator('#wp').getByRole('button', { name: 'Edit card: Old page card' }).click();
  await page.locator('#pane button', { hasText: 'Done: back to all settings' }).click();
  check('no "Manager Hub" anywhere on the canvas or in the settings pane', !/manager.?hub/i.test(await visibleText()));
  check('font option reads "Resource Cards font (Inter)"', (await page.locator('#pane').textContent()).includes('Font') && (await page.evaluate(() => window.host.lastConfig.pages[0].groups[1].groupFields.filter((f) => f.targetProperty === 'useInterFont')[0].properties.onText)) === 'Resource Cards font (Inter)');

  await boot('props=' + encodeURIComponent(JSON.stringify({})));
  check('a web part with no properties renders nothing in view mode', (await page.locator('#wp').innerHTML()) === '');
  await page.evaluate(() => window.host.setMode(2));
  check('…and shows an empty state with Add card in edit mode', (await page.locator('#wp').textContent()).includes('No cards yet'));
  await axe('empty state', '#wp');

  const xss = { cards: [{ id: 'a', icon: 'Star', title: '<img src=x onerror=alert(1)>', description: '"><svg onload=alert(2)>', linkText: '<script>alert(3)</script>', url: 'https://x.test/"onmouseover="alert(4)', openInNewTab: false, accentColor: '"><style>body{display:none}</style>' }] };
  await boot('props=' + encodeURIComponent(JSON.stringify(xss)));
  await page.locator('#wp a').hover();
  check('script/HTML in every text field is inert', (await page.locator('#wp img, #wp script, #wp style').count()) === 0 && (await page.locator('#wp svg[onload]').count()) === 0);

  // Many cards
  const many = { cards: Array.from({ length: 120 }, (_, i) => ({ id: 'c' + i, icon: 'Star', title: 'Card ' + i, description: 'Description ' + i, linkText: 'View Resource', url: '/sites/hr/' + i, openInNewTab: false, accentColor: '' })) };
  await boot('props=' + encodeURIComponent(JSON.stringify(many)));
  const t = await page.evaluate(async () => {
    const s = performance.now();
    window.host.setMode(2);
    while (document.querySelectorAll('#wp [role=group]').length < 120) await new Promise((r) => requestAnimationFrame(() => r()));
    return performance.now() - s;
  });
  check('120 cards render and switch to edit mode quickly', (await page.locator('#wp li').count()) === 120 && t < 500, `${Math.round(t)} ms`);

  // ---------------------------------------------------------------- section backgrounds
  // SharePoint sections can be white, neutral, soft or strong (theme colour).
  // The web part has no background, so whatever the section has shows through.
  for (const [name, bg] of [['white', '#ffffff'], ['neutral', '#f3f2f1'], ['soft', '#eff6fc'], ['strong (theme colour)', '#0078d4'], ['dark theme', '#1b1a19']]) {
    await boot('props=' + encodeURIComponent(JSON.stringify(saved)));
    await page.evaluate((bg) => { document.getElementById('section').style.backgroundColor = bg; document.getElementById('section').style.padding = '24px'; }, bg);
    const seen = await page.evaluate(() => { const r = document.querySelector('#wp > div').getBoundingClientRect(); const ul = document.querySelector('#wp ul');
      const hit = document.elementFromPoint(r.left + r.width / 2, r.top + 2); return { gapBg: getComputedStyle(ul).backgroundColor, cardBg: getComputedStyle(ul.querySelector('li')).backgroundColor }; });
    check(`${name} section: background shows around the cards; cards keep their own colours`, seen.gapBg === 'rgba(0, 0, 0, 0)' && seen.cardBg === 'rgb(250, 251, 254)', seen);
    if (name.startsWith('strong')) await page.locator('#section').screenshot({ path: path.join(ROOT, 'temp', 'e2e-strong-section.png') });
    await axe(`view mode on a ${name} section`, '#wp');
    await page.evaluate(() => window.host.setMode(2));
    await axe(`edit mode on a ${name} section`, '#wp');
  }
  await page.evaluate(() => window.host.instance.properties.cards.splice(0));
  await page.evaluate(() => window.host.instance.render());
  await axe('empty state on a dark section', '#wp');

  // ---------------------------------------------------------------- layout
  await boot();
  for (const [w, cols] of [[300, 1], [557, 1], [558, 2], [933, 2], [934, 3], [1200, 3]]) {
    await page.evaluate((w) => { document.getElementById('wp').style.width = w + 'px'; }, w);
    const n = await page.evaluate(() => getComputedStyle(document.querySelector('#wp ul')).gridTemplateColumns.split(' ').length);
    check(`web part ${w}px wide → ${cols} column(s)`, n === cols, n);
  }
  await boot('', { width: 320, height: 700 });
  check('no horizontal scrolling at 320px (WCAG 1.4.10 reflow)', await page.evaluate(() => document.documentElement.scrollWidth <= 320));
  await boot('edit=1', { width: 320, height: 700 });
  check('edit toolbar fits at 320px', await page.evaluate(() => document.documentElement.scrollWidth <= 320 && [...document.querySelectorAll('#wp [role=group]')].every((g) => { const r = g.getBoundingClientRect(); const c = g.closest('li').getBoundingClientRect(); return r.left >= c.left && r.right <= c.right; })));
  await boot('', { width: 1400, height: 900 }, { reducedMotion: 'reduce' });
  await page.locator('#wp li').first().hover();
  check('reduced motion: no hover lift', (await page.locator('#wp li').first().evaluate((e) => getComputedStyle(e).transform)) === 'none');
  await boot('', { width: 1400, height: 900 }, { forcedColors: 'active' });
  check('forced colours (High Contrast): card edges stay visible', (await page.locator('#wp li').first().evaluate((e) => getComputedStyle(e).borderTopStyle)) === 'solid');
  await boot();
  await page.evaluate(() => document.documentElement.setAttribute('dir', 'rtl'));
  check('right-to-left: arrows flip', true); // arrows only render with links; covered below
  await boot('props=' + encodeURIComponent(JSON.stringify(saved)));
  await page.evaluate(() => document.documentElement.setAttribute('dir', 'rtl'));
  results.pop();
  check('right-to-left: link arrow is mirrored', (await page.locator('#wp a svg').first().evaluate((e) => getComputedStyle(e).transform)).startsWith('matrix(-1'));

  // ---------------------------------------------------------------- keyboard, lifecycle
  await boot('props=' + encodeURIComponent(JSON.stringify(many)));
  await page.keyboard.press('Tab');
  const first = await page.evaluate(() => ({ tag: document.activeElement.tagName, label: document.activeElement.getAttribute('aria-label'), outline: getComputedStyle(document.activeElement).outlineStyle }));
  check('Tab reaches the first card link with a visible focus ring', first.tag === 'A' && first.label === 'View Resource: Card 0' && first.outline === 'solid', first);
  await page.evaluate(() => window.host.instance.onDispose());
  check('dispose unmounts cleanly', (await page.locator('#wp').innerHTML()) === '' && (await logs()).length === 0);
  await page.evaluate(() => window.host.mount({ cards: [] }));
  check('a new instance can mount after dispose', (await logs()).length === 0);

  await boot('edit=1');
  await page.locator('#wp').getByRole('button', { name: 'Edit card: Interviewing' }).click();
  await page.evaluate(() => window.host.setMode(1));
  check('leaving edit mode closes the pane and removes edit controls', !(await page.evaluate(() => window.host.paneOpen)) && (await page.locator('#wp button').count()) === 0);


  // ---------------------------------------------------------------- custom images
  const LOGO = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAYAAADimHc4AAAB20lEQVR42u2dvVHDQBSErR3l0AHtUAsBZFADZKYat+MGGGgAiJQwgKzR3b29e9/GHvlpv9s9yb+HA0IIoayaSh3o6vH+K5t5Hy+vUyiAjKaXhjFhfCwIYX5ZbfVHmB8LQZgfC0GYHwtBmB8LQZgfC0GYHwtBWBMrsfpjU0ACHBOAAgFQP21riARQQQBAgZp7G/j9+bj6mOunBwC0Nv2vx7vDmEcxvVcYGtn8VscdJgEtDFqewyUNymS+YxqU0XwnCMpqvsvzK/PJO8yh7OZHz8NLERn3AMfr8ai5hPmx81FB3AcAgPoJnJMEUEEAoH4C5yUBVBAAEAAAgAAAAAQAAFRTTx8XbDkvCaCCADBcrHuakwRQQQAYNt49zKcMJ+k8FxWUdQ9wS0HUPMp40k5zzA4n/9/7r5/n2/2r7OZkuwiUeQU6JNBmE25thkv9WX1JbzGl5kdC3DZ/y8vQWiY53n/YflF7Mevtbuy7b27EAAAABAAAIAAAAAEAAAgAAEAAAABqpF//dIwf8K6nn3/0RgKoIACsxgTVqR8S4FxBpKD+6l9NABDqmn9RBQGhnvkX7wFAqGP+pk0YCOXN33wVBITyPqnmwTF/XbvM5DWj/Quz2GrOCIM2QAh1rm/Z9LcB+cthCAAAAABJRU5ErkJggg==', 'base64');
  const SVG = Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" width="40" height="40"><circle cx="20" cy="20" r="18" fill="#7c3aed"/><script>alert("svg")</script></svg>');
  await boot('edit=1');
  const card0 = () => page.locator('#wp li').nth(0);
  await page.locator('#wp').getByRole('button', { name: 'Edit card: Recruiting Process' }).click();
  check('older saved cards (no image fields) default to an icon', (await page.locator('#pane input[type=radio][value=icon]').isChecked()) && (await card0().locator('svg.lucide-workflow').count()) === 1);
  await page.locator('#pane input[type=radio][value=image]').check();
  check('choosing "My own image" swaps the icon picker for the image fields', (await page.locator('#pane legend', { hasText: /^Image$/ }).count()) === 1 &&
    (await page.locator('#pane legend', { hasText: /^Icon$/ }).count()) === 0 && (await page.locator('#pane label', { hasText: 'Image description (alt text)' }).count()) === 1);
  check('until an image is chosen the card keeps its icon, with a note for the author', (await card0().locator('svg.lucide-workflow').count()) === 1 && (await card0().textContent()).includes('No image yet'));
  await axe('property pane (image fields)', '#pane');

  const fileInput = page.locator('#pane input[type=file]');
  check('file chooser only offers image types', (await fileInput.getAttribute('accept')) === 'image/png,image/jpeg,image/gif,image/webp,image/svg+xml');
  check('the hidden file input is out of the tab order (the Upload button opens it)', (await fileInput.getAttribute('tabindex')) === '-1' && (await page.locator('#pane button', { hasText: 'Upload image' }).count()) === 1);

  await page.evaluate(() => { window.host.httpCalls = []; });
  await fileInput.setInputFiles({ name: 'policy.pdf', mimeType: 'application/pdf', buffer: Buffer.from('%PDF-1.4') });
  check('a non-image file is refused with an explanation, nothing uploaded', (await page.locator('#pane [role=alert]').textContent()).includes("isn't a supported image") && (await page.evaluate(() => window.host.httpCalls.length)) === 0);
  await fileInput.setInputFiles({ name: 'huge.png', mimeType: 'image/png', buffer: Buffer.alloc(5 * 1024 * 1024 + 1) });
  check('an image over 5 MB is refused with an explanation, nothing uploaded', (await page.locator('#pane [role=alert]').textContent()).includes('up to 5 MB') && (await page.evaluate(() => window.host.httpCalls.length)) === 0);

  await page.evaluate(() => { window.__uploadFail = 403; });
  await fileInput.setInputFiles({ name: 'logo.png', mimeType: 'image/png', buffer: LOGO });
  await page.locator('#pane [role=alert]').waitFor();
  check('a refused upload (no permission) shows a clear message and changes nothing', (await page.locator('#pane [role=alert]').textContent()).includes("don't have permission") && (await props()).cards[0].imageUrl === '');
  await page.evaluate(() => { window.__uploadFail = 0; window.host.httpCalls = []; });

  await fileInput.setInputFiles({ name: 'Team Logo #1.png', mimeType: 'image/png', buffer: LOGO });
  await page.waitForFunction(() => JSON.parse(window.host.serialize()).cards[0].imageUrl !== '');
  const calls = await page.evaluate(() => window.host.httpCalls.map((c) => `${c.method} ${c.url.replace(location.origin, '')}${c.bodyType ? ' ' + c.bodyType : ''}`));
  check('upload uses SharePoint REST: ensure Site Assets → its folder → add file (never overwrite)', calls.length === 3 &&
    calls[0] === 'POST /sites/hr/_api/web/lists/EnsureSiteAssetsLibrary' &&
    calls[1] === "GET /sites/hr/_api/web/lists(guid'2f1f0b52-0000-4000-8000-000000000001')/RootFolder?$select=ServerRelativeUrl" &&
    /^POST \/sites\/hr\/_api\/web\/GetFolderByServerRelativePath\(decodedurl='%2Fsites%2Fhr%2FSiteAssets'\)\/Files\/AddUsingPath\(decodedurl='resource-card-\d{14}-Team-Logo-1\.png',overwrite=false\) \[object File\]$/.test(calls[2]), calls);
  const stored = (await props()).cards[0];
  check('the uploaded image address is saved on the card', /^\/sites\/hr\/SiteAssets\/resource-card-\d{14}-Team-Logo-1\.png$/.test(stored.imageUrl) && stored.mediaType === 'image', stored.imageUrl);
  check('the pane confirms the upload to screen readers', (await page.locator('#pane [role=status]').last().textContent()).includes('Uploaded Team Logo #1.png'));
  const img = card0().locator('img');
  await img.waitFor();
  await page.waitForFunction(() => { const i = document.querySelector('#wp li img'); return i && i.complete && i.naturalWidth > 0; });
  const box = await img.evaluate((i) => ({ w: i.getBoundingClientRect().width, h: i.getBoundingClientRect().height, alt: i.getAttribute('alt'), loaded: i.naturalWidth }));
  check('the card shows the uploaded image at icon size (32px) in place of the icon', box.w === 32 && box.h === 32 && box.loaded === 96 && (await card0().locator('svg.lucide-workflow').count()) === 0, box);
  check('with no alt text the image is decorative (alt="")', box.alt === '');
  check('the pane shows a preview and the file name', (await page.locator('#pane img').count()) === 1 && (await page.locator('#pane').textContent()).includes('Team-Logo-1.png'));
  const titleTop = async () => card0().locator('h3').evaluate((h) => h.getBoundingClientRect().top - h.closest('li').getBoundingClientRect().top);
  check('a small image keeps the card layout identical to an icon card', (await titleTop()) === (await page.locator('#wp li').nth(1).locator('h3').evaluate((h) => h.getBoundingClientRect().top - h.closest('li').getBoundingClientRect().top)));

  await page.locator('#pane select').selectOption('large');
  check('image size Large shows it at 64px', (await img.evaluate((i) => i.getBoundingClientRect().width)) === 64);
  await page.locator('#pane select').selectOption('medium');
  check('image size Medium shows it at 48px', (await img.evaluate((i) => i.getBoundingClientRect().width)) === 48);
  await field('Image description (alt text)').fill('Talent Acquisition team logo');
  check('alt text typed in the pane is applied to the image', (await img.getAttribute('alt')) === 'Talent Acquisition team logo');

  // Paste a link instead
  const linkBox = page.locator('#pane').getByLabel('Or paste an image link');
  await linkBox.fill('javascript:alert(1)');
  // React 18 renders the validation message just after the input event.
  await page.locator('#pane', { hasText: "aren't allowed" }).waitFor({ timeout: 3000 }).catch(() => {});
  check('pasting a javascript: link is refused and the saved image is kept', (await page.locator('#pane').textContent()).includes("aren't allowed") && (await props()).cards[0].imageUrl === stored.imageUrl);
  await linkBox.fill('/sites/hr/SiteAssets/missing.png');
  await page.waitForFunction(() => document.querySelector('#wp li').textContent.includes("couldn't be loaded"));
  check('a broken image link falls back to the icon, with a note for the author', (await card0().locator('img').count()) === 0 && (await card0().locator('svg.lucide-workflow').count()) === 1);
  check('…and the pane preview says it could not be loaded', (await page.locator('#pane').textContent()).includes("This image couldn't be loaded"));

  // SVG upload (scripts inside an SVG never run when shown with <img>)
  await fileInput.setInputFiles({ name: 'badge.svg', mimeType: 'image/svg+xml', buffer: SVG });
  await page.waitForFunction(() => /badge\.svg$/.test(JSON.parse(window.host.serialize()).cards[0].imageUrl));
  await page.waitForFunction(() => { const i = document.querySelector('#wp li img'); return i && i.complete && i.naturalWidth > 0; });
  check('SVG images upload and display (embedded script stays inert)', (await card0().locator('img').count()) === 1);

  // A second card with a PNG for the view-mode checks
  await page.locator('#wp').getByRole('button', { name: 'Edit card: Interviewing' }).click();
  await page.locator('#pane input[type=radio][value=image]').check();
  await page.locator('#pane input[type=file]').setInputFiles({ name: 'logo.png', mimeType: 'image/png', buffer: LOGO });
  await page.waitForFunction(() => JSON.parse(window.host.serialize()).cards[2].imageUrl !== '');
  check('each upload gets its own file (no overwriting)', (await props()).cards[2].imageUrl !== (await props()).cards[0].imageUrl);
  await page.locator('#pane button', { hasText: 'Remove image' }).click();
  check('Remove image clears it and the icon returns', (await props()).cards[2].imageUrl === '' && (await page.locator('#wp li').nth(2).locator('svg.lucide-messages-square').count()) === 1);
  await page.locator('#pane input[type=file]').setInputFiles({ name: 'logo.png', mimeType: 'image/png', buffer: LOGO });
  await page.waitForFunction(() => JSON.parse(window.host.serialize()).cards[2].imageUrl !== '');
  await page.locator('#wp').getByRole('button', { name: 'Edit card: Greenhouse - Our ATS' }).click();
  await page.locator('#pane input[type=radio][value=image]').check();
  await page.locator('#pane input[type=radio][value=icon]').check();
  check('switching a card back to Icon shows the icon again', (await page.locator('#wp li').nth(1).locator('svg.lucide-monitor').count()) === 1 && (await page.locator('#pane legend', { hasText: /^Icon$/ }).count()) === 1);
  await axe('edit mode with images', '#wp');

  const withImages = await props();
  await boot('props=' + encodeURIComponent(JSON.stringify(withImages)));
  await page.waitForFunction(() => [...document.querySelectorAll('#wp li img')].every((i) => i.complete && i.naturalWidth > 0));
  const view = await page.locator('#wp li img').evaluateAll((els) => els.map((i) => ({ alt: i.getAttribute('alt'), w: i.getBoundingClientRect().width })));
  check('after saving and reopening, images show for readers with their alt text and size', view.length === 2 && view[0].alt === 'Talent Acquisition team logo' && view[0].w === 48 && view[1].alt === '' && view[1].w === 32, view);
  check('readers never see author notes', !(await page.locator('#wp').textContent()).includes('No image yet') && !(await page.locator('#wp').textContent()).includes("couldn't"));
  await page.locator('#wp').screenshot({ path: path.join(ROOT, 'temp', 'e2e-images.png') });
  await axe('view mode with images', '#wp');

  const broken = JSON.parse(JSON.stringify(withImages));
  broken.cards[0].imageUrl = '/sites/hr/SiteAssets/deleted.png';
  broken.cards[1] = { ...broken.cards[1], mediaType: 'image', imageUrl: 'data:image/png;base64,AAAA' };
  await boot('props=' + encodeURIComponent(JSON.stringify(broken)));
  await page.waitForTimeout(400);
  check('readers see the icon, not a broken image, when an image was deleted', (await page.locator('#wp li').nth(0).locator('img').count()) === 0 && (await page.locator('#wp li').nth(0).locator('svg').count()) >= 1);
  check('a stored data: image address is never rendered', (await page.locator('#wp li').nth(1).locator('img').count()) === 0);

  await browser.close();
  server.close();
  const failed = results.filter((r) => !r.ok);
  console.log(`\n${results.length - failed.length}/${results.length} passed`);
  if (failed.length) { console.log('FAILURES:', JSON.stringify(failed, null, 1)); process.exit(1); }
})().catch(async (e) => { console.error('CRASH', e); try { console.log(await logs()); } catch {} process.exit(2); });

