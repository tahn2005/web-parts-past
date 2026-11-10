/*
 * Link Cards — end-to-end tests against the production bundle (link-cards/dist)
 * running in a stand-in SharePoint host.
 */
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const AXE = fs.readFileSync(require.resolve('../harness/node_modules/axe-core/axe.min.js'), 'utf8');
const manifest = require('/home/user/Manager/link-cards/src/webparts/linkCards/LinkCardsWebPart.manifest.json');
const DEFAULTS = manifest.preconfiguredEntries[0].properties;
const SPPKG = '/home/user/Manager/link-cards/sharepoint/solution/link-cards.sppkg';
const BASE = 'http://localhost:4380/';

const results = [];
/** Names that must not appear in the package (other products and web parts), stored encoded so they don't appear in this file either. */
const OTHER_NAMES = new RegExp(Buffer.from('bWFuYWdlci1odWJ8bWFuYWdlciBodWJ8bWFuYWdlcmlhbHxhY2NvcmRpb258Y3VzdG9tLXdpZHRofHBhZ2UtaGVhZGVyfGNsaWVudC1zaWRlLXNvbHV0aW9u', 'base64').toString(), 'i');
/** The two React dev-mode notices PnP's controls can emit; anything else is a failure. */
const KNOWN_PNP_NOTICES = [/`key` is not a prop/, /Can't perform a React state update on an unmounted component/];
const thirdParty = new Set();
function check(name, ok, detail) {
  results.push({ name, ok: !!ok });
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${!ok && detail !== undefined ? '  -> ' + JSON.stringify(detail) : ''}`);
}

const card = (id, extra = {}) => ({ ...DEFAULTS.cards[0], id, ...extra });

async function open(browser, { props = DEFAULTS, mode = 2, viewport = { width: 1440, height: 900 }, media, two = false } = {}) {
  const page = await browser.newPage({ viewport });
  if (media) await page.emulateMedia(media);
  const requests = [];
  const consoleErrors = [];
  page.on('request', (r) => requests.push(r.url()));
  page.on('console', (m) => {
    if (m.type() !== 'error' && m.type() !== 'warning') return;
    if (/was used but not registered/.test(m.text())) return; // Fluent icon font comes from SharePoint pages
    if (KNOWN_PNP_NOTICES.some((r) => r.test(m.text()))) return void thirdParty.add(m.text().split('\n')[0].slice(0, 150));
    consoleErrors.push(m.text());
  });
  page.on('pageerror', (e) => consoleErrors.push('pageerror: ' + e.message));
  page.on('dialog', (d) => (page.__dialogAnswer === false ? d.dismiss() : d.accept()));
  await page.goto(BASE + 'index.html');
  await page.evaluate(
    async ({ props, mode, two }) => {
      window.wp1 = await mountWebPart({ mountId: 'wp1', properties: JSON.parse(JSON.stringify(props)), displayMode: mode, instanceId: 'l1' });
      if (two) window.wp2 = await mountWebPart({ mountId: 'wp2', properties: JSON.parse(JSON.stringify(props)), displayMode: mode, instanceId: 'l2' });
    },
    { props, mode, two }
  );
  await page.waitForTimeout(120);
  return { page, requests, consoleErrors };
}

async function axeRun(page, label, rules) {
  await page.evaluate(AXE);
  const r = await page.evaluate((rules) => axe.run(document, { runOnly: rules || ['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa', 'wcag22aa', 'best-practice'] }), rules);
  const v = r.violations.map((x) => `${x.id}: ${x.nodes.map((n) => n.target.join(' ')).slice(0, 2).join(' | ')}`);
  check(`axe: ${label}`, v.length === 0, v);
  const review = r.incomplete.filter((x) => x.id === 'color-contrast').reduce((n, x) => n + x.nodes.length, 0);
  check(`axe: ${label} — nothing left for manual contrast review`, review === 0, review);
}

/** WCAG contrast of an element's (text/icon/outline) colour against the card behind it. */
function contrastIn(page, selector, prop) {
  return page.$eval(
    selector,
    (el, prop) => {
      const parse = (c) => c.match(/[\d.]+/g).slice(0, 3).map(Number);
      const lum = (c) => {
        const ch = (v) => ((v /= 255) <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4));
        return 0.2126 * ch(c[0]) + 0.7152 * ch(c[1]) + 0.0722 * ch(c[2]);
      };
      const ratio = (a, b) => {
        const [x, y] = [lum(a), lum(b)];
        return (Math.max(x, y) + 0.05) / (Math.min(x, y) + 0.05);
      };
      const cardEl = el.closest('[class*=card]');
      const fg = parse(getComputedStyle(el)[prop]);
      const bg = parse(getComputedStyle(cardEl).backgroundColor);
      return { ratio: Math.round(ratio(fg, bg) * 100) / 100, againstWhite: Math.round(ratio(fg, [255, 255, 255]) * 100) / 100 };
    },
    prop
  );
}

const props = (page) => page.evaluate(() => JSON.parse(JSON.stringify(wp1.properties)));
const setText = async (page, p, v) => {
  const sel = `#pane .field[data-path="${p}"] input`;
  await page.fill(sel, v);
  await page.dispatchEvent(sel, 'change');
  await page.waitForTimeout(30);
};
const paneHeader = (page) => page.$eval('#pane .pane-header', (e) => e.textContent).catch(() => '');
const paneButton = (page, text) => page.click(`#pane button:text-is("${text}")`);

(async () => {
  const server = require('child_process').spawn(process.execPath, [path.join(__dirname, 'server.js')]);
  process.on('exit', () => server.kill());
  server.stderr.on('data', (d) => {
    console.error('SERVER: ' + d);
    process.exit(3);
  });
  await new Promise((r) => server.stdout.once('data', r));
  const browser = await chromium.launch();

  // ------------------------------------------------------------ 0. the package
  {
    const xml = execFileSync('unzip', ['-p', SPPKG, '*.xml']).toString();
    const app = execFileSync('unzip', ['-p', SPPKG, 'AppManifest.xml']).toString();
    const files = execFileSync('unzip', ['-Z1', SPPKG]).toString();
    const all = execFileSync('unzip', ['-p', SPPKG], { maxBuffer: 100 * 1024 * 1024 }).toString('latin1');
    check('naming: package file is link-cards.sppkg', path.basename(SPPKG) === 'link-cards.sppkg');
    check('naming: shows as "Link Cards" (app, feature and toolbox)', /<Title>Link Cards<\/Title>/.test(app) && /Name="Link Cards"/.test(app) && /Title="Link Cards"/.test(xml) && /&quot;title&quot;:\{&quot;default&quot;:&quot;Link Cards&quot;/.test(xml));
    check('naming: no other product or web part names anywhere inside (SharePoint API names aside)', !OTHER_NAMES.test(all.replace(/displayGroupsAsAccordion/g, '')));
    check('naming: every bundled file is named link-cards (no numbered chunks)', files.split('\n').filter((f) => /^ClientSideAssets\/.+\.js$/.test(f)).every((f) => /^ClientSideAssets\/(link-cards-web-part_|chunk\.link-cards-property-pane_|LinkCardsWebPartStrings_|PropertyControlStrings_)/.test(f)));
    check('package: NOT available in full-width sections', /&quot;supportsFullBleed&quot;:false/.test(xml) && !/&quot;supportsFullBleed&quot;:true/.test(xml));
    check('package: SharePoint pages only', /&quot;supportedHosts&quot;:\[&quot;SharePointWebPart&quot;\]/.test(xml));
    check('package: tenant-wide deployment', /SkipFeatureDeployment="true"/.test(xml));
    check('package: no custom script required', !/requiresCustomScript&quot;:true/i.test(xml));
    check('package: version 1.0.1.0', /Version="1\.0\.1\.0"/.test(app));
    check('package: starts with one card', (xml.match(/default-link-1/g) || []).length >= 1);
  }

  // ------------------------------------------------------------ 1. dropped onto a page (edit mode)
  {
    const { page, requests, consoleErrors } = await open(browser);
    check('drop-in: one card plus the Add tile', (await page.$$('#wp1 li')).length === 2);
    check('drop-in: card shows "Link title" with the external-link icon', /Link title/.test(await page.textContent('#wp1 li')) && !!(await page.$('#wp1 li svg.lucide-external-link')));
    check('drop-in: no address yet -> not a link, author is told', !(await page.$('#wp1 a')) && /No link address yet/.test(await page.textContent('#wp1')));
    check('drop-in: card toolbar in edit mode', (await page.$$('#wp1 [role=group] button')).length === 4);
    check('drop-in: no pane code downloaded until the pane opens', !requests.some((u) => /chunk\./.test(u)));
    await axeRun(page, 'edit mode, default card');
    check('drop-in: no console errors and no unknown SharePoint APIs', consoleErrors.length === 0 && (await page.evaluate(() => __missing.length)) === 0, consoleErrors);
    await page.close();
  }

  // ------------------------------------------------------------ 2. readers
  {
    const sample = {
      ...DEFAULTS,
      palette: 'amber',
      cards: [
        card('a', { text: 'Offboarding SOP', url: 'https://contoso.sharepoint.com/sites/hr/Shared%20Documents/sop.docx' }),
        card('b', { text: 'Exit interview form', url: '/sites/hr/SitePages/Exit.aspx', openInNewTab: false, iconName: 'ClipboardList' }),
        card('c', { text: 'Orientation deck coming soon', url: '' })
      ]
    };
    const { page, consoleErrors } = await open(browser, { props: sample, mode: 1 });
    const links = await page.$$eval('#wp1 a', (a) => a.map((x) => ({ href: x.getAttribute('href'), target: x.target, rel: x.rel, name: x.textContent })));
    check('reader: cards with addresses are links to them', links.length === 2 && links[0].href.endsWith('sop.docx') && links[1].href === '/sites/hr/SitePages/Exit.aspx', links);
    check('reader: new-tab links are safe and say so to screen readers', links[0].target === '_blank' && links[0].rel === 'noopener noreferrer' && links[0].name === 'Offboarding SOP (opens in a new tab)', links[0]);
    check('reader: same-tab links are plain links', links[1].target === '' && links[1].name === 'Exit interview form', links[1]);
    check('reader: a card without an address is plain text, not a fake link', (await page.$$('#wp1 li')).length === 3 && (await page.$eval('#wp1 li:nth-child(3) [class*=card]', (c) => c.tagName)) === 'DIV');
    check('reader: no edit toolbar, Add tile or author hints', !(await page.$('#wp1 [role=group]')) && !(await page.$('#wp1 [data-focus-key=add]')) && !/No link address/.test(await page.textContent('#wp1')));
    check('reader: cards are a list', (await page.$eval('#wp1 ul', (u) => u.children.length)) === 3);
    await axeRun(page, 'reader with links and a plain card');

    // Same look as the React version
    const s = await page.evaluate(() => {
      const q = (sel) => getComputedStyle(document.querySelector(sel));
      return { pad: q('#wp1 a').padding, radius: q('#wp1 a').borderRadius, gap: q('#wp1 a').columnGap, icon: q('#wp1 a svg').width, size: q('#wp1 a [class*=text]').fontSize, weight: q('#wp1 a [class*=text]').fontWeight, bg: q('#wp1 a').backgroundColor, iconColor: q('#wp1 a svg').color, deco: q('#wp1 a').textDecorationLine, gridGap: q('#wp1 ul').rowGap };
    });
    check('look: React sizes (12px padding, 10px radius, 12px gap, 16px icon, 13.6px/500 text)', s.pad === '12px' && s.radius === '10px' && s.gap === '12px' && s.icon === '16px' && s.size === '13.6px' && s.weight === '500' && s.gridGap === '12px', s);
    check('look: React colours (card #fafbfe, amber icon) and no underline', s.bg === 'rgb(250, 251, 254)' && s.deco === 'none', s);
    await page.hover('#wp1 li:nth-child(1) a');
    await page.waitForTimeout(250);
    const hover = await page.$eval('#wp1 li:nth-child(1) a', (a) => ({ bg: getComputedStyle(a).backgroundColor, border: getComputedStyle(a).borderTopColor }));
    check('look: hover tints the card and border like the React version (amber)', hover.bg === 'rgb(255, 251, 235)' && hover.border === 'rgb(253, 230, 138)', hover);
    await axeRun(page, 'reader while hovering a card', ['wcag2aa']);
    await page.screenshot({ path: path.join(__dirname, 'reader.png'), clip: { x: 0, y: 0, width: 1120, height: 200 } });
    check('reader: no console errors', consoleErrors.length === 0, consoleErrors);
    await page.close();
  }

  // ------------------------------------------------------------ 3. authoring
  {
    const { page, requests, consoleErrors } = await open(browser);
    await page.click('#wp1 button[aria-label="Edit card"]');
    await page.waitForSelector('#pane .pane-header');
    check('pane: selecting a card opens its settings', (await paneHeader(page)) === 'Card 1 of 1');

    {
      // Every collapsed group must be openable: SharePoint only expands groups on accordion pages.
      const reachable = await page.evaluate(() => {
        const cfg = __pane.wp.getPropertyPaneConfiguration();
        return cfg.pages.every((pg) => !pg.groups.some((g) => g.isCollapsed) || pg.displayGroupsAsAccordion === true);
      });
      check('pane: collapsed groups (Colors, Font, Accessibility…) can be opened in SharePoint', reachable && !(await page.$('#pane section[data-unreachable="true"]')));
    }
    check('pane: pane code loads on demand', requests.some((u) => /chunk\.link-cards-property-pane/.test(u)));
    check('pane: selected card is highlighted and its edit button pressed', !!(await page.$('#wp1 li[class*=selected]')) && (await page.getAttribute('#wp1 button[aria-label="Edit card"]', 'aria-pressed')) === 'true');

    await setText(page, 'cards[0].text', 'Offboarding SOP');
    check('pane: link text updates the card', /Offboarding SOP/.test(await page.textContent('#wp1 li')));
    await setText(page, 'cards[0].url', 'javascript:alert(1)');
    check('pane: unsafe address rejected with a message', /Enter a full address/.test(await page.textContent('#pane .field[data-path="cards[0].url"] .error')) && (await props(page)).cards[0].url === '');
    await setText(page, 'cards[0].url', 'https://contoso.sharepoint.com/sites/hr/sop.docx');
    check('pane: address makes it a link', (await page.getAttribute('#wp1 a', 'href')) === 'https://contoso.sharepoint.com/sites/hr/sop.docx');
    await page.click('#wp1 a');
    check('pane: links do not navigate while editing (selects the card)', page.url() === BASE + 'index.html');
    await page.uncheck('#pane .field[data-path="cards[0].openInNewTab"] input');
    check('pane: new-tab setting', (await page.getAttribute('#wp1 a', 'target')) === null);
    await page.check('#pane .field[data-path="cards[0].openInNewTab"] input');

    await page.fill('#pane .custom input', 'file');
    await page.click('#pane .custom button[aria-label="File text"]');
    check('pane: icon picker changes the icon', (await props(page)).cards[0].iconName === 'FileText' && !!(await page.$('#wp1 li svg.lucide-file-text')));

    await page.check('#pane .field[data-path="cards[0].iconColorMode"] input[value=custom]');
    check('pane: custom icon colour starts from the current one', (await props(page)).cards[0].iconColor === '#2563eb' && !!(await page.$('#pane .field[data-path="cards[0].iconColor"] .custom *')));
    await page.evaluate(() => __paneChange('cards[0].iconColor', '#7c3aed', 'custom'));
    check('pane: icon colour applied', (await page.$eval('#wp1 li svg', (s) => getComputedStyle(s).color)) === 'rgb(124, 58, 237)');

    await paneButton(page, 'Duplicate card');
    check('arrange: duplicate selects the copy', (await paneHeader(page)) === 'Card 2 of 2' && (await page.$$('#wp1 li a')).length === 2);
    await setText(page, 'cards[1].text', 'Copy');
    await paneButton(page, 'Move card earlier');
    check('arrange: move earlier', (await paneHeader(page)) === 'Card 1 of 2' && /^Copy/.test(await page.textContent('#wp1 li:nth-child(1)')));
    page.__dialogAnswer = false;
    await paneButton(page, 'Delete card');
    check('arrange: cancelling the delete keeps the card', (await page.$$('#wp1 li a')).length === 2);
    page.__dialogAnswer = true;
    await paneButton(page, 'Delete card');
    check('arrange: delete removes it and returns to the settings', (await page.$$('#wp1 li a')).length === 1 && /Link Cards settings/.test(await paneHeader(page)));

    check('section: cards listed with an Add button', /1\. Offboarding SOP/.test(await page.textContent('#pane')) && !!(await page.$('#pane button:text-is("Add a link card")')));
    await page.selectOption('#pane .field[data-path="palette"] select', 'custom');
    check('section: custom colour starts from the previous palette and shows a picker', (await props(page)).iconColor === '#2563eb' && !!(await page.$('#pane .field[data-path="iconColor"] .custom *')));
    await page.selectOption('#pane .field[data-path="palette"] select', 'green');
    check('section: preset hides the picker', !(await page.$('#pane .field[data-path="iconColor"]')));
    await page.evaluate(() => __paneChange('textColor', '#bbbbbb', 'custom'));
    await page.evaluate(() => wp1.context.propertyPane.refresh());
    check('section: a low-contrast text colour is corrected and explained', /Card text: #bbbbbb → #/.test(await page.textContent('#pane section[data-group="Accessibility"]')));
    const t = await contrastIn(page, '#wp1 li [class*=text]', 'color');
    check('section: corrected text passes 4.5:1', t.ratio >= 4.5, t);
    await page.evaluate(() => __paneChange('textColor', '#0f1c36', 'custom'));
    await page.selectOption('#pane .field[data-path="columns"] select', '1');
    check('section: columns setting', (await page.$eval('#wp1 ul', (u) => getComputedStyle(u).gridTemplateColumns.split(' ').length)) === 1);
    await page.uncheck('#pane .field[data-path="useInterFont"] input');
    check('section: font can fall back to SharePoint\'s', /Segoe UI/.test(await page.$eval('#wp1 > div', (d) => getComputedStyle(d).fontFamily.split(',')[0])));

    await paneButton(page, 'Add a link card');
    check('section: Add opens the new card', (await paneHeader(page)) === 'Card 2 of 2');
    await paneButton(page, 'Back to Link Cards settings');
    check('pane: Back returns to the settings', /Link Cards settings/.test(await paneHeader(page)));
    await page.click('#wp1 li:nth-child(2) [class*=card]');
    check('canvas: clicking a card selects it', (await paneHeader(page)) === 'Card 2 of 2');
    await page.evaluate(() => wp1.context.propertyPane.close());
    check('pane: closing clears the selection', !(await page.$('#wp1 li[class*=selected]')));

    const saved = await props(page);
    check('saving: no helper values stored', !('paneAction' in saved), Object.keys(saved));
    fs.writeFileSync(path.join(__dirname, 'authored.json'), JSON.stringify(saved, null, 2));
    check('pane: no errors and no unknown SharePoint APIs', consoleErrors.length === 0 && (await page.evaluate(() => __missing.length + __errors.length)) === 0, consoleErrors.concat(await page.evaluate(() => __missing.concat(__errors))));
    await page.close();
  }

  // ------------------------------------------------------------ 4. saved page as a reader
  {
    const saved = JSON.parse(fs.readFileSync(path.join(__dirname, 'authored.json'), 'utf8'));
    const { page, consoleErrors } = await open(browser, { props: saved, mode: 1 });
    check('round trip: authored link renders as a link', (await page.$$('#wp1 li a')).length === 1 && (await page.textContent('#wp1 a')) === 'Offboarding SOP (opens in a new tab)');
    check('round trip: the added card with text but no address shows as plain text', (await page.$$('#wp1 li')).length === 2 && (await page.$eval('#wp1 li:nth-child(2) [class*=card]', (c) => c.tagName)) === 'DIV');
    await axeRun(page, 'saved page, reader');
    check('round trip: no console errors', consoleErrors.length === 0, consoleErrors);
    await page.close();
  }

  // ------------------------------------------------------------ 5. canvas toolbar and focus
  {
    const three = { ...DEFAULTS, cards: ['A', 'B', 'C'].map((x) => card('id' + x, { text: x, url: '/' + x })) };
    const { page } = await open(browser, { props: three });
    const order = () => page.$$eval('#wp1 li a', (a) => a.map((x) => x.textContent.replace(' (opens in a new tab)', '')).join(''));
    const active = () => page.evaluate(() => document.activeElement.dataset.focusKey);
    check('toolbar: first can\'t move earlier, last can\'t move later', (await page.$eval('[data-focus-key="earlier-idA"]', (b) => b.disabled)) && (await page.$eval('[data-focus-key="later-idC"]', (b) => b.disabled)));
    await page.focus('[data-focus-key="later-idA"]');
    await page.keyboard.press('Enter');
    await page.waitForTimeout(60);
    check('toolbar: move with the keyboard; focus stays on the button', (await order()) === 'BAC' && (await active()) === 'later-idA');
    await page.focus('#wp1 [aria-label="Delete card"]');
    await page.keyboard.press('Enter');
    await page.waitForTimeout(60);
    check('toolbar: delete moves focus to the next card', (await order()) === 'AC' && (await active()) === 'edit-idA');
    for (let i = 0; i < 2; i++) {
      await page.focus('#wp1 [aria-label="Delete card"]');
      await page.keyboard.press('Enter');
      await page.waitForTimeout(60);
    }
    check('toolbar: deleting the last card focuses Add', (await page.$$('#wp1 li a')).length === 0 && (await active()) === 'add');
    await page.keyboard.press('Enter');
    await page.waitForSelector('#pane .pane-header');
    check('toolbar: Add with the keyboard adds and opens a card', (await paneHeader(page)) === 'Card 1 of 1');
    await page.close();
  }

  // ------------------------------------------------------------ 6. robustness and security
  {
    for (const [label, bad] of [
      ['empty property bag', {}],
      ['null and partial cards', { cards: [null, { id: 'p', text: 'Partial' }, 'junk', { id: 'q', url: 42 }] }],
      ['nonsense values', { palette: 'nope', iconColor: 'x', textColor: 'blue-ish', cardBackground: '?', columns: '9', cards: [card('z', { iconName: 'Nope', iconColorMode: 'custom', iconColor: 'zzz', text: 'Still fine', url: '/ok' })] }]
    ]) {
      const r = await open(browser, { props: bad, mode: 1 });
      const errs = await r.page.evaluate(() => __errors);
      check(`robustness: ${label} renders without errors`, errs.length === 0 && r.consoleErrors.length === 0, errs.concat(r.consoleErrors));
      await r.page.close();
    }
    const empty = await open(browser, { props: { ...DEFAULTS, cards: [] }, mode: 1 });
    check('robustness: no cards -> nothing shown to readers (no empty list)', !(await empty.page.$('#wp1 ul')));
    await empty.page.close();

    const evil = {
      ...DEFAULTS,
      cards: [
        card('e1', { text: '<img src=x onerror="window.__pwned=1">', url: 'javascript:window.__pwned=1' }),
        card('e2', { text: 'Protocol-relative', url: '//evil.example.com' }),
        card('e3', { text: 'Data', url: 'data:text/html,<script>alert(1)</script>' }),
        card('e4', { text: '', url: ' JAVASCRIPT:alert(1)' }),
        card('e5', { text: 'Mail HR', url: 'mailto:hr@contoso.com' }),
        card('e6', { text: '', url: 'https://contoso.sharepoint.com/sites/hr' })
      ]
    };
    const e = await open(browser, { props: evil, mode: 1 });
    const hrefs = await e.page.$$eval('#wp1 a', (a) => a.map((x) => x.getAttribute('href')));
    check('security: only safe addresses become links', hrefs.join() === 'mailto:hr@contoso.com,https://contoso.sharepoint.com/sites/hr', hrefs);
    check('security: markup in text is shown as text, never run', !(await e.page.evaluate(() => window.__pwned)) && /<img/.test(await e.page.textContent('#wp1')));
    check('security: an unsafe address with no text is hidden, not shown as text', !/JAVASCRIPT/i.test(await e.page.textContent('#wp1')));
    check('a link with no text is named by its address', (await e.page.textContent('#wp1 li:last-child a')).startsWith('https://contoso.sharepoint.com/sites/hr'));
    await axeRun(e.page, 'odd content');
    await e.page.close();
  }

  // ------------------------------------------------------------ 7. layout
  {
    const many = { ...DEFAULTS, cards: Array.from({ length: 6 }, (_, i) => card('m' + i, { text: 'Resource link ' + (i + 1), url: '/p' + i })) };
    const { page } = await open(browser, { props: many, mode: 1 });
    const cols = () => page.$eval('#wp1 ul', (u) => getComputedStyle(u).gridTemplateColumns.split(' ').length);
    const width = async (w) => {
      await page.$eval('#wp1', (el, w) => (el.style.width = w + 'px'), w);
      await page.waitForTimeout(60);
    };
    for (const [w, n] of [[1000, 2], [540, 2], [539, 1], [320, 1]]) {
      await width(w);
      check(`layout (automatic): ${w}px wide -> ${n} column(s)`, (await cols()) === n, await cols());
    }
    await page.evaluate(() => __registry && 0);
    for (const [setting, w, n] of [['3', 760, 3], ['3', 759, 2], ['3', 539, 1], ['2', 1000, 2], ['1', 1000, 1]]) {
      await page.evaluate((s) => { wp1.properties.columns = s; wp1.render(); }, setting);
      await width(w);
      check(`layout (up to ${setting}): ${w}px wide -> ${n} column(s)`, (await cols()) === n, await cols());
    }
    await page.$eval('#wp1', (el) => (el.style.width = ''));
    const s = await page.evaluate(() => {
      const q = (sel) => getComputedStyle(document.querySelector(sel));
      return { ulPad: q('#wp1 ul').paddingLeft, ulStyle: q('#wp1 ul').listStyleType, liMargin: q('#wp1 li').marginBottom, aColor: q('#wp1 a [class*=text]').color, aDeco: q('#wp1 a').textDecorationLine, aWeight: q('#wp1 a [class*=text]').fontWeight };
    });
    check('isolation: SharePoint-style globals do not leak in', s.ulPad === '0px' && s.ulStyle === 'none' && s.liMargin === '0px' && s.aColor === 'rgb(15, 28, 54)' && s.aDeco === 'none' && s.aWeight === '500', s);
    await page.close();

    const long = { ...DEFAULTS, cards: [card('l', { text: 'A-very-long-unbroken-link-title-that-must-wrap-instead-of-pushing-the-page-sideways-on-small-phones', url: '/x' })] };
    for (const mode of [1, 2]) {
      const m = await open(browser, { props: long, mode, viewport: { width: 320, height: 700 } });
      const overflow = await m.page.evaluate(() => document.documentElement.scrollWidth - document.documentElement.clientWidth);
      check(`reflow: nothing scrolls sideways at 320px (${mode === 1 ? 'reading' : 'editing'})`, overflow <= 0, overflow);
      await m.page.close();
    }
  }

  // ------------------------------------------------------------ 8. keyboard, contrast, preferences
  {
    const mixed = { ...DEFAULTS, cards: [card('k1', { text: 'One', url: '/1' }), card('k2', { text: 'Plain', url: '' }), card('k3', { text: 'Three', url: '/3' })] };
    const { page } = await open(browser, { props: mixed, mode: 1 });
    await page.keyboard.press('Tab');
    const first = await page.evaluate(() => document.activeElement.textContent);
    await page.keyboard.press('Tab');
    const second = await page.evaluate(() => document.activeElement.textContent);
    await page.waitForTimeout(250); // the tint fades in over 150ms, as on hover
    check('keyboard: Tab visits each link once and skips plain-text cards', first.startsWith('One') && second.startsWith('Three'), [first, second]);
    const ring = await page.evaluate(() => {
      const s = getComputedStyle(document.activeElement);
      return { style: s.outlineStyle, width: s.outlineWidth, bg: s.backgroundColor };
    });
    const focusContrast = await contrastIn(page, '#wp1 li:nth-child(3) a', 'outlineColor');
    check('keyboard: visible 2px focus ring that meets 3:1 on the page and the card', ring.style === 'solid' && ring.width === '2px' && focusContrast.ratio >= 3 && focusContrast.againstWhite >= 3, { ring, focusContrast });
    check('keyboard: focus also shows the hover tint', ring.bg !== 'rgb(250, 251, 254)', ring);
    const box = await page.$eval('#wp1 a', (a) => a.getBoundingClientRect().height);
    check('target size: each link is at least 44px tall', box >= 44, box);
    await page.close();

    for (const palette of ['blue', 'indigo', 'green', 'amber']) {
      const p = await open(browser, { props: { ...mixed, palette }, mode: 1 });
      const icon = await contrastIn(p.page, '#wp1 li:nth-child(1) svg', 'color');
      await p.page.hover('#wp1 li:nth-child(1) a');
      await p.page.waitForTimeout(250);
      const iconHover = await contrastIn(p.page, '#wp1 li:nth-child(1) svg', 'color');
      const text = await contrastIn(p.page, '#wp1 li:nth-child(1) [class*=text]', 'color');
      check(`${palette}: icon ≥3:1 at rest and on hover, text ≥4.5:1`, icon.ratio >= 3 && iconHover.ratio >= 3 && text.ratio >= 4.5, { icon, iconHover, text });
      await axeRun(p.page, `${palette} while hovering`, ['wcag2aa']);
      await p.page.close();
    }
    const harsh = { ...mixed, palette: 'custom', iconColor: '#fde047', textColor: '#cbd5e1', cardBackground: '#ffffff', cards: mixed.cards.concat(card('k4', { text: 'Own colour', url: '/4', iconColorMode: 'custom', iconColor: '#e0f2fe' })) };
    const h = await open(browser, { props: harsh, mode: 1 });
    for (const n of [1, 4]) {
      const icon = await contrastIn(h.page, `#wp1 li:nth-child(${n}) svg`, 'color');
      const text = await contrastIn(h.page, `#wp1 li:nth-child(${n}) [class*=text]`, 'color');
      check(`pale custom colours are corrected (card ${n}): icon ≥3:1, text ≥4.5:1`, icon.ratio >= 3 && text.ratio >= 4.5, { icon, text });
    }
    await axeRun(h.page, 'pale custom colours');
    await h.page.close();

    const rm = await open(browser, { props: mixed, mode: 1, media: { reducedMotion: 'reduce' } });
    check('reduced motion: no transitions', (await rm.page.$eval('#wp1 a', (a) => getComputedStyle(a).transitionDuration)) === '0s');
    await rm.page.close();
    const hc = await open(browser, { props: mixed, mode: 1, media: { forcedColors: 'active' } });
    await axeRun(hc.page, 'Windows High Contrast');
    await hc.page.close();
    const two = await open(browser, { props: mixed, mode: 2, two: true });
    const ids = await two.page.$$eval('[id]', (e) => e.map((x) => x.id));
    check('two Link Cards on a page: ids unique, edit controls clean', new Set(ids).size === ids.length);
    await axeRun(two.page, 'two Link Cards in edit mode');
    await two.page.close();
  }

  await browser.close();
  console.log('\nThird-party (PnP) dev-mode notices:');
  thirdParty.forEach((t) => console.log('  - ' + t));
  const failed = results.filter((r) => !r.ok);
  console.log(`\n${results.length - failed.length}/${results.length} checks passed`);
  process.exit(failed.length ? 1 : 0);
})().catch((e) => {
  console.error(e);
  process.exit(2);
});
