const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const AXE = fs.readFileSync(require.resolve('../harness/node_modules/axe-core/axe.min.js'), 'utf8');
const manifest = require('/home/user/Manager/accordion-cards/src/webparts/resourceAccordion/ResourceAccordionWebPart.manifest.json');
const DEFAULTS = manifest.preconfiguredEntries[0].properties;
const URL_BASE = 'http://localhost:4377/';

const results = [];
/** The two React dev-mode notices PnP's file picker emits; anything else is a failure. */
const KNOWN_PNP_NOTICES = [/`key` is not a prop/, /Can't perform a React state update on an unmounted component/];
const thirdParty = new Set();
function check(name, ok, detail) {
  results.push({ name, ok: !!ok, detail });
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${!ok && detail !== undefined ? '  -> ' + JSON.stringify(detail) : ''}`);
}

async function open(browser, { props = DEFAULTS, mode = 2, viewport = { width: 1440, height: 1000 }, two = false, media } = {}) {
  const page = await browser.newPage({ viewport });
  if (media) await page.emulateMedia(media);
  const requests = [];
  const consoleErrors = [];
  page.on('request', (r) => requests.push(r.url()));
  page.on('response', (r) => { if (r.status() === 404) console.log('404: ' + r.url()); });
  page.on('console', (m) => {
    if (m.type() !== 'error' && m.type() !== 'warning') return;
    if (/was used but not registered/.test(m.text())) return; // Fluent icon font is registered by SharePoint pages
    const where = m.location() && m.location().url;
    if (KNOWN_PNP_NOTICES.some((r) => r.test(m.text()))) {
      thirdParty.add(m.text().split('\n')[0].slice(0, 160));
      return;
    }
    consoleErrors.push(m.text());
  });
  page.on('pageerror', (e) => consoleErrors.push('pageerror: ' + e.message));
  page.on('dialog', (d) => (page.__dialogAnswer === false ? d.dismiss() : d.accept()));
  await page.goto(URL_BASE + 'index.html');
  await page.evaluate(
    async ({ props, mode, two }) => {
      window.wp1 = await mountWebPart({ mountId: 'wp1', properties: JSON.parse(JSON.stringify(props)), displayMode: mode, instanceId: 'a1' });
      if (two) window.wp2 = await mountWebPart({ mountId: 'wp2', properties: JSON.parse(JSON.stringify(props)), displayMode: mode, instanceId: 'b2' });
    },
    { props, mode, two }
  );
  await page.waitForTimeout(150);
  return { page, requests, consoleErrors };
}

const axeRun = async (page, label, rules) => {
  await page.evaluate(AXE);
  const r = await page.evaluate(
    (rules) =>
      axe.run('#wp1, #wp2', {
        runOnly: rules || ['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa', 'wcag22aa', 'best-practice']
      }),
    rules
  );
  const v = r.violations.map((x) => `${x.id}: ${x.nodes.map((n) => n.target.join(' ')).slice(0, 3).join(' | ')}`);
  check(`axe: ${label}`, v.length === 0, v);
};

const props = (page) => page.evaluate(() => JSON.parse(JSON.stringify(wp1.properties)));
const paneHeader = (page) => page.$eval('#pane .pane-header', (e) => e.textContent).catch(() => '');
const setText = async (page, pathName, value) => {
  const sel = `#pane .field[data-path="${pathName}"] input, #pane .field[data-path="${pathName}"] textarea`;
  await page.fill(sel, value);
  await page.dispatchEvent(sel, 'change');
  await page.waitForTimeout(30);
};
const paneButton = (page, text) => page.click(`#pane button:text-is("${text}")`);

(async () => {
  const server = require('child_process').spawn(process.execPath, [path.join(__dirname, 'server.js')]);
  process.on('exit', () => server.kill());
  server.stderr.on('data', (d) => { console.error('SERVER: ' + d); process.exit(3); });
  await new Promise((r) => server.stdout.once('data', r));
  const browser = await chromium.launch();

  // ------------------------------------------------------------------ 0. the package itself
  {
    const sppkg = '/home/user/Manager/accordion-cards/sharepoint/solution/accordion-cards.sppkg';
    const list = require('child_process').execFileSync('unzip', ['-Z1', sppkg]).toString();
    const xml = require('child_process').execFileSync('unzip', ['-p', sppkg, '*.xml']).toString();
    check('package: tenant-wide deployment (SkipFeatureDeployment)', /SkipFeatureDeployment="true"/.test(xml));
    check('package: SharePoint pages only (no Teams / app-page hosts)', /&quot;supportedHosts&quot;:\[&quot;SharePointWebPart&quot;\]/.test(xml));
    check('package: assets are inside the package (no external CDN needed)', /ClientSideAssets\/resource-accordion-web-part_/.test(list) && /ClientSideAssets\/chunk\.resource-accordion-property-pane_/.test(list));
    check('package: one-card default is the shipped default', /Resource title/.test(xml) && (xml.match(/default-card-1/g) || []).length >= 1 && !/sample-recruiting/.test(xml));
    check('package: no custom script required', !/RequiresCustomScript&quot;:true|requiresCustomScript&quot;:true/.test(xml));
    check('package: NOT available in full-width sections', /&quot;supportsFullBleed&quot;:false/.test(xml) && !/&quot;supportsFullBleed&quot;:true/.test(xml));
    const app = require('child_process').execFileSync('unzip', ['-p', sppkg, 'AppManifest.xml']).toString();

    {
      // Compare with the package that was originally deployed and worked: same identity and configuration,
      // only the web part name, version and intended behaviour settings may differ.
      const fsx = require('fs'), px = require('path');
      const ORIG = '/tmp/claude-0/-home-user-Manager/0cbf0501-6f58-5de0-b625-f92d54f271a2/scratchpad/pkgdiff/acc-old';
      const norm = (dir) => {
        const out = [];
        const walk = (d) => fsx.readdirSync(d, { withFileTypes: true }).forEach((e) => (e.isDirectory() ? walk(px.join(d, e.name)) : out.push(px.relative(dir, px.join(d, e.name)).replace(/_[0-9a-f]{20,32}\./, '.'))));
        walk(dir);
        return out.sort().join('\n');
      };
      const files = require('child_process').execFileSync('unzip', ['-Z1', sppkg]).toString();
      const newFiles = files.split('\n').filter((f) => f && !f.endsWith('/')).map((f) => f.replace(/_[0-9a-f]{20,32}\./, '.')).sort().join('\n');
      check('original config: same files in the package as the original', newFiles === norm(ORIG));
      const oApp = fsx.readFileSync(px.join(ORIG, 'AppManifest.xml'), 'utf8');
      const attr = (s, a) => (new RegExp(a + '="([^"]*)"').exec(s) || [])[1];
      check('original config: same app name, title, product ID and settings', ['Name', 'ProductID', 'SharePointMinVersion', 'IsClientSideSolution', 'SkipFeatureDeployment', 'IsDomainIsolated'].every((a) => attr(oApp, a) === attr(app, a)) && /<Title>([^<]*)<\/Title>/.exec(oApp)[1] === /<Title>([^<]*)<\/Title>/.exec(app)[1], ['Name', 'ProductID'].map((a) => [attr(oApp, a), attr(app, a)]));
      const featFile = fsx.readdirSync(ORIG).find((f) => /^feature_.*\.xml$/.test(f));
      const oFeat = fsx.readFileSync(px.join(ORIG, featFile), 'utf8');
      const nFeat = require('child_process').execFileSync('unzip', ['-p', sppkg, featFile]).toString();
      check('original config: same feature (id, title, description)', ['Id', 'Title', 'Description', 'Scope'].every((a) => attr(oFeat, a) === attr(nFeat, a)));
      const oWp = fsx.readFileSync(px.join(ORIG, 'bf03090c-f1c1-4238-b349-4d12f39cd67a/WebPart_b13052d3-2957-4800-83d5-cf5963855d13.xml'), 'utf8').replace(/&quot;/g, '"');
      const nWp = require('child_process').execFileSync('unzip', ['-p', sppkg, 'bf03090c-f1c1-4238-b349-4d12f39cd67a/WebPart_b13052d3-2957-4800-83d5-cf5963855d13.xml']).toString().replace(/&quot;/g, '"');
      const field = (s, k) => (new RegExp('"' + k + '":"([^"]*)"').exec(s) || [])[1];
      check('original config: same component id, alias, bundle and strings module', ['id', 'alias', 'entryModuleId', 'componentType'].every((k) => field(oWp, k) === field(nWp, k)) && /"ResourceAccordionWebPartStrings"|"PageHeaderWebPartStrings"/.test(oWp) === /"ResourceAccordionWebPartStrings"|"PageHeaderWebPartStrings"/.test(nWp), ['alias', 'entryModuleId'].map((k) => [field(oWp, k), field(nWp, k)]));
      check('web part name: "Accordion Cards" in the toolbox (was "Resource Accordion")', /<ClientSideComponent Name="Accordion Cards"/.test(nWp) && /"title":\{"default":"Accordion Cards"\}/.test(nWp));
      check('version: raised to 1.0.4.0 (above every earlier build) so the overwrite is an upgrade', /Version="1\.0\.4\.0"/.test(app) && /"version":"1\.0\.4"/.test(nWp));
    }
  }

  // ------------------------------------------------------------------ 1. drop-in, edit mode
  {
    const { page, requests, consoleErrors } = await open(browser);
    const cards = await page.$$('#wp1 li');
    check('drop-in: exactly one card plus the Add tile', cards.length === 2, cards.length);
    check('drop-in: card shows default title', (await page.textContent('#wp1 h3')) === 'Resource title');
    check('drop-in: count badge reads 1 / "1 resource"', (await page.textContent('#wp1 [class*=badge]')) === '11 resource');
    check('drop-in: section is open while editing', (await page.getAttribute('#wp1 [aria-expanded]', 'aria-expanded')) === 'true');
    check('drop-in: card toolbar present in edit mode', (await page.$$('#wp1 [role=group] button')).length === 4);
    check('drop-in: no pane code downloaded before the pane opens', !requests.some((u) => /chunk\./.test(u)), requests.filter((u) => /chunk/.test(u)));
    check('drop-in: card without a link shows author hint, not a fake link', !(await page.$('#wp1 a')) && /No link yet/.test(await page.textContent('#wp1')));
    await axeRun(page, 'edit mode, default drop-in');
    const missing = await page.evaluate(() => __missing);
    check('drop-in: no unknown SharePoint APIs used', missing.length === 0, missing);
    check('drop-in: no console errors/React warnings', consoleErrors.length === 0, consoleErrors);
    await page.close();
  }

  // ------------------------------------------------------------------ 2. reader view
  {
    const { page, requests, consoleErrors } = await open(browser, { mode: 1 });
    const trigger = '#wp1 button[aria-expanded]';
    check('reader: starts collapsed', (await page.getAttribute(trigger, 'aria-expanded')) === 'false');
    check('reader: collapsed panel is hidden', await page.$eval('#wp1 [role=region]', (r) => r.hidden && getComputedStyle(r).display === 'none'));
    check('reader: no edit toolbar or Add tile', !(await page.$('#wp1 [role=group]')) && !(await page.$('#wp1 [data-focus-key=add]')));
    await page.click(trigger);
    check('reader: click expands', (await page.getAttribute(trigger, 'aria-expanded')) === 'true' && (await page.isVisible('#wp1 h3')));
    check('reader: aria-controls points at the region', await page.$eval(trigger, (b) => !!document.getElementById(b.getAttribute('aria-controls'))));
    check('reader: region labelled by the title', await page.$eval('#wp1 [role=region]', (r) => document.getElementById(r.getAttribute('aria-labelledby')).textContent === 'Section title'));
    check('reader: no "No link yet" hint shown to readers', !/No link yet/.test(await page.textContent('#wp1')));
    check('reader: downloads only the main bundle (no pane chunks)', !requests.some((u) => /chunk\./.test(u)));
    await axeRun(page, 'reader, expanded');
    check('reader: no console errors', consoleErrors.length === 0, consoleErrors);
    await page.close();
  }

  // ------------------------------------------------------------------ 3. author flows through the property pane
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
    check('pane: pane code is lazy-loaded on open', requests.some((u) => /chunk\.resource-accordion-property-pane/.test(u)));
    check('pane: selected card is highlighted', !!(await page.$('#wp1 [class*=selected]')));
    check('pane: edit button reports pressed state', (await page.getAttribute('#wp1 button[aria-label="Edit card"]', 'aria-pressed')) === 'true');

    await setText(page, 'cards[0].title', 'Recruiting Process');
    await setText(page, 'cards[0].description', 'Learn the hiring process.');
    check('pane: title/description update the canvas', (await page.textContent('#wp1 h3')) === 'Recruiting Process' && /Learn the hiring process\./.test(await page.textContent('#wp1')));

    await setText(page, 'cards[0].linkUrl', 'javascript:alert(1)');
    check('pane: javascript: link rejected with a message', /Enter a full address/.test(await page.textContent('#pane .field[data-path="cards[0].linkUrl"] .error')) && (await props(page)).cards[0].linkUrl === '');
    await setText(page, 'cards[0].linkUrl', 'https://contoso.sharepoint.com/sites/hr/SitePages/Recruiting.aspx');
    await setText(page, 'cards[0].linkText', 'Open guide');
    await page.check('#pane .field[data-path="cards[0].openInNewTab"] input');
    const a = await page.$eval('#wp1 a', (a) => ({ href: a.getAttribute('href'), target: a.target, rel: a.rel, text: a.textContent }));
    check('pane: link, text and new-tab applied', a.href.endsWith('Recruiting.aspx') && a.target === '_blank' && a.rel === 'noopener noreferrer', a);
    check('pane: accessible link name starts with visible text', a.text === 'Open guide: Recruiting Process (opens in a new tab)', a.text);
    await page.click('#wp1 a');
    check('pane: links do not navigate away while editing', page.url() === URL_BASE + 'index.html');

    // icon picker
    await page.fill('#pane .custom input[type=text], #pane .custom input[role=searchbox], #pane .custom input', 'rock');
    await page.waitForTimeout(50);
    const visibleIcons = await page.$$eval('#pane .custom button[aria-pressed]', (b) => b.map((x) => x.getAttribute('aria-label')));
    check('icon picker: search filters icons', visibleIcons.length === 1 && visibleIcons[0] === 'Rocket', visibleIcons);
    await page.click('#pane .custom button[aria-label="Rocket"]');
    check('icon picker: choosing an icon updates the card', (await props(page)).cards[0].iconName === 'Rocket' && !!(await page.$('#wp1 li svg.lucide-rocket')));
    check('icon picker: chosen icon is marked pressed', (await page.getAttribute('#pane .custom button[aria-label="Rocket"]', 'aria-pressed')) === 'true');

    // image: switch, PnP picker renders, link + upload + failure
    await page.check('#pane .field[data-path="cards[0].iconSource"] input[value=image]');
    await page.waitForTimeout(100);
    check('image: PnP file picker renders in the pane', /Choose or upload an image/.test(await page.textContent('#pane')));
    const pickerFieldInfo = await page.evaluate(() => {
      const f = __pane.config.pages[0].groups.flatMap((g) => g.groupFields).find((x) => x.targetProperty === 'imagePickerResult');
      return f ? Object.keys(f.properties) : null;
    });
    check('image: picker field exposes onSave/onPropertyChange', pickerFieldInfo && pickerFieldInfo.includes('onSave') && pickerFieldInfo.includes('onPropertyChange'), pickerFieldInfo);

    // Simulates exactly what PnP's host does on save: onSave, write result to targetProperty, onPropertyChange.
    const pnpSave = (result) =>
      page.evaluate(async (result) => {
        const f = __pane.config.pages[0].groups.flatMap((g) => g.groupFields).find((x) => x.targetProperty === 'imagePickerResult');
        if (result.upload) {
          const name = result.name;
          const bytes = await (await fetch('/logo.png')).arrayBuffer();
          result = { fileName: name, fileNameWithoutExtension: 'x', fileAbsoluteUrl: undefined, downloadFileContent: async () => new File([bytes], name, { type: 'image/png' }) };
        }
        f.properties.onSave(result);
        lodash.set(wp1.properties, 'imagePickerResult', result);
        f.properties.onPropertyChange('imagePickerResult', undefined, result);
        await new Promise((r) => setTimeout(r, 150));
      }, result);

    await pnpSave({ fileName: 'logo.png', fileNameWithoutExtension: 'logo', fileAbsoluteUrl: 'http://localhost:4377/logo.png' });
    let p = await props(page);
    check('image (site/link/stock): URL stored, scratch value removed', p.cards[0].imageUrl === 'http://localhost:4377/logo.png' && !('imagePickerResult' in p), p.cards[0].imageUrl);
    check('image: card renders the image, decorative', (await page.getAttribute('#wp1 li img', 'alt')) === '' && (await page.$eval('#wp1 li img', (i) => i.closest('[aria-hidden=true]') !== null)));
    await page.waitForFunction(() => document.querySelector('#wp1 li img').complete);
    check('image: image actually loads', await page.$eval('#wp1 li img', (i) => i.naturalWidth === 40));

    await page.evaluate(() => (__rest.calls = []));
    await pnpSave({ upload: true, name: 'My Logo!.png' });
    p = await props(page);
    const calls = await page.evaluate(() => __rest.calls.filter((c) => !/FilePickerTabOptions/.test(c.url)).map((c) => `${c.method} ${c.url.replace('https://contoso.sharepoint.com/sites/hr/_api', '')} ${c.bodyBytes || ''}`.trim()));
    check('upload: REST sequence ensure library -> root folder -> folder -> file', calls.length === 4 && /EnsureSiteAssetsLibrary/.test(calls[0]) && /RootFolder/.test(calls[1]) && /folders\/add/.test(calls[2]) && /Files\/add/.test(calls[3]), calls);
    check('upload: file bytes sent', / \d+$/.test(calls[3]) && Number(calls[3].split(' ').pop()) > 50, calls[3]);
    check('upload: unique, sanitised file name in Site Assets/AccordionCards', /^\/sites\/hr\/SiteAssets\/AccordionCards\/\d+-My-Logo-\.png$/.test(p.cards[0].imageUrl), p.cards[0].imageUrl);
    await page.evaluate(() => (__rest.fail = true));
    await pnpSave({ upload: true, name: 'second.png' });
    check('upload failure: author sees the reason, previous image kept', /couldn't be uploaded: Uploading: Access denied\./.test(await page.textContent('#pane')) && (await props(page)).cards[0].imageUrl === p.cards[0].imageUrl);
    await page.evaluate(() => (__rest.fail = false));
    await page.evaluate(() => (__rest.calls = []));
    await pnpSave({ upload: true, name: 'payload.html' });
    check('upload: non-image files are refused before anything is uploaded',
      /couldn't be uploaded: only \.png/.test(await page.textContent('#pane')) && (await page.evaluate(() => __rest.calls.filter((c) => /Files\/add/.test(c.url)).length)) === 0);
    await setText(page, 'cards[0].imageUrl', 'javascript:alert(1)');
    check('image URL field rejects javascript:', /Enter an image address/.test(await page.textContent('#pane .field[data-path="cards[0].imageUrl"] .error')));

    // card colour
    await page.check('#pane .field[data-path="cards[0].accentMode"] input[value=custom]');
    p = await props(page);
    check('card color: custom starts from the section accent', p.cards[0].accentColor === '#2563eb', p.cards[0].accentColor);
    check('card color: PnP color picker renders', !!(await page.$('#pane .field[data-path="cards[0].accentColor"] .custom *')));
    await page.evaluate(() => __paneChange('cards[0].accentColor', '#f59e0b', 'custom'));
    const linkColor = await page.$eval('#wp1 a', (a) => getComputedStyle(a).color);
    const iconColor = await page.$eval('#wp1 li [class^="cardIcon_"]', (a) => getComputedStyle(a).color);
    check('card color: icon uses exact accent, link darkened to pass AA', iconColor === 'rgb(245, 158, 11)' && linkColor !== iconColor, { iconColor, linkColor });

    // arrange
    await paneButton(page, 'Duplicate card');
    check('arrange: duplicate selects the copy', (await paneHeader(page)) === 'Card 2 of 2' && (await page.$$('#wp1 h3')).length === 2);
    p = await props(page);
    check('arrange: copy has a new id and same content', p.cards[0].id !== p.cards[1].id && p.cards[1].title === 'Recruiting Process');
    await setText(page, 'cards[1].title', 'Copy');
    await paneButton(page, 'Move card earlier');
    check('arrange: move earlier from the pane', (await paneHeader(page)) === 'Card 1 of 2' && (await page.textContent('#wp1 h3')) === 'Copy');
    page.__dialogAnswer = false;
    await paneButton(page, 'Delete card');
    check('arrange: cancelling the delete keeps the card', (await page.$$('#wp1 h3')).length === 2);
    page.__dialogAnswer = true;
    await paneButton(page, 'Delete card');
    check('arrange: delete removes card and returns to section settings', (await page.$$('#wp1 h3')).length === 1 && /Section settings/.test(await paneHeader(page)));

    // section settings
    check('section pane: lists cards and Add card', /1\. Recruiting Process/.test(await page.textContent('#pane')) && !!(await page.$('#pane button:text-is("Add a resource card")')));
    await setText(page, 'title', 'Recruiting & Hiring');
    await setText(page, 'description', 'Greenhouse, interviews, offers');
    check('section: header text updates', /Recruiting & Hiring/.test(await page.textContent('#wp1 button[aria-expanded]')));
    await page.uncheck('#pane .field[data-path="showCount"] input');
    check('section: count can be hidden', !(await page.$('#wp1 [class*=badge]')));
    await page.check('#pane .field[data-path="showCount"] input');
    await page.selectOption('#pane .field[data-path="headingLevel"] select', '3');
    check('section: heading level 3 -> cards h4', !!(await page.$('#wp1 h3 > button[aria-expanded]')) && !!(await page.$('#wp1 li h4')));
    check('section color: in the section\'s own settings, labelled "Section color"', !!(await page.$('#pane section[data-group="Section header"] .field[data-path="palette"] select')) && /Section color/.test(await page.textContent('#pane section[data-group="Section header"] .field[data-path="palette"]')));
    check('section color: other colors are under "More colors"', !!(await page.$('#pane section[data-group="More colors"] .field[data-path="headerBackground"]')) && !(await page.$('#pane section[data-group="More colors"] .field[data-path="palette"]')));
    await page.selectOption('#pane .field[data-path="palette"] select', 'custom');
    p = await props(page);
    check('section: custom palette starts from the previous palette', p.accentColor === '#2563eb' && p.tintColor === '#eff6ff' && p.borderColor === '#bfdbfe');
    check('section color: custom pickers appear in the section settings', !!(await page.$('#pane section[data-group="Section header"] .field[data-path="accentColor"]')));
    await page.evaluate(() => __paneChange('accentColor', '#be185d', 'custom'));
    const secIcon = await page.$eval('#wp1 button[aria-expanded] [class*=headerIcon]', (e) => getComputedStyle(e).color);
    const secBadge = await page.$eval('#wp1 [class*=badge]', (e) => getComputedStyle(e).borderTopColor);
    const tints = await page.evaluate(() => ({ tint: wp1.properties.tintColor, border: wp1.properties.borderColor }));
    check('section color: changes the section header icon, and its tint and border follow (count badge)', secIcon === 'rgb(190, 24, 93)' && secBadge !== 'rgb(191, 219, 254)' && tints.tint !== '#eff6ff' && tints.border !== '#bfdbfe', { secIcon, secBadge, tints });
    await page.evaluate(() => __paneChange('accentColor', '#2563eb', 'custom'));
    check('section: custom palette shows accent/tint/border pickers', !!(await page.$('#pane .field[data-path="accentColor"]')) && !!(await page.$('#pane .field[data-path="borderColor"]')));
    await page.selectOption('#pane .field[data-path="palette"] select', 'amber');
    check('section: preset hides custom pickers', !(await page.$('#pane .field[data-path="accentColor"]')));
    check('section: accessibility notes list the amber adjustments', /Resource count: #d97706/.test(await page.textContent('#pane')), await page.textContent('#pane section[data-group="Accessibility"]'));
    await page.check('#pane .field[data-path="iconSource"] input[value=image]');
    await setText(page, 'imageUrl', '/logo.png');
    check('section: header image', (await page.getAttribute('#wp1 button[aria-expanded] img', 'src')) === '/logo.png');
    await page.uncheck('#pane .field[data-path="useInterFont"] input');
    check('section: font toggle switches to SharePoint font', /Segoe UI/.test(await page.$eval('#wp1 > div', (d) => getComputedStyle(d).fontFamily.split(',')[0])));

    await paneButton(page, 'Add a resource card');
    check('section: Add card opens the new card', (await paneHeader(page)) === 'Card 2 of 2');
    await paneButton(page, 'Back to section settings');
    check('pane: Back returns to section settings', /Section settings/.test(await paneHeader(page)));
    await page.click('#wp1 li:nth-child(1) [class*=card]');
    check('canvas: clicking a card body selects it', (await paneHeader(page)) === 'Card 1 of 2');
    await page.evaluate(() => wp1.context.propertyPane.close());
    check('pane: closing clears the selection', !(await page.$('#wp1 [class*=selected]')));

    p = await props(page);
    check('persistence: no helper values saved in the web part data', !('paneAction' in p) && !('imagePickerResult' in p), Object.keys(p));
    p.headingLevel = '2';
    fs.writeFileSync(path.join(__dirname, 'authored.json'), JSON.stringify(p, null, 2));
    const missing = await page.evaluate(() => __missing);
    check('pane: no unknown SharePoint APIs used', missing.length === 0, [...new Set(missing)]);
    const errs = await page.evaluate(() => __errors);
    check('pane: no uncaught errors', errs.length === 0 && consoleErrors.filter((e) => !/Failed to load resource/.test(e)).length === 0, errs.concat(consoleErrors));
    await page.screenshot({ path: path.join(__dirname, 'authoring.png'), fullPage: true });
    await page.close();
  }

  // ------------------------------------------------------------------ 4. persistence round trip -> reader
  {
    const authored = JSON.parse(fs.readFileSync(path.join(__dirname, 'authored.json'), 'utf8'));
    const { page, consoleErrors } = await open(browser, { props: authored, mode: 1 });
    await page.click('#wp1 button[aria-expanded]');
    const titles = await page.$$eval('#wp1 li h3', (h) => h.map((x) => x.textContent));
    check('round trip: saved page renders the authored content', titles.join('|') === 'Recruiting Process|Resource title', titles);
    check('round trip: uploaded image path renders', /\/sites\/hr\/SiteAssets\/AccordionCards\//.test(await page.getAttribute('#wp1 li img', 'src')));
    await page.waitForFunction(() => document.querySelector('#wp1 li img').complete);
    check('round trip: uploaded image displays', await page.$eval('#wp1 li img', (i) => i.naturalWidth > 0));
    await axeRun(page, 'reader, authored content (amber, image, custom card accent)');
    check('round trip: no console errors', consoleErrors.filter((e) => !/Failed to load resource/.test(e)).length === 0, consoleErrors);
    await page.close();
  }

  // ------------------------------------------------------------------ 5. canvas toolbar + focus management
  {
    const three = { ...DEFAULTS, cards: ['A', 'B', 'C'].map((t) => ({ ...DEFAULTS.cards[0], id: 'id' + t, title: t })) };
    const { page } = await open(browser, { props: three });
    const order = () => page.$$eval('#wp1 li h3', (h) => h.map((x) => x.textContent).join(''));
    check('toolbar: first card cannot move earlier, last cannot move later',
      (await page.$eval('#wp1 [data-focus-key="earlier-idA"]', (b) => b.disabled)) && (await page.$eval('#wp1 [data-focus-key="later-idC"]', (b) => b.disabled)));
    await page.focus('#wp1 [data-focus-key="later-idA"]');
    await page.keyboard.press('Enter');
    await page.waitForTimeout(60);
    check('toolbar: move later via keyboard', (await order()) === 'BAC');
    check('toolbar: focus stays on the moved card\'s button', await page.evaluate(() => document.activeElement.dataset.focusKey === 'later-idA'));
    await page.keyboard.press('Enter');
    await page.waitForTimeout(60);
    check('toolbar: moved again', (await order()) === 'BCA');
    check('toolbar: focus falls back when the button becomes disabled', await page.evaluate(() => !!document.activeElement.dataset.focusKey));
    await page.focus('#wp1 [aria-label="Delete card"]');
    await page.keyboard.press('Enter');
    await page.waitForTimeout(60);
    check('toolbar: delete -> focus moves to the next card', (await order()) === 'CA' && (await page.evaluate(() => document.activeElement.dataset.focusKey)) === 'edit-idC');
    for (let i = 0; i < 2; i++) {
      await page.focus('#wp1 [aria-label="Delete card"]');
      await page.keyboard.press('Enter');
      await page.waitForTimeout(60);
    }
    check('toolbar: deleting the last card focuses Add', (await page.$$('#wp1 li h3')).length === 0 && (await page.evaluate(() => document.activeElement.dataset.focusKey)) === 'add');
    await page.keyboard.press('Enter');
    await page.waitForSelector('#pane .pane-header');
    check('toolbar: Add via keyboard adds and opens the card', (await page.$$('#wp1 li h3')).length === 1 && (await paneHeader(page)) === 'Card 1 of 1');
    await page.close();
  }

  // ------------------------------------------------------------------ 5b. section color vs card colors
  {
    const mixed = {
      ...DEFAULTS,
      palette: 'custom',
      accentColor: '#be185d',
      tintColor: '',
      borderColor: '',
      startExpanded: true,
      cards: [
        { ...DEFAULTS.cards[0], id: 'follows', title: 'Follows', accentMode: 'section' },
        { ...DEFAULTS.cards[0], id: 'own', title: 'Own', accentMode: 'custom', accentColor: '#0d9488' }
      ]
    };
    const { page, consoleErrors } = await open(browser, { props: mixed, mode: 1 });
    const c = await page.evaluate(() => ({
      section: getComputedStyle(document.querySelector('#wp1 button[aria-expanded] [class*=headerIcon]')).color,
      follows: getComputedStyle(document.querySelector('#wp1 li:nth-child(1) [class^="cardIcon_"]')).color,
      own: getComputedStyle(document.querySelector('#wp1 li:nth-child(2) [class^="cardIcon_"]')).color
    }));
    check('section color: custom section color shown on the header', c.section === 'rgb(190, 24, 93)', c);
    check('section color: "Match the section" cards follow it, cards with their own color keep theirs', c.follows === 'rgb(190, 24, 93)' && c.own === 'rgb(13, 148, 136)', c);
    await axeRun(page, 'custom section color with mixed card colors');
    check('section color: no console errors', consoleErrors.length === 0, consoleErrors);
    await page.close();
  }

  // ------------------------------------------------------------------ 6. robustness
  {
    for (const [label, bad] of [
      ['empty property bag', {}],
      ['null/partial cards', { title: 'X', cards: [null, { id: 'p', title: 'Partial' }, 'junk'] }],
      ['unknown palette & bad colors', { ...DEFAULTS, palette: 'nope', titleColor: 'not-a-color', cards: [{ ...DEFAULTS.cards[0], accentMode: 'custom', accentColor: 'zzz' }] }],
      ['missing icon name', { ...DEFAULTS, iconName: 'DoesNotExist', cards: [{ ...DEFAULTS.cards[0], iconName: 'Nope' }] }]
    ]) {
      const { page, consoleErrors } = await open(browser, { props: bad, mode: 1 });
      const ok = await page.$('#wp1 button[aria-expanded]');
      const errs = await page.evaluate(() => __errors);
      check(`robustness: ${label} renders without errors`, ok && errs.length === 0 && consoleErrors.length === 0, errs.concat(consoleErrors));
      await page.close();
    }
    const { page } = await open(browser, { props: { cards: [] }, mode: 1 });
    await page.click('#wp1 button[aria-expanded]');
    check('robustness: empty section tells readers', /No resources have been added/.test(await page.textContent('#wp1')));
    await axeRun(page, 'empty section');
    await page.close();
  }

  // ------------------------------------------------------------------ 7. security
  {
    const evil = {
      ...DEFAULTS,
      title: '<img src=x onerror="window.__pwned=1">',
      imageUrl: 'javascript:alert(1)',
      iconSource: 'image',
      cards: [
        { ...DEFAULTS.cards[0], id: 'e1', title: '<script>window.__pwned=1</script>', linkUrl: 'javascript:window.__pwned=1' },
        { ...DEFAULTS.cards[0], id: 'e2', linkUrl: '//evil.example.com', iconSource: 'image', imageUrl: 'data:text/html,<script>alert(1)</script>' },
        { ...DEFAULTS.cards[0], id: 'e3', linkUrl: ' JAVASCRIPT:alert(1)' },
        { ...DEFAULTS.cards[0], id: 'e4', linkUrl: 'mailto:hr@contoso.com' }
      ]
    };
    const { page } = await open(browser, { props: evil, mode: 1 });
    await page.click('#wp1 button[aria-expanded]');
    const hrefs = await page.$$eval('#wp1 a', (a) => a.map((x) => x.getAttribute('href')));
    check('security: only the mailto link is rendered', hrefs.length === 1 && hrefs[0] === 'mailto:hr@contoso.com', hrefs);
    check('security: unsafe image URLs are not rendered', (await page.$$('#wp1 img')).length === 0);
    check('security: markup in text is shown as text, never executed', !(await page.evaluate(() => window.__pwned)) && /<script>/.test(await page.textContent('#wp1')));
    await page.close();
  }

  // ------------------------------------------------------------------ 8. two instances on one page
  {
    const { page } = await open(browser, { mode: 1, two: true });
    const ids = await page.$$eval('[id]', (e) => e.map((x) => x.id));
    check('two instances: all element ids unique', new Set(ids).size === ids.length);
    await page.click('#wp1 button[aria-expanded]');
    check('two instances: toggling one leaves the other alone',
      (await page.getAttribute('#wp1 button[aria-expanded]', 'aria-expanded')) === 'true' && (await page.getAttribute('#wp2 button[aria-expanded]', 'aria-expanded')) === 'false');
    await page.close();
  }

  // ------------------------------------------------------------------ 9. layout, responsiveness, style isolation
  {
    const seven = { ...DEFAULTS, startExpanded: true, cards: Array.from({ length: 7 }, (_, i) => ({ ...DEFAULTS.cards[0], id: 'c' + i, title: 'Card ' + i, linkUrl: '/x' })) };
    const { page } = await open(browser, { props: seven, mode: 1 });
    const cols = () => page.$eval('#wp1 ul', (u) => getComputedStyle(u).gridTemplateColumns.split(' ').length);
    const widthCases = [[1180, 3], [968, 3], [967, 2], [700, 2], [592, 2], [591, 1], [320, 1]];
    for (const [w, expected] of widthCases) {
      await page.$eval('#wp1', (el, w) => (el.style.width = w + 'px'), w);
      await page.waitForTimeout(60);
      const c = await cols();
      check(`layout: web part ${w}px wide -> ${expected} column(s)`, c === expected, c);
    }
    await page.$eval('#wp1', (el) => (el.style.width = ''));
    await page.waitForTimeout(60);
    const s = await page.evaluate(() => {
      const q = (sel) => getComputedStyle(document.querySelector(sel));
      return {
        cardTitle: q('#wp1 li h3').fontSize,
        cardTitleWeight: q('#wp1 li h3').fontWeight,
        titleMargin: q('#wp1 li h3').marginTop,
        linkColor: q('#wp1 li a').color,
        linkDecoration: q('#wp1 li a').textDecorationLine,
        ulPadding: q('#wp1 ul').paddingLeft,
        liMargin: q('#wp1 li').marginBottom,
        pSize: q('#wp1 li p').fontSize,
        pMargin: q('#wp1 li p').marginTop,
        trigger: q('#wp1 button[aria-expanded]').fontFamily,
        triggerBorder: q('#wp1 button[aria-expanded]').borderTopStyle,
        headerTitle: q('#wp1 [id^=ra-title]').fontSize
      };
    });
    check('isolation: SharePoint-style globals do not leak in', s.cardTitle === '14px' && s.cardTitleWeight === '600' && s.titleMargin === '0px' && s.linkColor === 'rgb(37, 99, 235)' && s.linkDecoration === 'none' && s.ulPadding === '16px' && s.liMargin === '0px' && s.pSize === '12.48px' && s.pMargin === '0px' && /Inter/.test(s.trigger) && s.triggerBorder === 'none' && s.headerTitle === '15.2px', s);
    await page.close();

    const small = await open(browser, { props: { ...seven, title: 'An unusually long section title that has to wrap on small phones without overflowing', description: 'Supercalifragilisticexpialidocious-unbroken-description-text-that-must-wrap' }, mode: 1, viewport: { width: 320, height: 800 } });
    const overflow = await small.page.evaluate(() => document.documentElement.scrollWidth - window.innerWidth);
    check('reflow: no horizontal scrolling at 320px (WCAG 1.4.10)', overflow <= 0, overflow);
    await small.page.screenshot({ path: path.join(__dirname, 'mobile-320.png'), fullPage: true });
    await small.page.close();
  }

  // ------------------------------------------------------------------ 10. keyboard, focus, motion, high contrast
  {
    const withLinks = { ...DEFAULTS, cards: [0, 1].map((i) => ({ ...DEFAULTS.cards[0], id: 'k' + i, title: 'K' + i, linkUrl: '/page' + i })) };
    const { page } = await open(browser, { props: withLinks, mode: 1 });
    await page.keyboard.press('Tab');
    const first = await page.evaluate(() => document.activeElement.getAttribute('aria-expanded'));
    check('keyboard: Tab reaches the accordion button first', first === 'false');
    const ring = await page.evaluate(() => {
      const s = getComputedStyle(document.activeElement);
      return { style: s.outlineStyle, width: s.outlineWidth };
    });
    check('keyboard: visible focus outline on the header', ring.style === 'solid' && ring.width === '2px', ring);
    await page.keyboard.press('Tab');
    check('keyboard: collapsed cards are skipped', (await page.evaluate(() => document.activeElement.tagName)) === 'BODY' || !(await page.evaluate(() => !!document.activeElement.closest('#wp1 li'))));
    await page.keyboard.press('Shift+Tab');
    await page.keyboard.press('Enter');
    await page.keyboard.press('Tab');
    const onLink = await page.evaluate(() => ({ tag: document.activeElement.tagName, ring: getComputedStyle(document.activeElement).outlineStyle }));
    check('keyboard: Enter opens, Tab moves into the card links with a focus ring', onLink.tag === 'A' && onLink.ring === 'solid', onLink);
    await page.keyboard.press('Shift+Tab');
    await page.keyboard.press('Space');
    check('keyboard: Space collapses', (await page.getAttribute('#wp1 button[aria-expanded]', 'aria-expanded')) === 'false');
    await page.close();

    const rm = await open(browser, { props: withLinks, mode: 1, media: { reducedMotion: 'reduce' } });
    check('reduced motion: transitions disabled', (await rm.page.$eval('#wp1 [class*=chevron]', (c) => getComputedStyle(c).transitionDuration)) === '0s');
    await rm.page.close();

    const hc = await open(browser, { props: { ...withLinks, startExpanded: true }, mode: 1, media: { forcedColors: 'active' } });
    await axeRun(hc.page, 'forced colors (Windows High Contrast)');
    await hc.page.close();

    // hover state contrast for every palette
    for (const palette of ['blue', 'indigo', 'green', 'amber']) {
      const hv = await open(browser, { props: { ...withLinks, palette, startExpanded: true }, mode: 1 });
      await hv.page.hover('#wp1 li [class*=card]');
      await hv.page.hover('#wp1 button[aria-expanded]');
      await axeRun(hv.page, `contrast in ${palette} while hovering`, ['wcag2aa']);
      await hv.page.close();
    }
  }

  await browser.close();
  server.kill();
  console.log('\nThird-party (PnP chunk) dev-mode notices:');
  thirdParty.forEach((t) => console.log('  - ' + t));
  const failed = results.filter((r) => !r.ok);
  console.log(`\n${results.length - failed.length}/${results.length} checks passed`);
  process.exit(failed.length ? 1 : 0);
})().catch((e) => {
  console.error(e);
  process.exit(2);
});
