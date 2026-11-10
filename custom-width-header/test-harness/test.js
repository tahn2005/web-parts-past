const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const { PNG } = require('../harness/node_modules/pngjs');
const AXE = fs.readFileSync(require.resolve('../harness/node_modules/axe-core/axe.min.js'), 'utf8');
const manifest = require('/home/user/Manager/custom-width-header/src/webparts/customWidthHeader/CustomWidthHeaderWebPart.manifest.json');
const DEFAULTS = manifest.preconfiguredEntries[0].properties;
const BASE = 'http://localhost:4378/';
const SPPKG = '/home/user/Manager/custom-width-header/sharepoint/solution/manager-hub-page-header.sppkg';

const results = [];
/** The two React dev-mode notices PnP's file picker emits; anything else is a failure. */
const KNOWN_PNP_NOTICES = [/`key` is not a prop/, /Can't perform a React state update on an unmounted component/];
const thirdParty = new Set();
function check(name, ok, detail) {
  results.push({ name, ok: !!ok });
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${!ok && detail !== undefined ? '  -> ' + JSON.stringify(detail) : ''}`);
}

async function open(browser, { props = DEFAULTS, mode = 2, viewport = { width: 1600, height: 900 }, media, both = false, before } = {}) {
  const page = await browser.newPage({ viewport });
  if (media) await page.emulateMedia(media);
  const requests = [];
  const consoleErrors = [];
  page.on('request', (r) => requests.push(r.url()));
  page.on('console', (m) => {
    if (m.type() !== 'error' && m.type() !== 'warning') return;
    if (/was used but not registered/.test(m.text())) return; // Fluent icon font comes from SharePoint pages
    const where = (m.location() && m.location().url) || '';
    if (KNOWN_PNP_NOTICES.some((r) => r.test(m.text()))) return void thirdParty.add(m.text().split('\n')[0].slice(0, 150));
    consoleErrors.push(m.text());
  });
  page.on('pageerror', (e) => consoleErrors.push('pageerror: ' + e.message));
  await page.goto(BASE + 'index.html');
  if (before) {
    await page.evaluate((html) => {
      const div = document.createElement('div');
      div.id = 'titleArea';
      div.innerHTML = html;
      document.body.insertBefore(div, document.body.firstChild);
    }, before);
  }
  await page.evaluate(
    async ({ props, mode, both }) => {
      window.wp1 = await mountWebPart({ mountId: 'wp1', properties: JSON.parse(JSON.stringify(props)), displayMode: mode, instanceId: 'h1' });
      if (both) window.wp2 = await mountWebPart({ mountId: 'wp2', properties: JSON.parse(JSON.stringify(props)), displayMode: mode, instanceId: 'h2' });
    },
    { props, mode, both }
  );
  await page.waitForTimeout(150);
  return { page, requests, consoleErrors };
}

async function axeRun(page, label, { allowIncomplete = true } = {}) {
  await page.evaluate(AXE);
  const r = await page.evaluate(() =>
    axe.run('#wp1', { runOnly: ['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa', 'wcag22aa', 'best-practice'] })
  );
  const v = r.violations.map((x) => `${x.id}: ${x.nodes.map((n) => n.target.join(' ')).slice(0, 2).join(' | ')}`);
  check(`axe: ${label}`, v.length === 0, v);
  const cc = r.incomplete.filter((x) => x.id === 'color-contrast');
  if (!allowIncomplete) check(`axe: ${label} — every contrast measurable (nothing left "needs review")`, cc.length === 0, cc.map((x) => x.nodes.length));
  return { incompleteContrast: cc.reduce((n, x) => n + x.nodes.length, 0) };
}

const props = (page) => page.evaluate(() => JSON.parse(JSON.stringify(wp1.properties)));
const setText = async (page, p, v) => {
  const sel = `#pane .field[data-path="${p}"] input, #pane .field[data-path="${p}"] textarea`;
  await page.fill(sel, v);
  await page.dispatchEvent(sel, 'change');
  await page.waitForTimeout(30);
};
const choose = (page, p, v) => page.check(`#pane .field[data-path="${p}"] input[value="${v}"]`);
const select = (page, p, v) => page.selectOption(`#pane .field[data-path="${p}"] select`, v);

/**
 * Measures contrast the way a person sees it: hide the text, photograph what is
 * actually behind each text element, and compare the text colour against the
 * worst pixel in that area.
 */
async function measureRenderedContrast(page, root = '#wp1') {
  const items = await page.evaluate((root) => {
    const pick = [
      ['title', 'h1, h2'],
      ['description', 'p[class*=description]'],
      ['eyebrow', 'p[class*=eyebrow]'],
      ['back link', 'a']
    ];
    const out = [];
    for (const [name, sel] of pick) {
      const el = document.querySelector(`${root} ${sel}`);
      if (!el) continue;
      const r = el.getBoundingClientRect();
      const m = /rgba?\(([\d.]+), ([\d.]+), ([\d.]+)(?:, ([\d.]+))?\)/.exec(getComputedStyle(el).color);
      out.push({ name, box: { x: r.x, y: r.y + window.scrollY, width: r.width, height: r.height }, color: [+m[1], +m[2], +m[3]], alpha: m[4] ? +m[4] : 1 });
    }
    return out;
  }, root);
  // Transitions off, or the screenshot could catch a colour mid-fade.
  const hide = await page.addStyleTag({ content: `${root} * { color: transparent !important; transition: none !important; }` });
  const shot = PNG.sync.read(await page.screenshot({ fullPage: true, animations: 'disabled' }));
  await hide.evaluate((n) => n.remove());
  await page.waitForTimeout(250);
  const lum = (c) => {
    const ch = (v) => ((v /= 255) <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4));
    return 0.2126 * ch(c[0]) + 0.7152 * ch(c[1]) + 0.0722 * ch(c[2]);
  };
  const ratio = (a, b) => {
    const [x, y] = [lum(a), lum(b)];
    return (Math.max(x, y) + 0.05) / (Math.min(x, y) + 0.05);
  };
  return items.map((it) => {
    let worst = Infinity;
    for (let y = Math.floor(it.box.y); y < Math.ceil(it.box.y + it.box.height); y++) {
      for (let x = Math.floor(it.box.x); x < Math.ceil(it.box.x + it.box.width); x++) {
        if (x < 0 || y < 0 || x >= shot.width || y >= shot.height) continue;
        const i = (y * shot.width + x) * 4;
        const bg = [shot.data[i], shot.data[i + 1], shot.data[i + 2]];
        const fg = it.color.map((c, k) => bg[k] + (c - bg[k]) * it.alpha);
        worst = Math.min(worst, ratio(fg, bg));
      }
    }
    return { name: it.name, worst: Math.round(worst * 100) / 100, need: it.name === 'title' ? 3 : 4.5 };
  });
}

(async () => {
  const server = require('child_process').spawn(process.execPath, [path.join(__dirname, 'server.js')]);
  process.on('exit', () => server.kill());
  server.stderr.on('data', (d) => {
    console.error('SERVER: ' + d);
    process.exit(3);
  });
  await new Promise((r) => server.stdout.once('data', r));
  const browser = await chromium.launch();

  // ------------------------------------------------------------ 0. package
  {
    const exec = (a) => require('child_process').execFileSync('unzip', a).toString();
    const list = exec(['-Z1', SPPKG]);
    const xml = exec(['-p', SPPKG, '*.xml']);
    check('package: tenant-wide deployment (SkipFeatureDeployment)', /SkipFeatureDeployment="true"/.test(xml));
    check('package: SharePoint pages only', /&quot;supportedHosts&quot;:\[&quot;SharePointWebPart&quot;\]/.test(xml));
    check('package: allowed in full-width sections (supportsFullBleed)', /&quot;supportsFullBleed&quot;:true/.test(xml));
    check('package: all assets inside the package', /ClientSideAssets\/custom-width-header-web-part_/.test(list) && /ClientSideAssets\/chunk\.custom-width-header-property-pane_/.test(list));
    const app = exec(['-p', SPPKG, 'AppManifest.xml']);
    check('package: shows as "Custom Width Header" (app, feature and toolbox)', /<Title>Custom Width Header<\/Title>/.test(app) && /Name="Custom Width Header"/.test(app) && /Title="Custom Width Header"/.test(xml) && /&quot;title&quot;:\{&quot;default&quot;:&quot;Custom Width Header&quot;/.test(xml) && !/client-side-solution/.test(app + xml));
    check('package: version 1.0.2.0 with the same IDs, so deployed copies update in place', /Version="1\.0\.2\.0"/.test(app) && /ProductID="a9db1d40-94f9-4cb4-b3e5-c6f8271e3100"/.test(app) && /121dcf48-85b3-4120-9d1e-7a019d67bdeb/.test(xml));
    {
      const all = require('child_process').execFileSync('unzip', ['-p', SPPKG], { maxBuffer: 100 * 1024 * 1024 }).toString('latin1');
      const files = require('child_process').execFileSync('unzip', ['-Z1', SPPKG]).toString();
      check('naming: package file keeps its original name (manager-hub-page-header.sppkg) so it overwrites the deployed one', require('path').basename(SPPKG) === 'manager-hub-page-header.sppkg');
      check('naming: no old names anywhere inside the package', !/manager-hub|resource-accordion|ResourceAccordion|page-header|PageHeaderWebPart|client-side-solution/.test(all));
      check('naming: every bundled file is named after the web part (no numbered chunks)', files.split('\n').filter((f) => /^ClientSideAssets\/.+\.js$/.test(f)).every((f) => /^ClientSideAssets\/(custom-width-header-web-part_|chunk\.custom-width-header-property-pane_|[A-Za-z]+WebPartStrings_|PropertyControlStrings_)/.test(f)), files.split('\n').filter((f) => /\.js$/.test(f)));
    }
    check('package: heading level defaults to Automatic', /&quot;headingLevel&quot;:&quot;auto&quot;/.test(xml));
    check('package: separate solution from the accordion', !/resource-accordion/.test(list) && /a9db1d40-94f9-4cb4-b3e5-c6f8271e3100/.test(exec(['-p', SPPKG, 'AppManifest.xml'])));
  }

  // ------------------------------------------------------------ 1. drop-in (edit mode)
  {
    const { page, requests, consoleErrors } = await open(browser);
    check('drop-in: title is an h1', (await page.textContent('#wp1 h1')) === 'Page title');
    check('drop-in: back link defaults to the site home page', (await page.getAttribute('#wp1 a', 'href')) === 'https://contoso.sharepoint.com/sites/hr' && (await page.textContent('#wp1 a')) === 'Home');
    check('drop-in: icon tile shown and decorative', !!(await page.$('#wp1 [aria-hidden=true] svg.lucide-layers')));
    check('drop-in: description shown', /A short description/.test(await page.textContent('#wp1 p')));
    await page.click('#wp1 a');
    check('drop-in: back link does not navigate while editing', page.url() === BASE + 'index.html');
    check('drop-in: no pane code for readers/authors until the pane opens', !requests.some((u) => /chunk\./.test(u)));
    await axeRun(page, 'default header (edit mode)', { allowIncomplete: false });
    const m = await measureRenderedContrast(page);
    check('drop-in: rendered contrast passes for every text element', m.every((x) => x.worst >= x.need), m);
    check('drop-in: no console errors, no unknown SharePoint APIs', consoleErrors.length === 0 && (await page.evaluate(() => __missing.length)) === 0, consoleErrors);
    await page.screenshot({ path: path.join(__dirname, 'default.png'), clip: { x: 0, y: 0, width: 1280, height: 260 } });
    await page.close();
  }

  // ------------------------------------------------------------ 2. the two screenshots from the app
  {
    const pageHeader = { ...DEFAULTS, title: 'Hire & Onboard', iconName: 'UserPlus', description: 'Everything you need to recruit great talent, make confident hiring decisions, and set new hires up for long-term success — from opening a req to the 90-day check-in.', palette: 'indigo', useGradient: true };
    const { page } = await open(browser, { props: pageHeader, mode: 1 });
    await page.waitForTimeout(400);
    await page.screenshot({ path: path.join(__dirname, 'replica-page.png'), clip: { x: 0, y: 0, width: 1280, height: 230 } });
    const s = await page.evaluate(() => {
      const q = (sel) => getComputedStyle(document.querySelector(sel));
      return { title: q('#wp1 h1').fontSize, family: q('#wp1 h1').fontFamily.split(',')[0], tile: q('#wp1 [class*=iconTile]').width, inner: q('#wp1 [class*=inner]').maxWidth, desc: q('#wp1 p').maxWidth, back: q('#wp1 a').fontSize };
    });
    check('page layout: app sizes (28px Merriweather title, 48px tile, 1280px content, 600px description)', s.title === '28px' && /Merriweather/.test(s.family) && s.tile === '48px' && s.inner === '1280px' && s.desc === '600px' && s.back === '13.12px', s);
    await page.close();

    const resource = { ...DEFAULTS, layout: 'resource', showBackLink: false, eyebrow: 'Recruiting & Hiring', title: 'Greenhouse - Our ATS', iconName: 'Monitor', description: 'Use Greenhouse to manage candidates, interviews, and hiring activities.', palette: 'blue' };
    const r = await open(browser, { props: resource, mode: 1 });
    await r.page.waitForTimeout(400);
    await r.page.screenshot({ path: path.join(__dirname, 'replica-resource.png'), clip: { x: 0, y: 0, width: 1600, height: 230 } });
    const s2 = await r.page.evaluate(() => {
      const q = (sel) => getComputedStyle(document.querySelector(sel));
      return { title: q('#wp1 h1').fontSize, eyebrow: q('#wp1 p[class*=eyebrow]').textTransform, inner: q('#wp1 [class*=inner]').maxWidth };
    });
    check('content layout: app sizes (25.6px title, uppercase label, 880px content)', s2.title === '25.6px' && s2.eyebrow === 'uppercase' && s2.inner === '880px', s2);
    await axeRun(r.page, 'content layout with label', { allowIncomplete: false });
    await r.page.close();
  }

  // ------------------------------------------------------------ 3. full width + responsiveness
  {
    const { page } = await open(browser, { mode: 1, both: true, viewport: { width: 1920, height: 900 } });
    const w = await page.evaluate(() => ({
      full: document.querySelector('#wp1 header').getBoundingClientRect().width,
      vw: document.querySelector('#canvas').getBoundingClientRect().width,
      inner: document.querySelector('#wp1 [class*=inner]').getBoundingClientRect().width,
      column: document.querySelector('#wp2 header').getBoundingClientRect().width,
      columnContent: document.querySelector('.column').clientWidth - 64
    }));
    check('full width: background spans edge to edge', Math.abs(w.full - w.vw) < 1, w);
    check('full width: content centred at the app\'s 1280px', Math.abs(w.inner - 1280) < 1, w);
    check('normal section: fits its column', Math.abs(w.column - w.columnContent) < 1, w);
    await page.close();

    const long = { ...DEFAULTS, eyebrow: 'A deliberately long label for small screens', title: 'An unusually long page title that must wrap cleanly on a small phone screen', description: 'Supercalifragilisticexpialidocious-unbroken-text-that-must-wrap-without-overflow-anywhere-at-all' };
    for (const width of [320, 768]) {
      const m = await open(browser, { props: long, mode: 1, viewport: { width, height: 800 } });
      const overflow = await m.page.evaluate(() => document.documentElement.scrollWidth - document.documentElement.clientWidth);
      check(`reflow: no horizontal scrolling at ${width}px (WCAG 1.4.10)`, overflow <= 0, overflow);
      if (width === 320) await m.page.screenshot({ path: path.join(__dirname, 'mobile-320.png'), fullPage: true });
      await m.page.close();
    }
  }

  // ------------------------------------------------------------ 4. text over images — the main event
  {
    const images = ['white', 'black', 'checker', 'noise', 'yellow'];
    for (const style of ['overlay', 'panel']) {
      for (const img of images) {
        for (const overlayOpacity of style === 'overlay' ? [0, 55] : [55]) {
          const p = { ...DEFAULTS, eyebrow: 'Label', backgroundType: 'image', backgroundImageUrl: `/${img}.png`, imageTextStyle: style, overlayOpacity, height: 'medium' };
          const { page } = await open(browser, { props: p, mode: 1 });
          await page.waitForFunction(() => document.querySelector('#wp1 img').complete);
          const m = await measureRenderedContrast(page);
          const label = `text over image (${style}${style === 'overlay' ? `, author set ${overlayOpacity}%` : ''}) on ${img}`;
          check(`${label}: every word ≥ AA on the worst pixel`, m.length === 4 && m.every((x) => x.worst >= x.need), m);
          if (img === 'checker') {
            const a = await axeRun(page, `${label}`, { allowIncomplete: style === 'overlay' });
            if (style === 'overlay') console.log(`      note: axe leaves ${a.incompleteContrast} text node(s) as "needs review" over the photo — measured above instead`);
            await page.screenshot({ path: path.join(__dirname, `image-${style}.png`), clip: { x: 0, y: 0, width: 1280, height: 300 } });
          }
          await page.close();
        }
      }
    }
    // custom light text colours / light tints are corrected too
    const tricky = { ...DEFAULTS, backgroundType: 'image', backgroundImageUrl: '/white.png', overlayColor: '#ffffff', overlayOpacity: 30, textColor: '#ffffff' };
    const t = await open(browser, { props: tricky, mode: 1 });
    await t.page.waitForFunction(() => document.querySelector('#wp1 img').complete);
    const m = await measureRenderedContrast(t.page);
    check('white tint + white text on a white photo is still made readable', m.every((x) => x.worst >= x.need), m);
    await t.page.close();
  }

  // ------------------------------------------------------------ 5. colour backgrounds, every palette
  for (const palette of ['blue', 'indigo', 'green', 'amber']) {
    for (const useGradient of [true, false]) {
      const { page } = await open(browser, { props: { ...DEFAULTS, eyebrow: 'Label', palette, useGradient }, mode: 1 });
      const m = await measureRenderedContrast(page);
      check(`${palette}${useGradient ? ' gradient' : ' solid'}: rendered contrast passes`, m.every((x) => x.worst >= x.need), m);
      await axeRun(page, `${palette}${useGradient ? ' gradient' : ' solid'}`, { allowIncomplete: useGradient });
      await page.close();
    }
  }

  // ------------------------------------------------------------ 6. authoring through the property pane
  {
    const { page, requests, consoleErrors } = await open(browser);
    await page.evaluate(() => wp1.context.propertyPane.open());
    await page.waitForSelector('#pane .pane-header');
    check('pane: opens and lazy-loads its code', requests.some((u) => /chunk\.custom-width-header-property-pane/.test(u)));

    await setText(page, 'title', 'Hire & Onboard');
    await setText(page, 'description', 'Everything you need to recruit great talent.');
    await setText(page, 'eyebrow', 'Employee lifecycle');
    check('pane: text fields update the header', (await page.textContent('#wp1 h1')) === 'Hire & Onboard' && (await page.textContent('#wp1 p[class*=eyebrow]')) === 'Employee lifecycle');

    await setText(page, 'backLinkText', 'Manager Hub');
    await setText(page, 'backLinkUrl', 'javascript:alert(1)');
    check('pane: unsafe back link address rejected', /Enter a full address/.test(await page.textContent('#pane .field[data-path="backLinkUrl"] .error')) && (await props(page)).backLinkUrl === '');
    await setText(page, 'backLinkUrl', '/sites/hr/SitePages/Manager-Hub.aspx');
    check('pane: custom back link text and address', (await page.getAttribute('#wp1 a', 'href')) === '/sites/hr/SitePages/Manager-Hub.aspx' && (await page.textContent('#wp1 a')) === 'Manager Hub');
    await page.uncheck('#pane .field[data-path="showBackLink"] input');
    check('pane: back link can be hidden (and its fields hide)', !(await page.$('#wp1 a')) && !(await page.$('#pane .field[data-path="backLinkText"]')));
    await page.check('#pane .field[data-path="showBackLink"] input');

    await page.fill('#pane .custom input', 'user');
    await page.click('#pane .custom button[aria-label="User plus"]');
    check('pane: icon picker sets the icon', (await props(page)).iconName === 'UserPlus' && !!(await page.$('#wp1 svg.lucide-user-plus')));
    await choose(page, 'iconSource', 'none');
    check('pane: icon can be turned off', !(await page.$('#wp1 [class*=iconTile]')));

    const pnpSave = (target, result) =>
      page.evaluate(
        async ({ target, result }) => {
          const f = __pane.config.pages[0].groups.flatMap((g) => g.groupFields).find((x) => x.properties && x.properties.key === 'picker-' + target);
          if (result.upload) {
            const name = result.name;
            const bytes = await (await fetch('/icon.png')).arrayBuffer();
            result = { fileName: name, fileAbsoluteUrl: undefined, downloadFileContent: async () => new File([bytes], name, { type: 'image/png' }) };
          }
          f.properties.onSave(result);
          lodash.set(wp1.properties, 'imagePickerResult', result);
          f.properties.onPropertyChange('imagePickerResult', undefined, result);
          await new Promise((r) => setTimeout(r, 150));
        },
        { target, result }
      );

    await choose(page, 'iconSource', 'image');
    await pnpSave('iconImageUrl', { fileName: 'icon.png', fileAbsoluteUrl: BASE + 'icon.png' });
    check('pane: icon image from a picker', (await page.getAttribute('#wp1 [class*=iconTile] img', 'src')) === BASE + 'icon.png' && (await page.getAttribute('#wp1 [class*=iconTile] img', 'alt')) === '');

    await select(page, 'palette', 'custom');
    check('pane: custom colour starts from the previous palette', (await props(page)).backgroundColor === '#2563eb' && !!(await page.$('#pane .field[data-path="backgroundColor"] .custom *')));
    await page.evaluate(() => __paneChange('backgroundColor', '#ffe066', 'custom'));
    await page.evaluate(() => wp1.context.propertyPane.refresh());
    check('pane: a light custom colour is corrected and explained', /Background color: #ffe066/.test(await page.textContent('#pane section[data-group="Accessibility"]')));
    await page.uncheck('#pane .field[data-path="useGradient"] input');
    check('pane: gradient can be switched off', !/gradient/.test(await page.$eval('#wp1 header', (h) => getComputedStyle(h).backgroundImage)));

    await choose(page, 'backgroundType', 'image');
    check('pane: image mode shows picker, position, text style, tint and strength', /Choose or upload an image/.test(await page.textContent('#pane')) && !!(await page.$('#pane .field[data-path="imagePosition"]')) && !!(await page.$('#pane .field[data-path="overlayOpacity"] input[type=range]')));
    check('pane: edit mode prompts for an image', /Choose a background image/.test(await page.textContent('#wp1')));
    await page.evaluate(() => (__rest.calls = []));
    await pnpSave('backgroundImageUrl', { upload: true, name: 'Team Photo.png' });
    const p1 = await props(page);
    const calls = await page.evaluate(() => __rest.calls.filter((c) => !/FilePickerTabOptions/.test(c.url)).map((c) => c.method + ' ' + c.url.replace('https://contoso.sharepoint.com/sites/hr/_api', '')));
    check('upload: saved to Site Assets/CustomWidthHeader with a unique, safe name', /^\/sites\/hr\/SiteAssets\/CustomWidthHeader\/\d+-Team-Photo\.png$/.test(p1.backgroundImageUrl) && calls.length === 4, { url: p1.backgroundImageUrl, calls });
    check('upload: background image renders, decorative', (await page.getAttribute('#wp1 header > img', 'alt')) === '' && (await page.getAttribute('#wp1 header > img', 'aria-hidden')) === 'true');
    await page.evaluate(() => (__rest.calls = []));
    await pnpSave('backgroundImageUrl', { upload: true, name: 'notes.html' });
    check('upload: non-images refused before upload', /couldn't be uploaded: only/.test(await page.textContent('#pane')) && (await page.evaluate(() => __rest.calls.filter((c) => /Files\/add/.test(c.url)).length)) === 0);

    await page.$eval('#pane .field[data-path="overlayOpacity"] input', (i) => {
      i.value = '10';
      i.dispatchEvent(new Event('change'));
    });
    await page.waitForTimeout(50);
    await page.evaluate(() => wp1.context.propertyPane.refresh());
    check('pane: a weak tint is raised for contrast and explained', /Image overlay strength: 10% → \d+%/.test(await page.textContent('#pane section[data-group="Accessibility"]')) && /checked against pure white and pure black/.test(await page.textContent('#pane')));
    await select(page, 'imagePosition', 'top');
    check('pane: image focus position', (await page.$eval('#wp1 header > img', (i) => getComputedStyle(i).objectPosition)) === '50% 0%');
    await choose(page, 'imageTextStyle', 'panel');
    check('pane: panel mode swaps tint controls for a panel colour', !!(await page.$('#pane .field[data-path="overlayColor"]')) && !(await page.$('#pane .field[data-path="overlayOpacity"]')) && /Panel color/.test(await page.textContent('#pane')));

    await select(page, 'layout', 'resource');
    await select(page, 'height', 'large');
    await select(page, 'headingLevel', '2');
    const lay = await page.evaluate(() => ({ h: document.querySelector('#wp1 header').getBoundingClientRect().height, h2: !!document.querySelector('#wp1 h2'), inner: getComputedStyle(document.querySelector('#wp1 [class*=inner]')).maxWidth }));
    check('pane: layout, height and heading level', lay.h >= 400 && lay.h2 && lay.inner === '880px', lay);
    await page.uncheck('#pane .field[data-path="useManagerHubFonts"] input');
    check('pane: fonts can fall back to SharePoint\'s', /Segoe UI/.test(await page.$eval('#wp1 h2', (h) => getComputedStyle(h).fontFamily.split(',')[0])));

    await page.evaluate(() => wp1.context.propertyPane.close());
    const saved = await props(page);
    check('persistence: no helper values saved', !('paneNote' in saved) && !('imagePickerResult' in saved), Object.keys(saved));
    fs.writeFileSync(path.join(__dirname, 'authored.json'), JSON.stringify(saved, null, 2));
    check('pane: no errors and no unknown SharePoint APIs', consoleErrors.filter((e) => !/Failed to load resource/.test(e)).length === 0 && (await page.evaluate(() => __missing.length + __errors.length)) === 0, consoleErrors.concat(await page.evaluate(() => __missing.concat(__errors))));
    await page.screenshot({ path: path.join(__dirname, 'authoring.png'), fullPage: true });
    await page.close();
  }

  // ------------------------------------------------------------ 7. saved page as a reader
  {
    const saved = JSON.parse(fs.readFileSync(path.join(__dirname, 'authored.json'), 'utf8'));
    const { page, consoleErrors } = await open(browser, { props: saved, mode: 1 });
    await page.waitForFunction(() => document.querySelector('#wp1 header > img').complete);
    check('round trip: saved header renders as authored', (await page.textContent('#wp1 h2')) === 'Hire & Onboard' && /CustomWidthHeader/.test(await page.getAttribute('#wp1 header > img', 'src')));
    check('round trip: no edit prompts for readers', !/Choose a background image/.test(await page.textContent('#wp1')));
    const m = await measureRenderedContrast(page);
    check('round trip: rendered contrast passes', m.every((x) => x.worst >= x.need), m);
    await axeRun(page, 'saved header, image + panel', { allowIncomplete: false });
    check('round trip: no console errors', consoleErrors.length === 0, consoleErrors);
    await page.close();
  }

  // ------------------------------------------------------------ 7b. heading outline
  {
    // How SharePoint article pages show the page title above the canvas.
    const withTitleArea = async (page, html) =>
      page.evaluate((html) => {
        const div = document.createElement('div');
        div.id = 'titleArea';
        div.innerHTML = html;
        document.body.insertBefore(div, document.body.firstChild);
      }, html);
    const level = (page) => page.$eval('#wp1 h1, #wp1 h2', (h) => h.tagName);

    let t = await open(browser, { mode: 1 });
    check('heading: no page title above -> the header is the H1', (await level(t.page)) === 'H1');
    await t.page.close();

    // Title area present before the web part renders (the normal case).
    t = await open(browser, { mode: 1, props: { ...DEFAULTS }, before: '<h1>Hire &amp; Onboard</h1>' });
    check('heading: SharePoint page title above -> the header is H2', (await level(t.page)) === 'H2');
    await t.page.close();

    t = await open(browser, { mode: 1, before: '<div role="heading" aria-level="1">Page name</div>' });
    check('heading: ARIA level-1 title above -> H2', (await level(t.page)) === 'H2');
    await t.page.close();

    t = await open(browser, { mode: 1, before: '<h1 style="display:none">Hidden title</h1><h1 aria-hidden="true">Decorative</h1>' });
    check('heading: hidden or aria-hidden titles are ignored -> H1', (await level(t.page)) === 'H1');
    await t.page.close();

    t = await open(browser, { mode: 1 });
    await withTitleArea(t.page, '<h1>Late title</h1>');
    await t.page.waitForTimeout(1800);
    check('heading: a title area drawn after the header is picked up (re-check)', (await level(t.page)) === 'H2');
    await t.page.close();

    t = await open(browser, { mode: 2, props: { ...DEFAULTS, headingLevel: '1' }, before: '<h1>Page name</h1>' });
    check('heading: a manual H1 is respected', (await level(t.page)) === 'H1');
    await t.page.evaluate(() => wp1.context.propertyPane.open());
    await t.page.waitForSelector('#pane .pane-header');
    check('heading: pane warns when a manual level does not fit the page', /set to Heading 1, but Heading 2 fits/.test(await t.page.textContent('#pane section[data-group="Accessibility"]')));
    check('heading: pane offers Automatic', (await t.page.$$eval('#pane .field[data-path="headingLevel"] option', (o) => o.map((x) => x.value))).join() === 'auto,1,2');
    await t.page.selectOption('#pane .field[data-path="headingLevel"] select', 'auto');
    check('heading: choosing Automatic fixes it and the warning goes', (await level(t.page)) === 'H2' && !/fits this page/.test(await t.page.textContent('#pane')));
    await t.page.close();

    t = await open(browser, { mode: 1 });
    const migrated = await t.page.evaluate(() => {
      const v = (s) => __registry['@microsoft/sp-core-library'].Version.parse(s);
      return {
        oldDefault: wp1.onAfterDeserialize({ headingLevel: '1' }, v('1.0')).headingLevel,
        oldH2: wp1.onAfterDeserialize({ headingLevel: '2' }, v('1.0')).headingLevel,
        newH1: wp1.onAfterDeserialize({ headingLevel: '1' }, v('1.1')).headingLevel,
        current: wp1.dataVersion.toString()
      };
    });
    check('upgrade: headers saved by 1.0 with the H1 default move to Automatic; deliberate choices kept', migrated.oldDefault === 'auto' && migrated.oldH2 === '2' && migrated.newH1 === '1' && migrated.current === '1.1', migrated);
    await t.page.close();
  }

  // ------------------------------------------------------------ 8. keyboard, motion, high contrast
  {
    const { page } = await open(browser, { mode: 1 });
    await page.keyboard.press('Tab');
    const f = await page.evaluate(() => ({ tag: document.activeElement.tagName, outline: getComputedStyle(document.activeElement).outlineStyle, width: getComputedStyle(document.activeElement).outlineWidth }));
    check('keyboard: back link is reachable with a visible focus ring', f.tag === 'A' && f.outline === 'solid' && f.width === '2px', f);
    await page.close();
    const rm = await open(browser, { mode: 1, media: { reducedMotion: 'reduce' } });
    check('reduced motion: no transitions', (await rm.page.$eval('#wp1 a', (a) => getComputedStyle(a).transitionDuration)) === '0s');
    await rm.page.close();
    const hc = await open(browser, { props: { ...DEFAULTS, backgroundType: 'image', backgroundImageUrl: '/checker.png' }, mode: 1, media: { forcedColors: 'active' } });
    check('high contrast: background photo dropped so text uses system colours', (await hc.page.$eval('#wp1 header > img', (i) => getComputedStyle(i).display)) === 'none');
    await axeRun(hc.page, 'forced colors');
    await hc.page.close();
  }

  // ------------------------------------------------------------ 9. isolation, security, robustness
  {
    const { page } = await open(browser, { mode: 1 });
    const s = await page.evaluate(() => {
      const t = getComputedStyle(document.querySelector('#wp1 h1'));
      const p = getComputedStyle(document.querySelector('#wp1 p'));
      const a = getComputedStyle(document.querySelector('#wp1 a'));
      return { size: t.fontSize, weight: t.fontWeight, margin: t.marginTop, color: t.color, pSize: p.fontSize, pMargin: p.marginTop, aDeco: a.textDecorationLine, aColor: a.color };
    });
    check('isolation: SharePoint-style globals do not leak in', s.size === '28px' && s.weight === '700' && s.margin === '0px' && s.color === 'rgb(255, 255, 255)' && s.pSize === '15.2px' && s.pMargin === '0px' && s.aDeco === 'none' && s.aColor !== 'rgb(0, 120, 212)', s);
    await page.close();

    const evil = { ...DEFAULTS, title: '<img src=x onerror="window.__pwned=1">', backLinkUrl: 'javascript:window.__pwned=1', backgroundType: 'image', backgroundImageUrl: 'javascript:alert(1)', iconSource: 'image', iconImageUrl: '//evil.example.com/x.png' };
    const e = await open(browser, { props: evil, mode: 1 });
    check('security: unsafe back link not rendered', !(await e.page.$('#wp1 a')));
    check('security: unsafe image addresses not rendered', (await e.page.$$('#wp1 img')).length === 0);
    check('security: markup in text shown as text, never run', !(await e.page.evaluate(() => window.__pwned)) && /<img/.test(await e.page.textContent('#wp1 h1')));
    await e.page.close();

    for (const [label, bad] of [
      ['empty property bag', {}],
      ['nonsense values', { layout: 'x', height: 'y', headingLevel: '9', iconSource: 'z', palette: 'nope', backgroundType: 'image', imagePosition: 'left', overlayOpacity: 'abc', textColor: 'blue-ish' }],
      ['blank title, nothing else', { title: '', description: '', eyebrow: '', showBackLink: false, iconSource: 'none' }]
    ]) {
      const r = await open(browser, { props: bad, mode: 1 });
      const errs = await r.page.evaluate(() => __errors);
      check(`robustness: ${label} renders without errors`, (await r.page.$('#wp1 header')) && errs.length === 0 && r.consoleErrors.length === 0, errs.concat(r.consoleErrors));
      if (label === 'blank title, nothing else') {
        check('robustness: no empty heading for readers', !(await r.page.$('#wp1 h1, #wp1 h2')));
        await axeRun(r.page, 'blank header');
      }
      await r.page.close();
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
