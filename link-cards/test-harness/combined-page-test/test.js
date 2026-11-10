/*
 * Both web parts together on realistic SharePoint-style pages, checked as a
 * whole page: heading outline, one H1, unique ids, landmarks, contrast, reflow.
 */
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const AXE = fs.readFileSync(require.resolve('../harness/node_modules/axe-core/axe.min.js'), 'utf8');
const HEADER = require('/home/user/Manager/custom-width-header/src/webparts/pageHeader/PageHeaderWebPart.manifest.json').preconfiguredEntries[0].properties;
const ACC = require('/home/user/Manager/accordion-cards/src/webparts/resourceAccordion/ResourceAccordionWebPart.manifest.json').preconfiguredEntries[0].properties;
const LINKS = require('/home/user/Manager/link-cards/src/webparts/linkCards/LinkCardsWebPart.manifest.json').preconfiguredEntries[0].properties;
const LINKS_ID = '9cdfb3f9-fb97-44a5-865b-06621546998c_';
const HEADER_ID = '121dcf48-85b3-4120-9d1e-7a019d67bdeb_';
const ACC_ID = 'b13052d3-2957-4800-83d5-cf5963855d13_';

const results = [];
function check(name, ok, detail) {
  results.push({ name, ok: !!ok });
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${!ok && detail !== undefined ? '  -> ' + JSON.stringify(detail) : ''}`);
}

const cards = (n, prefix) =>
  Array.from({ length: n }, (_, i) => ({ ...ACC.cards[0], id: `${prefix}${i}`, title: `${prefix} resource ${i + 1}`, linkUrl: `/sites/hr/SitePages/${prefix}-${i}.aspx` }));

/**
 * A page is a list of blocks, top to bottom, mirroring SharePoint's structure:
 * optional title area, then sections holding web parts.
 */
async function buildPage(browser, blocks, { mode = 1, viewport = { width: 1440, height: 900 } } = {}) {
  const page = await browser.newPage({ viewport });
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  page.on('console', (m) => {
    if ((m.type() === 'error' || m.type() === 'warning') && !/was used but not registered/.test(m.text())) errors.push(m.text());
  });
  await page.goto('http://localhost:4379/index.html');
  await page.evaluate(
    async ({ blocks, mode, HEADER_ID, ACC_ID, LINKS_ID }) => {
      const main = document.getElementById('page');
      let n = 0;
      for (const b of blocks) {
        if (b.kind === 'title') {
          const t = document.createElement('div');
          t.id = 'titleArea';
          t.innerHTML = `<h1>${b.text}</h1>`;
          main.appendChild(t);
          continue;
        }
        const section = document.createElement('div');
        section.className = 'section' + (b.full ? ' full' : '');
        const inner = document.createElement('div');
        if (!b.full) inner.className = 'column' + (b.cols ? ' cols' : '');
        section.appendChild(inner);
        main.appendChild(section);
        for (const w of b.parts) {
          const host = document.createElement('div');
          host.className = 'wp';
          inner.appendChild(host);
          if (w.kind === 'text') {
            host.innerHTML = w.html;
            continue;
          }
          host.id = 'wp' + ++n;
          await mountWebPart({
            mountId: host.id,
            properties: JSON.parse(JSON.stringify(w.props)),
            displayMode: mode,
            instanceId: 'inst' + n,
            moduleId: w.kind === 'header' ? HEADER_ID : w.kind === 'links' ? LINKS_ID : ACC_ID
          });
        }
      }
    },
    { blocks, mode, HEADER_ID, ACC_ID, LINKS_ID }
  );
  await page.waitForTimeout(250);
  return { page, errors };
}

async function outline(page) {
  return page.evaluate(() =>
    Array.from(document.querySelectorAll('h1,h2,h3,h4,h5,h6,[role=heading]'))
      .filter((h) => !h.closest('[aria-hidden=true]') && h.getClientRects().length > 0)
      .map((h) => ({ level: h.getAttribute('aria-level') ? +h.getAttribute('aria-level') : +h.tagName[1], text: h.textContent.trim().slice(0, 40) }))
  );
}

function outlineProblems(list) {
  const problems = [];
  if (!list.length || list[0].level !== 1) problems.push('first heading is not level 1');
  if (list.filter((h) => h.level === 1).length !== 1) problems.push(`${list.filter((h) => h.level === 1).length} level-1 headings`);
  for (let i = 1; i < list.length; i++) {
    if (list[i].level > list[i - 1].level + 1) problems.push(`skip: h${list[i - 1].level} "${list[i - 1].text}" -> h${list[i].level} "${list[i].text}"`);
  }
  return problems;
}

async function audit(page, label, { photo = false } = {}) {
  await page.evaluate(AXE);
  const r = await page.evaluate(() =>
    axe.run(document, { runOnly: ['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa', 'wcag22aa', 'best-practice'] })
  );
  const v = r.violations.map((x) => `${x.id}: ${x.nodes.map((n) => n.target.join(' ')).slice(0, 2).join(' | ')}`);
  check(`${label}: whole page passes axe (incl. heading-order, one H1, unique ids, landmarks)`, v.length === 0, v);
  const review = r.incomplete.filter((x) => x.id === 'color-contrast').reduce((n, x) => n + x.nodes.length, 0);
  if (!photo) check(`${label}: nothing left for manual contrast review`, review === 0, review);
  const o = await outline(page);
  const p = outlineProblems(o);
  check(`${label}: heading outline ${o.map((h) => 'h' + h.level).join(' ')} has no skips or extra H1`, p.length === 0, { p, o });
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

  const header = (extra = {}) => ({ kind: 'header', props: { ...HEADER, title: 'Hire & Onboard', ...extra } });
  const accordion = (title, extra = {}) => ({ kind: 'accordion', props: { ...ACC, title, startExpanded: true, cards: cards(3, title.split(' ')[0]), ...extra } });
  const text = (html) => ({ kind: 'text', html });
  const links = (extra = {}) => ({
    kind: 'links',
    props: {
      ...LINKS,
      cards: [
        { ...LINKS.cards[0], id: 'l1', text: 'Offboarding SOP', url: '/sites/hr/Shared%20Documents/sop.docx' },
        { ...LINKS.cards[0], id: 'l2', text: 'Exit interview form', url: '/sites/hr/SitePages/exit.aspx', openInNewTab: false, iconName: 'ClipboardList' },
        { ...LINKS.cards[0], id: 'l3', text: 'Orientation deck coming soon', url: '' }
      ],
      ...extra
    }
  });

  // A. Article page — SharePoint shows the page title (H1) above the canvas.
  {
    const { page, errors } = await buildPage(browser, [
      { kind: 'title', text: 'Hire &amp; Onboard' },
      { full: true, parts: [header()] },
      { parts: [text('<h2>Overview</h2><p>How hiring works here.</p>'), links(), accordion('Recruiting & Hiring'), accordion('Onboarding')] }
    ]);
    await audit(page, 'article page');
    check('article page: header became H2 under the page title', (await page.$eval('#wp1 h2', (h) => h.textContent)) === 'Hire & Onboard');
    check('article page: no errors', errors.length === 0, errors);
    await page.screenshot({ path: path.join(__dirname, 'article-page.png'), fullPage: true });
    await page.close();
  }

  // B. Home-style page — no title area, the header is the page's main title.
  {
    const { page, errors } = await buildPage(browser, [
      { full: true, parts: [header()] },
      { parts: [accordion('Recruiting & Hiring'), links({ palette: 'green' }), accordion('Policies & Compliance')] }
    ]);
    await audit(page, 'home page');
    check('home page: header is the H1', (await page.$eval('#wp1 h1', (h) => h.textContent)) === 'Hire & Onboard');
    check('home page: no errors', errors.length === 0, errors);
    await page.close();
  }

  // C. Content header over a photo with the solid text panel; accordion one level down under a sub-heading.
  {
    const { page, errors } = await buildPage(browser, [
      { kind: 'title', text: 'Greenhouse' },
      { full: true, parts: [header({ layout: 'resource', eyebrow: 'Recruiting & Hiring', showBackLink: false, backgroundType: 'image', backgroundImageUrl: '/checker.png', imageTextStyle: 'panel' })] },
      { parts: [text('<h2>Using Greenhouse</h2>'), accordion('Interview kits', { headingLevel: '3' })] }
    ]);
    await page.waitForFunction(() => Array.from(document.images).every((i) => i.complete));
    await audit(page, 'photo header (panel) + nested accordion');
    check('photo header: no errors', errors.length === 0, errors);
    await page.close();
  }

  // D. Photo header with the tint (axe can't measure text on photos; the header suite measures pixels instead).
  {
    const { page, errors } = await buildPage(browser, [
      { kind: 'title', text: 'Support &amp; Retain' },
      { full: true, parts: [header({ backgroundType: 'image', backgroundImageUrl: '/checker.png', imageTextStyle: 'overlay', height: 'medium' })] },
      { parts: [accordion('Employee wellbeing')] }
    ]);
    await page.waitForFunction(() => Array.from(document.images).every((i) => i.complete));
    await audit(page, 'photo header (tint)', { photo: true });
    check('photo header (tint): no errors', errors.length === 0, errors);
    await page.close();
  }

  // E. Everything in edit mode (toolbars, Add tile, hints) is still clean.
  {
    const { page, errors } = await buildPage(
      browser,
      [
        { kind: 'title', text: 'Lead &amp; Develop' },
        { full: true, parts: [header({ palette: 'indigo' })] },
        { parts: [accordion('Performance Management'), links({ palette: 'indigo' }), accordion('Compensation')] }
      ],
      { mode: 2 }
    );
    await audit(page, 'edit mode');
    check('edit mode: no errors', errors.length === 0, errors);
    await page.close();
  }

  // F. Every palette side by side on one page.
  {
    const blocks = [{ kind: 'title', text: 'Palettes' }];
    for (const palette of ['blue', 'indigo', 'green', 'amber']) {
      blocks.push({ full: true, parts: [header({ palette, title: palette, eyebrow: 'Label' })] });
      blocks.push({ parts: [accordion(`${palette} section`, { palette }), links({ palette })] });
    }
    const { page, errors } = await buildPage(browser, blocks);
    await audit(page, 'four palettes: headers, accordions and link cards');
    const ids = await page.$$eval('[id]', (e) => e.map((x) => x.id));
    check('many web parts: every id on the page is unique', new Set(ids).size === ids.length);
    check('four palettes: no errors', errors.length === 0, errors);
    await page.close();
  }

  // G. Phone width: two-column section collapses, nothing scrolls sideways.
  {
    const { page, errors } = await buildPage(
      browser,
      [
        { kind: 'title', text: 'Exit &amp; Transition' },
        { full: true, parts: [header({ description: 'Guidance for managing departures with care — from separation conversations to backfill coordination.' })] },
        { cols: true, parts: [accordion('Offboarding'), links({ palette: 'amber' })] }
      ],
      { viewport: { width: 375, height: 800 } }
    );
    const overflow = await page.evaluate(() => document.documentElement.scrollWidth - document.documentElement.clientWidth);
    check('phone: no horizontal scrolling with both web parts (WCAG 1.4.10)', overflow <= 0, overflow);
    await audit(page, 'phone');
    check('phone: no errors', errors.length === 0, errors);
    await page.screenshot({ path: path.join(__dirname, 'phone-page.png'), fullPage: true });
    await page.close();
  }

  // H. Control: prove the outline check catches real problems (so the passes above mean something).
  {
    const { page } = await buildPage(browser, [
      { kind: 'title', text: 'Control' },
      { full: true, parts: [header({ headingLevel: '1' })] },
      { parts: [accordion('Skipped', { headingLevel: '4' })] }
    ]);
    const p = outlineProblems(await outline(page));
    check('control: a forced H1 + an H4 accordion is reported (duplicate H1 and skipped levels)', p.some((x) => /2 level-1/.test(x)) && p.some((x) => /skip/.test(x)), p);
    await page.close();
  }

  await browser.close();
  const failed = results.filter((r) => !r.ok);
  console.log(`\n${results.length - failed.length}/${results.length} checks passed`);
  process.exit(failed.length ? 1 : 0);
})().catch((e) => {
  console.error(e);
  process.exit(2);
});
