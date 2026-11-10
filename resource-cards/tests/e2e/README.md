# End-to-end tests

`e2e.cjs` loads the **production bundle** from `release/assets` (what's inside the
`.sppkg`) into `index.html`, a simulated SharePoint page. `host.js` supplies:

- an AMD loader and stand-ins for the SPFx modules the bundle imports
  (`@microsoft/sp-webpart-base`, `sp-property-pane`, `sp-core-library`), with React 18
  provided externally, as SharePoint does;
- a property pane that follows SPFx's rules: values are read and written with
  lodash `get`/`update` on each field's `targetProperty`, then
  `onPropertyPaneFieldChanged` runs and the web part re-renders (reactive mode).
  Button fields write their `onClick` result back, and custom fields get
  `onRender`/`onDispose`;
- a SharePoint REST stand-in for image uploads, which records each call and
  stores files in an in-memory "Site Assets" library that serves them back;
- hostile page CSS (red serif headings, large margins, green links) to show that
  SharePoint's styles don't leak into the cards.

The test drives it with Playwright in Chromium. It covers rendering, edit
mode, every property pane control, add/move/delete from both the canvas and
the pane, save and reload, malformed and malicious stored data, responsive
columns, reflow at 320px, reduced motion, forced colours, RTL, keyboard focus,
and axe-core accessibility scans.

```bash
npm run build                                   # produce release/
npm i --no-save playwright@1 axe-core@4         # one-time test tools
npx playwright install chromium                 # if you don't already have it
node tests/e2e/e2e.cjs
```

The simulator isn't SharePoint. Always finish with a smoke test on a real page
(see the main README).
