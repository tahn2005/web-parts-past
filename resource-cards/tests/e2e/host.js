// Minimal stand-in for the SharePoint page runtime: an AMD loader, the SPFx
// modules the bundle imports, and a property pane that follows SPFx's
// semantics (lodash get/update on targetProperty, then
// onPropertyPaneFieldChanged, then a reactive re-render).
(function () {
  const params = new URLSearchParams(location.search);
  const dev = params.get('dev') === '1';
  const log = (window.__log = []);

  const modules = {};
  let pendingName = null;
  const amdDefine = function (name, deps, factory) {
    if (typeof name !== 'string') { factory = deps; deps = name; name = pendingName; }
    if (typeof deps === 'function') { factory = deps; deps = []; }
    modules[name] = { deps, factory };
  };
  amdDefine.amd = {};
  function req(name) {
    const m = modules[name];
    if (!m) throw new Error('AMD module missing: ' + name);
    if (!('value' in m)) m.value = m.factory.apply(null, m.deps.map(req));
    return m.value;
  }
  function loadScript(src, amdName) {
    return new Promise((resolve, reject) => {
      pendingName = amdName || null;
      const s = document.createElement('script');
      s.src = src; s.onload = resolve; s.onerror = () => reject(new Error('load ' + src));
      document.head.appendChild(s);
    });
  }

  const host = (window.host = {
    mode: params.get('edit') === '1' ? 2 : 1,
    paneOpen: false,
    customFields: [],
    instance: null,
    WebPartClass: null
  });

  class BaseClientSideWebPart {
    constructor() { this.properties = {}; }
    get displayMode() { return host.mode; }
    get context() { return host.context; }
    onInit() { return Promise.resolve(); }
    onPropertyPaneFieldChanged() {}
    onDisplayModeChanged() {}
    onPropertyPaneConfigurationComplete() {}
    get disableReactivePropertyChanges() { return false; }
  }
  // Same shim sp-loader applies to framework classes so ES5-compiled web parts
  // (tslib __extends + Base.apply(this, arguments)) can extend native classes.
  (function shim(a) {
    a.apply = function (b, c) { const d = window.Reflect; return d && d.construct && this == a ? d.construct(a, c, b.constructor) : Function.apply.apply(this, [b, c]); };
    a.call = function (b, ...c) { const d = window.Reflect; return d && d.construct && this == a ? d.construct(a, c, b.constructor) : Function.apply.apply(this, [b, c]); };
  })(BaseClientSideWebPart);
  const fieldFactory = (type) => (targetProperty, properties) => ({ type, targetProperty, properties });

  modules['@microsoft/sp-core-library'] = { value: { Version: { parse: (v) => ({ v }) }, DisplayMode: { Read: 1, Edit: 2 } } };
  modules['@microsoft/sp-webpart-base'] = { value: { BaseClientSideWebPart } };
  modules['@microsoft/sp-property-pane'] = { value: {
    PropertyPaneFieldType: { Custom: 'Custom' },
    PropertyPaneButtonType: { Normal: 0, Primary: 1, Hero: 2, Compound: 3, Command: 4, Icon: 5 },
    PropertyPaneTextField: fieldFactory('Text'),
    PropertyPaneToggle: fieldFactory('Toggle'),
    PropertyPaneDropdown: fieldFactory('Dropdown'),
    PropertyPaneChoiceGroup: fieldFactory('ChoiceGroup'),
    PropertyPaneLabel: fieldFactory('Label'),
    PropertyPaneButton: fieldFactory('Button')
  } };

  // ---- property pane ---------------------------------------------------------
  const paneEl = () => document.getElementById('pane');
  function change(targetProperty, value) {
    const wp = host.instance;
    const old = _.get(wp.properties, targetProperty);
    _.update(wp.properties, targetProperty, () => value);
    wp.onPropertyPaneFieldChanged(targetProperty, old, value);
    wp.render();
  }
  function disposeCustom() {
    host.customFields.forEach(({ el, f }) => f.properties.onDispose && f.properties.onDispose(el));
    host.customFields = [];
  }
  function renderPane() {
    const wp = host.instance;
    const pane = paneEl();
    // Preserve focus target across rebuilds, as the real pane (React) does.
    const focusedId = document.activeElement && pane.contains(document.activeElement) ? document.activeElement.id : null;
    disposeCustom();
    pane.innerHTML = '';
    const close = document.createElement('button');
    close.textContent = 'Close pane'; close.id = 'pane-close';
    close.onclick = () => host.closePane();
    pane.appendChild(close);
    const config = wp.getPropertyPaneConfiguration();
    host.lastConfig = config;
    config.pages.forEach((page) => {
      const header = document.createElement('p'); header.className = 'pane-header'; header.textContent = page.header.description; pane.appendChild(header);
      page.groups.forEach((group, gi) => {
        const g = document.createElement('section'); g.className = 'group';
        const h = document.createElement('h2'); h.textContent = group.groupName; g.appendChild(h);
        group.groupFields.forEach((f, fi) => {
          const wrap = document.createElement('div'); wrap.className = 'field'; wrap.dataset.type = f.type; wrap.dataset.target = f.targetProperty;
          const id = `f-${gi}-${fi}-${f.type}-${f.targetProperty}`.replace(/[^\w-]/g, '_');
          const p = f.properties;
          const val = _.get(wp.properties, f.targetProperty);
          if (f.type === 'Text') {
            const lab = document.createElement('label'); lab.textContent = p.label; lab.htmlFor = id;
            const input = document.createElement(p.multiline ? 'textarea' : 'input');
            if (!p.multiline) input.type = 'text';
            input.id = id; input.value = val == null ? '' : val;
            const err = document.createElement('div'); err.className = 'err';
            const validate = () => { err.textContent = p.onGetErrorMessage ? p.onGetErrorMessage(input.value) || '' : ''; };
            input.oninput = () => { validate(); change(f.targetProperty, input.value); };
            validate();
            wrap.append(lab, input, err);
          } else if (f.type === 'Toggle') {
            const lab = document.createElement('label'); const cb = document.createElement('input'); cb.type = 'checkbox'; cb.id = id;
            cb.checked = !!val; cb.onchange = () => change(f.targetProperty, cb.checked);
            lab.append(cb, ' ' + p.label); wrap.append(lab);
          } else if (f.type === 'Dropdown') {
            const lab = document.createElement('label'); lab.textContent = p.label; lab.htmlFor = id;
            const sel = document.createElement('select'); sel.id = id;
            p.options.forEach((o) => { const opt = document.createElement('option'); opt.value = String(o.key); opt.textContent = o.text; sel.appendChild(opt); });
            sel.value = String(val);
            sel.onchange = () => { const o = p.options.filter((x) => String(x.key) === sel.value)[0]; change(f.targetProperty, o.key); };
            wrap.append(lab, sel);
          } else if (f.type === 'ChoiceGroup') {
            const fs = document.createElement('fieldset'); const lg = document.createElement('legend'); lg.textContent = p.label; fs.appendChild(lg);
            p.options.forEach((o) => {
              const lab = document.createElement('label'); const r = document.createElement('input'); r.type = 'radio'; r.name = id; r.value = String(o.key);
              r.checked = val === o.key; r.onchange = () => change(f.targetProperty, o.key);
              lab.append(r, ' ' + o.text); fs.appendChild(lab);
            });
            wrap.append(fs);
          } else if (f.type === 'Label') {
            const t = document.createElement('p'); t.className = 'pane-label'; t.textContent = p.text; wrap.append(t);
          } else if (f.type === 'Button') {
            const b = document.createElement('button'); b.id = id; b.textContent = p.text; b.disabled = !!p.disabled;
            if (p.ariaLabel) b.setAttribute('aria-label', p.ariaLabel);
            b.onclick = () => change(f.targetProperty, p.onClick(_.get(wp.properties, f.targetProperty)));
            wrap.append(b);
          } else if (f.type === 'Custom') {
            const el = document.createElement('div'); wrap.append(el);
            p.onRender(el, undefined, (tp, v) => change(tp, v));
            host.customFields.push({ el, f });
          } else {
            throw new Error('Unknown field type ' + f.type);
          }
          g.appendChild(wrap);
        });
        pane.appendChild(g);
      });
    });
    if (focusedId) { const el = document.getElementById(focusedId); if (el) el.focus(); }
  }
  host.closePane = () => { host.paneOpen = false; paneEl().classList.remove('open'); disposeCustom(); host.instance.onPropertyPaneConfigurationComplete(); };
  // SharePoint REST stand-in for the Site Assets upload. Records every call;
  // set window.__uploadFail = <status> to simulate a refusal.
  host.httpCalls = [];
  const response = (status, body) => ({ ok: status >= 200 && status < 300, status, json: async () => body });
  modules['@microsoft/sp-http'] = { value: { SPHttpClient: { configurations: { v1: { flavor: 'v1' } } } } };
  const WEB = location.origin + '/sites/hr';
  host.context_spHttpClient = {
    async post(url, config, options) {
      host.httpCalls.push({ method: 'POST', url, config, bodyType: options && options.body ? Object.prototype.toString.call(options.body) : undefined });
      if (window.__uploadFail) return response(window.__uploadFail, { error: { message: 'Access denied.' } });
      if (/\/_api\/web\/lists\/EnsureSiteAssetsLibrary$/.test(url)) return response(200, { Id: '2f1f0b52-0000-4000-8000-000000000001' });
      const m = /GetFolderByServerRelativePath\(decodedurl='([^']*)'\)\/Files\/AddUsingPath\(decodedurl='([^']*)',overwrite=false\)$/.exec(url);
      if (m) {
        const folder = decodeURIComponent(m[1]); const name = decodeURIComponent(m[2]);
        await fetch(folder + '/' + encodeURIComponent(name), { method: 'PUT', body: options.body });
        return response(200, { ServerRelativeUrl: folder + '/' + name });
      }
      return response(404, { error: { message: 'Unknown endpoint ' + url } });
    },
    async get(url, config) {
      host.httpCalls.push({ method: 'GET', url, config });
      if (/lists\(guid'2f1f0b52-0000-4000-8000-000000000001'\)\/RootFolder\?\$select=ServerRelativeUrl$/.test(url)) return response(200, { ServerRelativeUrl: '/sites/hr/SiteAssets' });
      return response(404, {});
    }
  };
  host.context = {
    pageContext: { web: { absoluteUrl: WEB } },
    spHttpClient: host.context_spHttpClient,
    propertyPane: {
      open() { host.paneOpen = true; paneEl().classList.add('open'); renderPane(); },
      refresh() { if (host.paneOpen) renderPane(); },
      isPropertyPaneOpen() { return host.paneOpen; }
    }
  };
  host.setMode = (mode) => { const old = host.mode; host.mode = mode; if (mode !== 2 && host.paneOpen) host.closePane(); host.instance.onDisplayModeChanged(old); };

  host.mount = async (properties) => {
    if (host.instance) { host.instance.onDispose(); document.getElementById('wp').innerHTML = ''; }
    const wp = new host.WebPartClass();
    wp.domElement = document.getElementById('wp');
    wp.properties = properties;
    host.instance = wp;
    await wp.onInit();
    wp.render();
    return wp;
  };
  host.serialize = () => JSON.stringify(host.instance.properties);

  window.addEventListener('error', (e) => log.push('window.error: ' + e.message));
  const origError = console.error; console.error = function () { log.push('console.error: ' + [].map.call(arguments, String).join(' ')); origError.apply(console, arguments); };
  const origWarn = console.warn; console.warn = function () { log.push('console.warn: ' + [].map.call(arguments, String).join(' ')); origWarn.apply(console, arguments); };

  (async () => {
    await loadScript(dev ? 'vendor/react.development.js' : 'vendor/react.production.min.js');
    await loadScript(dev ? 'vendor/react-dom.development.js' : 'vendor/react-dom.production.min.js');
    await loadScript('vendor/lodash.min.js');
    await loadScript('files.js');
    // Install the AMD loader only now: React/lodash UMD builds would
    // otherwise register as anonymous AMD modules instead of globals.
    window.define = amdDefine;
    modules['react'] = { value: window.React };
    modules['react-dom/client'] = { value: window.ReactDOM };
    modules['react-dom'] = { value: window.ReactDOM };
    const manifest = await (await fetch('manifests/manifest.json')).json();
    await loadScript('assets/' + window.__files.strings, 'ResourceCardsWebPartStrings');
    await loadScript('assets/' + window.__files.bundle);
    const mod = req(manifest.id + '_' + manifest.version);
    host.WebPartClass = mod.default || mod;
    host.manifest = manifest;
    const propsText = window.__props || params.get('props');
    const props = propsText ? JSON.parse(propsText) : JSON.parse(JSON.stringify(manifest.preconfiguredEntries[0].properties));
    await host.mount(props);
    window.__ready = true;
  })().catch((e) => { log.push('boot: ' + e.stack); window.__ready = 'failed'; });
})();
