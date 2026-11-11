/*
 * A stand-in for the SharePoint page host. Provides the nine externals the
 * production bundle asks for, instantiates the real web part class from dist/,
 * and emulates the property pane closely enough to drive every field.
 */
(function () {
  const registry = {};
  window.__missing = [];
  window.__errors = [];
  window.__defineName = null;

  window.define = function (name, deps, factory) {
    if (typeof name !== 'string') {
      factory = deps;
      deps = name;
      name = window.__defineName;
    }
    if (typeof deps === 'function') {
      factory = deps;
      deps = [];
    }
    const resolved = deps.map((d) => {
      if (!(d in registry)) throw new Error('Missing external: ' + d);
      return registry[d];
    });
    registry[name] = factory.apply(null, resolved);
  };
  window.define.amd = true;
  window.__registry = registry;

  const watch = (label, obj) =>
    new Proxy(obj, {
      get(target, key) {
        if (typeof key === 'string' && !(key in target) && !['__esModule', 'default', 'then', 'toJSON', '$$typeof'].includes(key)) {
          window.__missing.push(label + '.' + key);
        }
        return target[key];
      }
    });

  registry['react'] = window.React;
  registry['react-dom'] = window.ReactDOM;
  registry['@microsoft/sp-lodash-subset'] = window.lodash;

  registry['@microsoft/sp-core-library'] = watch('core', {
    Version: {
      parse: (v) => {
        const [major, minor, patch] = String(v).split('.').map(Number);
        return { major, minor: minor || 0, patch: patch || 0, toString: () => v };
      }
    },
    Guid: { newGuid: () => ({ toString: () => crypto.randomUUID() }) },
    DisplayMode: { Read: 1, Edit: 2, 1: 'Read', 2: 'Edit' },
    Environment: { type: 2 },
    EnvironmentType: { Test: 0, Local: 1, SharePoint: 2, ClassicSharePoint: 3, 0: 'Test', 1: 'Local', 2: 'SharePoint', 3: 'ClassicSharePoint' },
    Log: { info() {}, warn() {}, error() {}, verbose() {} },
    Text: { format: (s, ...a) => s.replace(/\{(\d+)\}/g, (m, i) => a[i]) }
  });

  registry['@microsoft/sp-http'] = watch('http', {
    SPHttpClient: { configurations: { v1: { flags: 'v1' } } },
    HttpClient: { configurations: { v1: {} } },
    MSGraphClientV3: function () {}
  });

  // SPFx loads framework libraries needed only by lazy chunks through the component loader.
  const COMPONENTS = { 'c07208f0-ea3b-4c1a-9965-ac1b825211a6': '@microsoft/sp-http' };
  window.__loadedComponents = [];
  registry['@microsoft/sp-loader'] = watch('loader', {
    SPComponentLoader: {
      loadComponentById: (id, version) => {
        __loadedComponents.push(id + '@' + version);
        return COMPONENTS[id] ? Promise.resolve(registry[COMPONENTS[id]]) : Promise.reject(new Error('Unknown component ' + id));
      },
      loadCss() {},
      loadScript: () => Promise.resolve()
    }
  });

  const FieldType = { Custom: 1, CheckBox: 2, TextField: 3, Toggle: 5, Dropdown: 6, Label: 7, Slider: 8, Heading: 9, ChoiceGroup: 10, Button: 11, HorizontalRule: 12, Link: 13 };
  const field = (type) => (targetProperty, properties) => ({ type, targetProperty, properties: properties || {} });
  registry['@microsoft/sp-property-pane'] = watch('pane', {
    PropertyPaneFieldType: FieldType,
    PropertyPaneButtonType: { Normal: 0, Primary: 1, Hero: 2, Compound: 3, Command: 4, Icon: 5 },
    PropertyPaneTextField: field(FieldType.TextField),
    PropertyPaneToggle: field(FieldType.Toggle),
    PropertyPaneDropdown: field(FieldType.Dropdown),
    PropertyPaneLabel: field(FieldType.Label),
    PropertyPaneChoiceGroup: field(FieldType.ChoiceGroup),
    PropertyPaneButton: field(FieldType.Button),
    PropertyPaneHorizontalRule: () => ({ type: FieldType.HorizontalRule, properties: {} }),
    PropertyPaneLink: field(FieldType.Link),
    PropertyPaneSlider: field(FieldType.Slider)
  });

  class BaseClientSideWebPart {
    constructor() {
      this.properties = {};
    }
    get displayMode() {
      return this.__displayMode;
    }
    get domElement() {
      return this.__dom;
    }
    get instanceId() {
      return this.__instanceId;
    }
    get context() {
      return this.__context;
    }
    get manifest() {
      return { id: '9cdfb3f9-fb97-44a5-865b-06621546998c', alias: 'LinkCardsWebPart' };
    }
    onInit() {
      return Promise.resolve();
    }
    onPropertyPaneFieldChanged() {}
    onPropertyPaneConfigurationStart() {}
    onPropertyPaneConfigurationComplete() {}
    loadPropertyPaneResources() {
      return Promise.resolve();
    }
  }
  registry['@microsoft/sp-webpart-base'] = watch('webpartbase', { BaseClientSideWebPart });

  // ---------------------------------------------------------------- fake REST
  window.__rest = { calls: [], fail: false };
  const respond = (ok, body, status = ok ? 200 : 403) => ({
    ok,
    status,
    statusText: ok ? 'OK' : 'Forbidden',
    json: async () => body
  });
  const spHttpClient = {
    get: async (url, config) => {
      __rest.calls.push({ method: 'GET', url, config });
      if (/RootFolder/.test(url)) return respond(true, { ServerRelativeUrl: '/sites/hr/SiteAssets' });
      return respond(true, { value: [] });
    },
    post: async (url, config, options) => {
      __rest.calls.push({ method: 'POST', url, config, bodyBytes: options && options.body ? options.body.byteLength : 0 });
      if (__rest.fail && /Files\/add/.test(url)) return respond(false, { error: { message: 'Access denied.' } });
      if (/EnsureSiteAssetsLibrary/.test(url)) return respond(true, { Id: '11111111-2222-3333-4444-555555555555' });
      if (/folders\/add/.test(url)) return respond(true, {});
      if (/Files\/add/.test(url)) {
        const name = decodeURIComponent(/url='([^']+)'/.exec(url)[1]);
        const folder = decodeURIComponent(/ServerRelativeUrl\('([^']+)'\)/.exec(url)[1]);
        return respond(true, { ServerRelativeUrl: folder + '/' + name });
      }
      return respond(true, {});
    }
  };

  // ----------------------------------------------------------- fake pane UI
  const pane = { open: false, disposers: [], wp: null, config: null };
  window.__pane = pane;

  function change(path, newValue, kind) {
    const wp = pane.wp;
    const old = lodash.get(wp.properties, path);
    if (!lodash.isEqual(old, newValue) || kind === 'button' || kind === 'custom') {
      lodash.update(wp.properties, path, () => newValue);
      wp.onPropertyPaneFieldChanged(path, old, newValue);
      wp.render();
    }
  }
  window.__paneChange = change;

  function renderPane() {
    const root = document.getElementById('pane');
    pane.disposers.forEach((d) => d());
    pane.disposers = [];
    root.innerHTML = '';
    if (!pane.open) return;
    const wp = pane.wp;
    const config = wp.getPropertyPaneConfiguration();
    pane.config = config;
    config.pages.forEach((page) => {
      const header = document.createElement('p');
      header.className = 'pane-header';
      header.textContent = page.header ? page.header.description : '';
      root.appendChild(header);
      page.groups.forEach((group) => {
        const g = document.createElement('section');
        g.dataset.group = group.groupName || '';
        if (group.groupName) {
          const h = document.createElement('h3');
          h.textContent = group.groupName;
          g.appendChild(h);
        }
        // Like SharePoint: a collapsed group can only be opened on an accordion page;
        // otherwise its heading shows but its fields are unreachable.
        const unreachable = !!group.isCollapsed && !page.displayGroupsAsAccordion;
        g.dataset.unreachable = unreachable ? 'true' : 'false';
        if (!unreachable) group.groupFields.forEach((f) => g.appendChild(renderField(f, wp)));
        root.appendChild(g);
      });
    });
  }

  function renderField(f, wp) {
    const p = f.properties;
    const wrap = document.createElement('div');
    wrap.className = 'field';
    wrap.dataset.type = f.type;
    wrap.dataset.path = f.targetProperty || '';
    const value = f.targetProperty ? lodash.get(wp.properties, f.targetProperty) : undefined;
    const label = (text) => {
      const l = document.createElement('label');
      l.textContent = text || '';
      wrap.appendChild(l);
      return l;
    };
    switch (f.type) {
      case FieldType.TextField: {
        label(p.label);
        const input = document.createElement(p.multiline ? 'textarea' : 'input');
        input.value = value == null ? '' : value;
        input.placeholder = p.placeholder || '';
        const err = document.createElement('div');
        err.className = 'error';
        input.addEventListener('change', () => {
          const msg = p.onGetErrorMessage ? p.onGetErrorMessage(input.value) : '';
          err.textContent = msg;
          if (!msg) change(f.targetProperty, input.value, 'text');
        });
        wrap.appendChild(input);
        wrap.appendChild(err);
        break;
      }
      case FieldType.Toggle: {
        label(p.label);
        const input = document.createElement('input');
        input.type = 'checkbox';
        input.checked = !!value;
        input.addEventListener('change', () => change(f.targetProperty, input.checked, 'toggle'));
        wrap.appendChild(input);
        break;
      }
      case FieldType.Dropdown: {
        label(p.label);
        const select = document.createElement('select');
        p.options.forEach((o) => {
          const opt = document.createElement('option');
          opt.value = o.key;
          opt.textContent = o.text;
          select.appendChild(opt);
        });
        select.value = value;
        select.addEventListener('change', () => change(f.targetProperty, select.value, 'dropdown'));
        wrap.appendChild(select);
        break;
      }
      case FieldType.ChoiceGroup: {
        label(p.label);
        p.options.forEach((o) => {
          const input = document.createElement('input');
          input.type = 'radio';
          input.name = f.targetProperty;
          input.value = o.key;
          input.checked = value === o.key;
          input.addEventListener('change', () => change(f.targetProperty, o.key, 'choice'));
          wrap.appendChild(input);
          wrap.appendChild(document.createTextNode(o.text));
        });
        break;
      }
      case FieldType.Slider: {
        label(p.label);
        const input = document.createElement('input');
        input.type = 'range';
        input.min = p.min;
        input.max = p.max;
        input.step = p.step;
        input.value = value;
        input.addEventListener('change', () => change(f.targetProperty, Number(input.value), 'slider'));
        wrap.appendChild(input);
        break;
      }
      case FieldType.Label: {
        const t = document.createElement('p');
        t.className = 'label';
        t.textContent = p.text;
        wrap.appendChild(t);
        break;
      }
      case FieldType.Button: {
        const b = document.createElement('button');
        b.textContent = p.text;
        b.addEventListener('click', () => change(f.targetProperty, p.onClick(value), 'button'));
        wrap.appendChild(b);
        break;
      }
      case FieldType.Custom: {
        const host = document.createElement('div');
        host.className = 'custom';
        wrap.appendChild(host);
        p.onRender(host, undefined, (target, v) => change(target || f.targetProperty, v, 'custom'));
        if (p.onDispose) pane.disposers.push(() => p.onDispose(host));
        break;
      }
      default:
        wrap.textContent = 'UNSUPPORTED FIELD TYPE ' + f.type;
    }
    return wrap;
  }

  // ------------------------------------------------------ pane requests
  // SharePoint keeps one pane for the page and one "was it opened by a web part" mark, set by
  // whoever asked last: a web part's own open() (true) or the page itself (false), for example
  // when the author selects a web part or uses its edit button. Opening a pane that is already
  // open tells the web part it was showing "configuration complete" and the new one
  // "configuration start". The classic pane sets the mark before those events and raises
  // "start" 400ms after a fresh open; the newer content panel sets it after them.
  window.__paneMode = new URLSearchParams(location.search).get('pane') || 'classic';
  async function requestPane(wp, byWebPart) {
    const wasOpen = pane.open;
    const previous = pane.wp;
    const classic = __paneMode === 'classic';
    if (!wp.__paneLoaded) {
      await wp.loadPropertyPaneResources();
      wp.__paneLoaded = true;
    }
    pane.wp = wp;
    pane.open = true;
    if (classic) pane.byWebPart = byWebPart;
    renderPane();
    if (wasOpen) {
      previous.onPropertyPaneConfigurationComplete();
      wp.onPropertyPaneConfigurationStart();
    } else if (classic) {
      setTimeout(() => pane.open && pane.wp === wp && wp.onPropertyPaneConfigurationStart(), 400);
    } else {
      wp.onPropertyPaneConfigurationStart();
    }
    if (!classic) pane.byWebPart = byWebPart;
  }
  /** The page opening the pane for a web part: its edit button, or selecting it while the pane is open. */
  window.hostOpenPane = (wp) => requestPane(wp, false);

  // -------------------------------------------------------------- the host
  window.mountWebPart = async function ({ mountId, properties, displayMode, instanceId, moduleId }) {
    const prefix = (moduleId || window.__MODULE_ID).split('_')[0] + '_';
    const mod = registry[Object.keys(registry).find((k) => k.startsWith(prefix))];
    const WebPart = mod.default || mod;
    const wp = new WebPart();
    wp.__dom = document.getElementById(mountId);
    wp.__instanceId = instanceId || mountId;
    wp.__displayMode = displayMode || 2;
    wp.properties = properties;
    wp.__context = watch('context', {
      spHttpClient,
      httpClient: {},
      pageContext: watch('pageContext', {
        web: { absoluteUrl: 'https://contoso.sharepoint.com/sites/hr', serverRelativeUrl: '/sites/hr', id: { toString: () => 'web' }, title: 'HR' },
        site: { absoluteUrl: 'https://contoso.sharepoint.com/sites/hr', serverRelativeUrl: '/sites/hr', id: { toString: () => 'site' } },
        cultureInfo: { currentUICultureName: 'en-US', currentCultureName: 'en-US' },
        user: { loginName: 'author@contoso.com', displayName: 'Author' },
        legacyPageContext: {}
      }),
      propertyPane: {
        // Like SharePoint: both are about the pane as a whole, whichever web part it shows.
        isPropertyPaneOpen: () => pane.open,
        isRenderedByWebPart: () => !!pane.byWebPart,
        open: () => requestPane(wp, true),
        close: () => {
          const current = pane.wp;
          pane.open = false;
          if (__paneMode === 'contentPanel') pane.byWebPart = undefined;
          renderPane();
          current.onPropertyPaneConfigurationComplete();
        },
        refresh: () => {
          if (pane.open && pane.wp === wp) renderPane();
        }
      }
    });
    await wp.onInit();
    wp.render();
    return wp;
  };

  window.setDisplayMode = function (wp, mode) {
    wp.__displayMode = mode;
    wp.render();
  };

  window.addEventListener('error', (e) => __errors.push(String(e.message)));
  window.addEventListener('unhandledrejection', (e) => __errors.push('unhandled: ' + (e.reason && e.reason.message)));
})();
