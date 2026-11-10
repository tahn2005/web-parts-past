/**
 * Link safety. Authors type URLs freely, so anything that could run script
 * (javascript:, data:, vbscript:) is refused rather than rendered.
 */

const SAFE_SCHEMES = ['http:', 'https:', 'mailto:', 'tel:'];

export type UrlCheck = { ok: true; url: string } | { ok: false; reason: string };

export function checkUrl(value: string | undefined): UrlCheck {
  // Browsers drop tabs, newlines and other control characters from URLs, so
  // "java\tscript:" would still run. Remove them before checking the scheme.
  // Plain spaces are kept: SharePoint paths ("Shared Documents") contain them.
  // eslint-disable-next-line no-control-regex
  const url = (value || '').replace(/[\u0000-\u001F\u007F]/g, '').trim();

  if (url === '') {
    return { ok: true, url: '' };
  }

  // Site-relative ("/sites/hr/..."), page-relative ("page.aspx") and anchors
  // ("#section") are fine; "//host" is protocol-relative and treated as a URL.
  const schemeMatch = /^([a-z][a-z0-9+.-]*):/i.exec(url);
  if (!schemeMatch) {
    return { ok: true, url };
  }

  const scheme = schemeMatch[0].toLowerCase();
  if (SAFE_SCHEMES.indexOf(scheme) === -1) {
    return { ok: false, reason: `Links starting with "${scheme}" aren't allowed. Use a web address (https://…), a mailto: or a tel: link.` };
  }
  return { ok: true, url };
}

/** The URL to render, or '' when it is empty or unsafe. */
export function safeUrl(value: string | undefined): string {
  const result = checkUrl(value);
  return result.ok ? result.url : '';
}

/**
 * Image addresses: web (https/http) and site-relative only. mailto:, tel: and
 * data: make no sense for an image and data: URLs would bloat the page.
 */
export function checkImageUrl(value: string | undefined): UrlCheck {
  const result = checkUrl(value);
  if (!result.ok) {
    return result;
  }
  if (/^(mailto|tel):/i.test(result.url)) {
    return { ok: false, reason: 'Use the web address of an image (https://…) or a link to an image in SharePoint.' };
  }
  return result;
}

/** The image address to render, or '' when it is empty or unsafe. */
export function safeImageUrl(value: string | undefined): string {
  const result = checkImageUrl(value);
  return result.ok ? result.url : '';
}
