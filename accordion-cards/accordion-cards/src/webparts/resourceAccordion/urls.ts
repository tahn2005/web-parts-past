/**
 * Only http(s), mailto, tel and site-relative URLs are rendered. Anything else
 * — notably `javascript:` — is dropped, so a pasted URL can never run script.
 */
export function safeHref(value: string | undefined): string | undefined {
  const url = (value ?? '').trim();
  if (!url || url === '#') return undefined;
  if (/^(https?:|mailto:|tel:)/i.test(url)) return url;
  if (/^\/(?!\/)/.test(url)) return url; // site-relative, but not protocol-relative "//host"
  if (/^[?#]/.test(url)) return url;
  return undefined;
}

/** Images are restricted further: no mailto/tel. */
export function safeImageSrc(value: string | undefined): string | undefined {
  const url = safeHref(value);
  if (!url || /^(mailto:|tel:|[?#])/i.test(url)) return undefined;
  return url;
}

/** Property pane validation message for link fields. Empty string = valid. */
export function validateLink(value: string): string {
  const url = (value ?? '').trim();
  if (!url) return '';
  return safeHref(url)
    ? ''
    : 'Enter a full address starting with https://, a site path starting with /, or a mailto:/tel: link.';
}

export function validateImageUrl(value: string): string {
  const url = (value ?? '').trim();
  if (!url) return '';
  return safeImageSrc(url) ? '' : 'Enter an image address starting with https:// or a site path starting with /.';
}
