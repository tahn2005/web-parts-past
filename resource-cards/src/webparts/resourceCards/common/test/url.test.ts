import { checkUrl, safeUrl } from '../url';

describe('checkUrl', () => {
  it('accepts web, site-relative, mail and phone links', () => {
    ['https://contoso.sharepoint.com/sites/HR/Shared Documents/a.pdf', '/sites/hr/SitePages/Home.aspx', 'page.aspx', '#top',
      'mailto:hr@contoso.com', 'tel:+12137129139', ''].forEach((url) => {
      expect(checkUrl(url).ok).toBe(true);
    });
  });

  it('refuses script and data links, including obfuscated ones', () => {
    // These are the inputs under test, never executed.
    // eslint-disable-next-line no-script-url
    ['javascript:alert(1)', ' JavaScript:alert(1)', 'java\tscript:alert(1)', 'data:text/html,<script>', 'vbscript:x'].forEach((url) => {
      expect(checkUrl(url).ok).toBe(false);
      expect(safeUrl(url)).toBe('');
    });
  });
});
