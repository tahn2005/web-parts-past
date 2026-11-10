import { uploadFileName, validateImageFile, MAX_IMAGE_BYTES } from '../imageFiles';
import { checkImageUrl } from '../url';
import { normalizeCards } from '../model';

describe('validateImageFile', () => {
  it('accepts common image types up to 5 MB', () => {
    ['image/png', 'image/jpeg', 'image/gif', 'image/webp', 'image/svg+xml'].forEach((type) => {
      expect(validateImageFile({ name: 'a', type, size: 1000 })).toBe('');
    });
    expect(validateImageFile({ name: 'a.png', type: 'image/png', size: MAX_IMAGE_BYTES })).toBe('');
  });

  it('explains why other files are refused', () => {
    expect(validateImageFile({ name: 'doc.pdf', type: 'application/pdf', size: 10 })).toContain('isn\'t a supported image');
    expect(validateImageFile({ name: 'huge.png', type: 'image/png', size: MAX_IMAGE_BYTES + 1 })).toContain('up to 5 MB');
    expect(validateImageFile({ name: 'empty.png', type: 'image/png', size: 0 })).toContain('empty');
  });
});

describe('uploadFileName', () => {
  const when = new Date('2026-10-30T21:05:09Z');

  it('makes a unique, SharePoint-safe name with the right extension', () => {
    expect(uploadFileName('Team Logo (final)#2.PNG', 'image/png', when)).toBe('resource-card-20261030210509-Team-Logo-final2.png');
    expect(uploadFileName('photo.jpeg', 'image/jpeg', when)).toBe('resource-card-20261030210509-photo.jpg');
  });

  it('strips characters SharePoint rejects and falls back to "image"', () => {
    const name = uploadFileName('..%*:<>?|.svg', 'image/svg+xml', when);
    expect(name).toBe('resource-card-20261030210509-image.svg');
    expect(/[#%*:<>?/\\|"]/.test(name)).toBe(false);
  });
});

describe('checkImageUrl', () => {
  it('allows web and SharePoint addresses, refuses everything else', () => {
    expect(checkImageUrl('/sites/hr/SiteAssets/logo.png').ok).toBe(true);
    expect(checkImageUrl('https://contoso.sharepoint.com/sites/hr/SiteAssets/a b.png').ok).toBe(true);
    // eslint-disable-next-line no-script-url
    ['javascript:alert(1)', 'data:image/png;base64,AAAA', 'mailto:a@b.c', 'tel:123'].forEach((url) => {
      expect(checkImageUrl(url).ok).toBe(false);
    });
  });
});

describe('normalizeCards', () => {
  it('upgrades cards saved before images existed', () => {
    const [card] = normalizeCards([{ id: 'a', icon: 'Star', title: 'Old card' }]);
    expect(card.mediaType).toBe('icon');
    expect(card.imageUrl).toBe('');
    expect(card.imageAlt).toBe('');
    expect(card.imageSize).toBe('small');
  });

  it('rejects unknown media types and sizes', () => {
    const [card] = normalizeCards([{ mediaType: 'video', imageSize: 'huge' }]);
    expect(card.mediaType).toBe('icon');
    expect(card.imageSize).toBe('small');
  });
});
