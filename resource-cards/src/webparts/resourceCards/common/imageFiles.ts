/**
 * Rules for image files authors upload. Kept free of SharePoint imports so
 * they can be unit tested and used by the property pane directly.
 */

export const MAX_IMAGE_BYTES = 5 * 1024 * 1024;

/** File types browsers display reliably in an <img>. */
export const ACCEPTED_IMAGE_TYPES: { [mime: string]: string } = {
  'image/png': 'png',
  'image/jpeg': 'jpg',
  'image/gif': 'gif',
  'image/webp': 'webp',
  'image/svg+xml': 'svg'
};

export const ACCEPT_ATTRIBUTE = Object.keys(ACCEPTED_IMAGE_TYPES).join(',');

/** Returns a message explaining why the file can't be used, or '' if it can. */
export function validateImageFile(file: { name: string; type: string; size: number }): string {
  if (!ACCEPTED_IMAGE_TYPES[file.type]) {
    return `"${file.name}" isn't a supported image. Use a PNG, JPG, GIF, WebP or SVG file.`;
  }
  if (file.size > MAX_IMAGE_BYTES) {
    const mb = (file.size / (1024 * 1024)).toFixed(1);
    return `"${file.name}" is ${mb} MB. Images can be up to 5 MB; try a smaller or compressed version.`;
  }
  if (file.size === 0) {
    return `"${file.name}" is empty.`;
  }
  return '';
}

/**
 * Builds a safe, unique file name: SharePoint rejects characters such as
 * # % * : < > ? / \ | and names that start or end with a dot or space, and a
 * timestamp keeps two uploads of "logo.png" from overwriting each other.
 */
export function uploadFileName(originalName: string, mimeType: string, now: Date = new Date()): string {
  const extension = ACCEPTED_IMAGE_TYPES[mimeType] || 'png';
  const base =
    originalName
      .replace(/\.[^.]*$/, '')
      .normalize('NFKD')
      .replace(/[^\w\- ]+/g, '')
      .trim()
      .replace(/\s+/g, '-')
      .slice(0, 60) || 'image';
  const stamp = now.toISOString().replace(/[-:T]/g, '').slice(0, 14);
  return `resource-card-${stamp}-${base}.${extension}`;
}
