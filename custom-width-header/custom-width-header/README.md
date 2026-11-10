# Custom Width Header (SPFx)

A header web part built from the hero banners in the Manager Hub app: the
lifecycle page header and the resource detail header. It's its own solution
with its own package, separate from Accordion Cards.

| | |
|---|---|
| SPFx | 1.23.2 (gulp toolchain), React 17 |
| Node | 22.14+ (22.x) |
| Package | `sharepoint/solution/manager-hub-page-header.sppkg` (shows as **Custom Width Header** in the App Catalog) |

## Using it

Add a **full-width section** to a page and drop in **Custom Width Header**
(in the *Advanced* group). The background then runs edge to edge and the text stays
centred at the app's width. The header also works in any normal section.

## What authors can change

**Text**
- Back link: show or hide it, its text (default "Home"), and its address. Leave
  the address blank to link to the site's home page.
- Label above the title, e.g. "RECRUITING & HIRING".
- Title and description.

**Icon:** none, an icon from the searchable picker, or an uploaded image.

**Background**
- **A color.** Pick one of the four Manager Hub stage colors or any custom
  color. You can add a soft gradient like the app's.
- **An image.** Upload one or pick it from the site, OneDrive, stock images or a
  link. Choose which part stays in view (center, top or bottom) and how the text
  sits on it: a tint over the whole image, or a solid panel behind the text.
- Text color.

**Layout**
- *Page header (wide)*: the app's lifecycle page hero, with content up to 1280px.
- *Content header (narrow)*: the app's resource header, with content up to 880px.
- Height: fit the content, medium (280px) or large (400px).
- Heading level: **Automatic** (the default), H1 or H2.

**Font:** Inter and Merriweather, or SharePoint's own font.

Uploaded images are saved to the site's **Site Assets** library, in the
`CustomWidthHeader` folder.

## Accessibility (WCAG 2.2 AA)

Headers with text on images or colored backgrounds are a common source of
SharePoint accessibility failures. This web part handles that directly.

**Text over images.** The contrast of text on a photo depends on which pixels
end up behind each letter, so automated checkers can't measure it. The web part
calculates the tint against the two worst possible images, pure white and pure
black. If the author's tint isn't strong enough for both, it's raised just
enough and the pane says so. Any real photo falls between those two extremes,
so the text passes on every image, crop and screen size. The tests prove this by
rendering hostile images and measuring the actual pixels behind every word.

For pages that must pass automated checkers cleanly, choose **Put the text on a
solid panel**. The text then sits on a solid color that checkers can measure,
and nothing is left for manual review.

**Colors.** Any background color, gradient or text color is checked. If
something would fail, the web part first raises the opacity of the softer
secondary text, then darkens (or lightens) the background, and only as a last
resort changes the text color. The pane lists every adjustment.

- The default background is **solid**, because checkers can measure solid
  colors but flag gradients for review. The gradient is still an option.
- Only text that's actually shown is checked. A header with just a title (large
  text, 3:1) keeps colors that small text would have needed changed.

**Heading outline.** The title's heading level is chosen automatically by
default. On article pages, SharePoint already shows the page title as the main
(H1) heading above the header, so the header uses H2. On home-style pages there
is no page title, so the header is the H1. Either way the page outline never
skips a level and never has two H1s. Authors can still force H1 or H2; if that
doesn't fit the page, the Accessibility section of the property pane says so.
Headers saved with version 1.0's fixed H1 default switch to Automatic.

**Other:**
- The title is a real heading, and it's never left empty for readers.
- Background and icon images are decorative and hidden from screen readers.
- Every link has a visible focus outline.
- No sideways scrolling at 320px wide (WCAG 1.4.10 reflow).
- Respects reduced-motion settings. In Windows High Contrast mode the photo is
  dropped so text uses the system colors.
- Unsafe link and image addresses (`javascript:`, `//host`, `data:`) are never
  rendered.

## Build and deploy

```bash
npm install
npm run package
```

1. In the SharePoint admin center, go to **More features > Apps > App Catalog**.
2. Upload `manager-hub-page-header.sppkg` (replacing the existing file).
3. Tick **Make this solution available to all sites in the organization**, then
   choose **Deploy**.

It only appears on SharePoint pages; Teams isn't supported.

## Performance and privacy

- Readers download one file of about 52 KB. The editing tools load only when an
  author opens the property pane.
- PnP's usage reporting is turned off. The only outside request is the optional
  Google Fonts stylesheet.
