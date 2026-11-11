# Accordion Cards (SPFx)

A SharePoint web part built from the accordion sections and resource cards of
the Manager Hub React app (`../src`). It's its own solution with its own
package, separate from Custom Width Header.

| | |
|---|---|
| SPFx | 1.23.2 (gulp toolchain), React 17 |
| Node | 22.14+ (22.x) |
| Package | `sharepoint/solution/accordion-cards.sppkg` (the original package name, so it overwrites the deployed one) |
| Web part name | **Accordion Cards** (toolbox) |
| Version | 1.0.5 |

## Web parts

### Accordion Cards

The accordion section and its resource cards from the lifecycle pages, as **one**
web part. SharePoint can't nest web parts, so the cards live inside this web part
and are edited the way you'd edit separate web parts: select a card on the page
and the property pane switches to that card's settings.

**Putting it on a page:** add a one-column section, then add **Accordion
Cards** (in the *Advanced* group of the toolbox). It starts with placeholder
header text and one card, ready to fill in. Add one web part for each accordion
section. It can't be placed in full-width sections; those are for the Custom
Width Header.

**On the canvas (edit mode)** each card has a toolbar: *Edit*, *Move earlier*,
*Move later*, *Delete*. The dashed **Add a resource card** tile adds a card. The
section opens automatically while editing, but you can still collapse it.

**Section settings** (property pane, with no card selected):

- **Section header:** the title and description; an icon or an image; whether to
  show the number of cards; whether the section starts expanded for readers; and
  the heading level.
- **Cards:** a button for each card (another way to open its settings), Add card,
  the default link text, and the card background and text colors.
- **Section color** (in the section's own settings, next to its icon): one of the app's four stage colors (Blue, Indigo, Green, Amber) or
  *Custom* for accent, tint and border. You can also set the header background,
  header hover color, title, description, the arrow, and the card area background.
- **Font:** loads Inter as the app does. Turn this off if your tenant blocks
  Google Fonts; the web part then uses SharePoint's font.
- **Accessibility:** lists any colors that were adjusted for contrast.

**Card settings** (select a card):

- the title and description;
- an icon from a searchable picker, or an image you upload or pick from the site,
  OneDrive, stock images or a link;
- the link address, link text and whether it opens in a new tab;
- an accent color that matches the section or is custom;
- buttons to move, duplicate or delete the card.

Uploaded images are saved to the site's **Site Assets** library, in the
`AccordionCards` folder.

### Rendering and responsiveness

Sizes, spacing, colors and hover effects come from the app's Tailwind values.
The number of card columns depends on **the web part's own width**, not the
browser window: 1 column below 592px, 2 columns from 592px, 3 columns from 968px.
In the app the cards switch at the same sizes, so they look alike in a wide one-column
section and in a narrow column.

### Accessibility (WCAG 2.2 AA)

- Follows the WAI-ARIA accordion pattern: the header is a real heading that wraps
  a `<button>` with `aria-expanded` and `aria-controls`. The collapsed panel uses
  `hidden`, so it leaves the tab order. Enter and Space both toggle the section.
- You can set the heading level. Card titles are one level below the section
  title, and the cards are a list so screen readers announce how many there are.
- Screen readers hear the count badge as "7 resources". Icons and images are
  decorative because the title next to them gives the meaning.
- A screen reader hears each link as "View Resource: Recruiting Process". It
  starts with the text on screen, so voice control works. Links that open a new
  tab say so.
- Every control has a visible focus outline, which the app left out. Hover
  effects also show on keyboard focus. Reduced-motion and Windows High Contrast
  settings are respected.
- Contrast is checked automatically. If a text color, the count, a link or the
  arrow would fall below AA (4.5:1 for text, 3:1 for the arrow and focus
  outlines), it is darkened just enough to pass. Colors that already pass are
  unchanged. With the app's Green and Amber colors, link and count text come out
  slightly darker than in the app, because the app's versions don't meet AA.
- Links only accept `https://`, `mailto:`, `tel:` and site-relative URLs.
  `javascript:` links are never rendered.

### Performance and privacy

- Page readers download one file of about 63 KB. The property pane code (PnP
  color and file pickers, and the icon picker) loads only when an author opens
  the pane.
- PnP's controls normally send usage data to an external endpoint. The web part
  turns that off. The only outside request is the optional Inter font from Google
  Fonts, which you can switch off.

## Configuration

The solution's configuration (solution name, package name, feature, web part
alias, bundle and strings names) is kept identical to the originally deployed
version, so a new build can be uploaded over the existing package in the App
Catalog. Only the web part's display name and the version number change.
Changing the solution name on an update can stop SharePoint registering the
web part, so leave `config/package-solution.json` alone apart from the version.

## Build and deploy

```bash
npm install
npm run package
```

This writes `sharepoint/solution/accordion-cards.sppkg`.

To make it available on every site in the tenant:

1. Open the **tenant app catalog** (SharePoint admin center > More features >
   Apps > App Catalog).
2. Upload the `.sppkg`.
3. In the dialog, tick **Make this solution available to all sites in the
   organization**, then choose **Deploy**.

After that, page authors on any site find **Accordion Cards** in the toolbox.
Nobody needs to add the app to their site first. It only appears on SharePoint
pages; Teams and single-part app pages aren't supported.

The web part only calls SharePoint's REST API to upload images to Site Assets,
which needs edit rights there. That's normal for page authors.

## Adding icons

Import the icon from `lucide-react` and add it to `ICONS` in
`src/webparts/resourceAccordion/icons.ts`. Icons are imported one at a time so
the bundle doesn't include all of lucide.
