# Link Cards (SPFx)

A SharePoint web part that shows a grid of link cards. Each card has its own
text, icon, icon color and link address.

| | |
|---|---|
| SPFx | 1.23.2 (gulp toolchain), React 17 |
| Node | 22.14+ (22.x) |
| Package | `sharepoint/solution/link-cards.sppkg` (shows as **Link Cards** in the App Catalog) |

## Using it

Add a one-column (or multi-column) section to a page and drop in **Link Cards**
(in the *Advanced* group). It starts with one card. Link Cards isn't available
in full-width sections.

**On the page (edit mode)**, each card has a small toolbar: *Edit*, *Move
earlier*, *Move later* and *Delete*. The dashed **Add a link card** tile adds a
card. Selecting a card opens its settings in the property pane.

**Each card:**
- Link text.
- Link address, and whether it opens in a new tab (on by default). Leave the
  address blank to show the card as plain text, for example while the page it
  will point to is being prepared.
- Icon, from a searchable picker.
- Icon color: the same as the other cards, or a custom color.
- Move, duplicate or delete.

**All cards:**
- Icon color scheme: Blue, Indigo, Green, Amber or any custom color. Each
  scheme includes the soft hover tint and border that go with it.
- Text color and card background.
- Columns: Automatic (1 or 2), Always 1, Up to 2 or Up to 3. Columns follow the
  web part's own width, so the layout fits any section.
- Font: Inter, or SharePoint's own font.

## Behavior

- The whole card is the link. On hover and keyboard focus it takes on its tint
  and border.
- Links that open in a new tab use `noopener noreferrer`.
- A card with no address is plain text, not a link, and it's skipped by the
  keyboard.
- Readers never see empty cards: a card needs text or a working address.
- A link with no text is labeled with its address, so it always has a name.

## Accessibility (WCAG 2.2 AA)

- The cards are a list, so screen readers announce how many there are.
- Each link's accessible name is its visible text. Links that open a new tab
  add "(opens in a new tab)".
- Every link has a visible 2px focus outline with at least 3:1 contrast against
  both the card and the page.
- Every link is at least 44px tall.
- Colors are checked automatically: text needs 4.5:1 and icons 3:1, at rest and
  on hover. Any color that falls short is darkened (or lightened on dark cards)
  just enough to pass, and the property pane lists the change. Colors that
  already pass are left exactly as chosen.
- Respects reduced-motion settings, and keeps borders and focus visible in
  Windows High Contrast.
- No sideways scrolling at 320px wide (WCAG 1.4.10 reflow).
- Unsafe addresses (`javascript:`, `//host`, `data:`) are rejected in the
  property pane and never rendered.

## Build and deploy

```bash
npm install
npm run package
```

1. In the SharePoint admin center, go to **More features > Apps > App Catalog**.
2. Upload `link-cards.sppkg`.
3. Tick **Make this solution available to all sites in the organization**, then
   choose **Deploy**.

It only appears on SharePoint pages; Teams isn't supported.

## Performance and privacy

- Readers download one file of about 51 KB. The editing tools load only when an
  author opens the property pane.
- The web part makes no SharePoint or other API calls and needs no permissions.
- PnP's usage reporting is turned off. The only outside request is the optional
  Inter font from Google Fonts.

## Adding icons

Import the icon from `lucide-react` and add it to `ICONS` in
`src/webparts/linkCards/icons.ts`.
