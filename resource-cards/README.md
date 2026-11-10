# Resource Cards (SharePoint Framework web part)

A SharePoint version of the resource cards from the original React app. Page authors
can add, edit and arrange them without writing code.

| Web part | What it is |
|---|---|
| **Resource Cards** | The resource cards from the original lifecycle pages (icon or image, title, description, "View Resource →" link), laid out in the same responsive 1/2/3-column grid. Just the cards, with no box around them, so they sit directly on the SharePoint section background. |

| | |
|---|---|
| SharePoint Framework | 1.24.0 (Heft toolchain) |
| React | 18.3.1 |
| Node.js | 22.14+ (22.x), 24.x or 26.x |
| Package | `sharepoint/solution/manager-hub-webparts.sppkg` |

---

## Deploying (SharePoint admin, one-time)

A ready-built package is included at `sharepoint/solution/manager-hub-webparts.sppkg`.

1. Open your tenant **App Catalog** (SharePoint admin center → *More features* → *Apps* → *App Catalog*).
2. Upload `manager-hub-webparts.sppkg`.
3. When prompted, choose **Enable app** (and *Only enable this app* / *Make available to all sites*).
   The package is built with `skipFeatureDeployment`, so it's available on every site straight away.
   It appears as **Resource Cards** in the App Catalog and in Site contents.

**Updating an earlier upload:** upload the new `manager-hub-webparts.sppkg` over the existing one
in the App Catalog (same file name) and choose **Replace**. Pages that already use the web part pick up
the new version automatically; their cards are kept.

All scripts, fonts and styles ship inside the package. Nothing loads from a
third-party CDN.

## Using it (page authors)

1. Edit a page, select **+** in a section, and choose **Resource Cards** (in the *Resource Cards* group).
   It starts with three sample cards.
2. **Add a card:** select **Add card** above the cards, or **Add card** in the property pane.
3. **Edit a card:** select the ✏️ button on the card. Its settings open in the property pane on the right:
   - **Title**, **Description**, **Link text** (default "View Resource") and **Link URL**
   - **Open link in a new tab**
   - **Show at the top of the card**: *An icon* or *My own image*
     - **Icon**: search by name or topic (for example "money" or "people") and select one
     - **Image**: select **Upload image** to pick a PNG, JPG, GIF, WebP or SVG (up to 5 MB) from your computer.
       It's saved in the site's **Site Assets** library, where SharePoint keeps page images.
       Or paste a link to an image that's already in SharePoint or on the web.
       Then choose a size (*Small* matches the icons; *Medium* and *Large* are 48 and 64px),
       and add a description (alt text) only if the image shows something the card title doesn't already say
   - **Card accent colour**: keep *Same as the web part accent colour*, or pick a different colour for this card only
   - **Move earlier / Move later / Delete card**, and **Done** to return to the general settings
4. **Reorder or delete:** use ‹ › and 🗑 on each card, or the buttons in the card's settings.
5. **General settings** (the web part's ✏️ in the SharePoint toolbar):
   - **Accent colour** for all cards: the five preset colours, or any custom colour by hex code or colour picker
   - **Font**: the Resource Cards font (Inter) or the site's own font
   - **Card title heading level** (see Accessibility below)

Changes show on the page as you type. They're saved when you save or republish the page.

Tips:
- Uploading an image needs permission to add files to the site's Site Assets library (site members usually have it).
  Without it, the pane says so; paste a link to an image instead. People viewing the page need read access
  to wherever the image is stored, so prefer images on the same site.
- If an image is later deleted or moved, readers see the card's icon instead of a broken image,
  and authors see a note on the card in edit mode.
- A card with no URL shows no link on the published page. In edit mode it shows a reminder.
- The layout follows the web part's width. Cards show three across in a one-column section,
  two in a two-thirds or half-width column, and one in a one-third column or on phones.
- Resource Cards isn't offered in **full-width** sections (the edge-to-edge section layout);
  use a one-column section for the widest layout.
- Links can be full addresses (`https://…`), site-relative paths (`/sites/hr/SitePages/x.aspx`),
  `mailto:` or `tel:`. Anything that could run script (such as `javascript:`) is refused.

## Accessibility

The web part meets WCAG 2.2 AA. It was checked with axe-core (no violations in
view mode, edit mode or the property pane controls, across every preset colour
and a very light custom colour) and by keyboard.

Fixes compared with the React original:

| Issue in the React app | What the web part does |
|---|---|
| "View Resource" text failed contrast on the green (3.6:1) and amber (3.1:1) pages | Link and description text are kept at 4.5:1 or better automatically. Light accent colours are darkened for text only; icons keep the exact colour. The property pane shows when this happens. |
| `outline: none` on the link button, so keyboard focus was invisible | A clearly visible focus ring. Keyboard focus also highlights the card the same way hover does. |
| Every link was named just "View Resource" | Each link is announced with its card title ("View Resource: Interviewing"), plus "(opens in a new tab)" where relevant. |
| Cards were plain `div`s with an `h4` | A real list (screen readers announce "list, 7 items"). The heading level is configurable so it fits the page outline. |
| Small click target | The link hit area is 24px tall (WCAG 2.5.8) with no change to the layout. |
| Hover lift animation for everyone | Turned off for people who've asked their system to reduce motion. |
| — | Works in Windows High Contrast (forced colours) and right-to-left languages. |

Edit-mode controls are also accessible. The toolbar buttons have descriptive
names ("Delete card: Interviewing"). Focus moves sensibly after adding, moving
or deleting a card. Each change is announced to screen readers. The icon and
colour pickers are standard radio groups, so arrow keys work.

**Heading level:** card titles look the same at every level. Choose the level
one below the heading above the cards. For example, if the section title is a
*Heading 2*, use *Heading 3* (the default).

## Testing

| Check | Result |
|---|---|
| Lint (SPFx ESLint rules) and TypeScript strict compile | Clean, no warnings |
| Unit tests: colour contrast, link and image-address safety, upload file rules (`npm run build`) | 13/13 pass |
| End-to-end on the **production bundle** in a simulated SharePoint page (`tests/e2e`) | 131/131 pass |
| Pixel comparison of the cards with the React app at 1440/1100/700/420px windows | Identical column widths at every size; about 1% of pixels differ (text anti-aliasing only) |
| axe-core (WCAG 2.2 AA + best practice): view, edit, empty state, property pane (icon and image fields), cards with images, all preset and a very light custom colour, on white, neutral, soft, strong and dark section backgrounds | 0 violations |
| `npm audit --omit=dev` (shipped dependencies) | 0 vulnerabilities |
| Package contents | Bundle, strings and fonts only; no third-party CDN; no `eval` |

The end-to-end suite covers: rendering, page CSS not leaking in, font loading,
edit mode, every property pane control, add/move/delete from both the canvas
and the pane, focus management and screen-reader announcements, saving and
reloading, malformed saved data, script injection attempts in every field,
`javascript:`/`data:` links, custom images (upload via SharePoint REST to Site Assets,
file type and size checks, refused uploads, pasted links, sizes, alt text, broken
or deleted images, SVG, remove and switch back to icon), 120 cards, column breakpoints, 320px reflow,
reduced motion, Windows High Contrast, right-to-left, keyboard focus, and
dispose/re-mount.

### Before going live: SharePoint smoke test (about 15 minutes)

The simulator follows SPFx's documented behaviour, but it isn't SharePoint.
Do these checks once on a real site, ideally a test site first:

1. Upload the `.sppkg` to the App Catalog and enable it. It should deploy with no errors.
2. Add **Resource Cards** to a page. Three sample cards appear.
3. Add a card, edit every field in the property pane, change its icon and colour, move it and delete one. Do this from both the card toolbar and the pane.
   Switch a card to **My own image**, upload a picture, and check it appears in the site's **Site Assets** library.
4. **Save as draft**, reload, and check everything stayed. **Republish** and view as a reader.
5. Check the web part in a one-third, half and one-column section (1/2/3 columns), and on a phone. Check that it is **not** offered in a full-width section.
6. Tab through the published page: each link shows a focus ring and opens the right place (new tab where set).

`npm audit` reports findings in the build toolchain (dev dependencies such as
Microsoft's Heft/webpack tooling). None of that code ships in the package.

---

## For developers

```bash
npm install
npm run build      # lint, compile, test, bundle and package → sharepoint/solution/*.sppkg
npm run start      # serve locally for the hosted workbench (https://<tenant>/_layouts/15/workbench.aspx)
```

### Layout

```
src/webparts/resourceCards/
  ResourceCardsWebPart.ts            web part: properties, property pane, add/edit/move/delete
  ResourceCardsWebPart.manifest.json toolbox entry and default (sample) cards
  components/
    ResourceCards.tsx                grid, edit-mode "Add card", focus management, announcements
    ResourceCard.tsx                 one card (mirrors src/app/components/ResourceCard.tsx)
    ResourceCards.module.scss        styles, with the React app's Tailwind values resolved
  propertyPane/
    IconPicker.tsx                   searchable icon radio group
    ImagePicker.tsx                  upload or paste an image, with preview
    AccentColorPicker.tsx            preset swatches, hex and system colour picker, live preview
    reactField.ts                    hosts React components as custom property pane fields
  common/
    model.ts                         card and property types, defaults
    color.ts                         palettes and WCAG contrast handling
    imageFiles.ts                    accepted image types, size limit, safe file names
    upload.ts                        uploads to Site Assets via SharePoint REST
    icons.ts                         the icons on offer (add more here)
    url.ts                           link validation
    fonts.ts                         registers the bundled Inter font
  fonts/                             Inter 400/600 (SIL Open Font License, see Inter-LICENSE.txt)
  loc/                               property pane text
tests/e2e/                           end-to-end tests of the production bundle (see its README)
```

### Fidelity to the React app

The card markup and styles reproduce `ResourceCard.tsx` and the card grid in
`LifecyclePageTemplate.tsx`. The white accordion panel around the grid is
deliberately left out: the web part is just the cards. Tailwind classes and `theme.css` tokens are
resolved to plain CSS (for example `rounded-lg` there is 10px, because the
theme sets `--radius: 0.625rem`). In a side-by-side test at matching widths,
every element lands at the same position and size, to the sub-pixel. The grid
columns break at the same points. The React breakpoints are window-based
(640px/1280px); here they're measured on the web part itself (558px/934px, the
widths the card grid has at those window sizes), so the grid also works in
narrower SharePoint columns.

### Adding an icon

Import it from `lucide-react` in `common/icons.ts` and add a row to `ICONS`.
The `name` is saved with pages, so never rename or remove an existing entry.
