# Quick Link Cards web part

A SharePoint Framework (SPFx 1.23) web part that renders the security quick-link card
grid. It appears in the page toolbox like any built-in web part. Page authors manage
it entirely from the web part's settings panel, without writing any code.

## For page authors (no code)

1. Edit the page, click **+**, and search for **Quick Link Cards**.
2. The web part starts with the six standard security cards. Click the **pencil (Edit
   web part)** icon to open the settings panel.
3. **Page 1 – Cards**
   - **+ Add a card** adds a new card at the end.
   - Click a card's name to open its settings: title, description, button text,
     button link, open in new tab, and **Choose image**. You can pick an existing image
     from the site, OneDrive, or Stock images, or upload one from your computer.
     Uploaded images are saved to this site's *Site Assets › QuickLinkCards* folder.
   - **Colors for this card**: turn this on to give that one card its own accent line,
     title, button, button text and background colors.
   - **Move card up / down** and **Remove this card** are at the bottom of each card.
4. **Page 2 – Layout & colors** (the **Next** button at the bottom of the panel): sets the
   number of cards per row, the image width, and the default colors for every card.
5. Republish the page.

Cards stack to two columns, and then to one, automatically when the section is narrow
(one-third columns, phones).

## For the administrator (one-time deployment)

The ready-built package is `sharepoint/solution/quick-links-cards.sppkg`.

1. Open the **SharePoint admin center › More features › Apps › App Catalog**.
   (Creating the tenant app catalog may take a few minutes the first time.)
2. Upload `quick-links-cards.sppkg` and choose **Enable this app and add it to all sites**.
   The package is built with `skipFeatureDeployment: true`, so no per-site install is
   needed. The web part shows up in the toolbox on every site.
3. To ship an update, increase `version` in `config/package-solution.json` (and in
   `package.json`), rebuild, and upload the new `.sppkg` over the old one. Pages that
   already use the web part pick up the change automatically.

No custom scripting is required (`requiresCustomScript: false`), unlike the Modern
Script Editor.

## For developers

Requires Node 22 (`>=22.14 <23`).

```bash
npm install
npm run start    # local workbench: https://aegworldwide.sharepoint.com/_layouts/workbench.aspx
npm run build    # production build + sharepoint/solution/quick-links-cards.sppkg
```

`npm run start` does not need admin rights. `config/local-dev-certificate.js` skips
SPFx's step of installing a trusted certificate, which needs admin rights. Instead, the
first time you start the server, open https://localhost:4321/temp/build/manifests.js
in the browser you test with and choose to continue past the certificate warning.
Skip `npx heft trust-dev-cert`.

| File | Purpose |
| --- | --- |
| `src/webparts/quickLinkCards/QuickLinkCardsWebPart.ts` | Rendering and the settings panel |
| `src/webparts/quickLinkCards/QuickLinkCardsWebPart.module.scss` | Card styles (ported from the original HTML/CSS) |
| `src/webparts/quickLinkCards/QuickLinkCardsWebPart.manifest.json` | Toolbox name/icon and the six default cards |
| `src/webparts/quickLinkCards/loc/en-us.js` | All text shown in the settings panel |

Image picking and color pickers come from
[`@pnp/spfx-property-controls`](https://pnp.github.io/sp-dev-fx-property-controls/).
