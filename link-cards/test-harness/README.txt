Tests for Link Cards (v1.0.1).

test.js loads the real production bundle (link-cards/dist) in Chromium inside a
stand-in SharePoint host (host.js). It covers the package (names, version, no
full-width sections), editing on the page and in the property pane, saving,
security, odd content, layout, keyboard, focus, target size, contrast at rest
and on hover for every color scheme, and Windows High Contrast.
Results: ../test-report.txt (119/119).

combined-page-test/ puts Link Cards together with the other web parts on
realistic SharePoint-style pages and audits whole pages: heading outline, one
H1, unique ids, landmarks, contrast, reflow.
Results: ../combined-page-report.txt (32/32).

The scripts use absolute paths from the machine where they were written, so
update them before running elsewhere. They need Playwright, axe-core and
esbuild. vendor.js (React 17 and lodash) is generated, not included.
These tests complement checking in a real tenant; they don't replace it.
