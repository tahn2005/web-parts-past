Tests for Accordion Cards (v1.0.4).

test.js loads the real production bundle (accordion-cards/dist) in Chromium
inside a stand-in SharePoint host (host.js). It covers the package (names,
version, no full-width sections), the property pane, uploads, editing on the
page, saving, security, layout, keyboard and accessibility.
Results: ../test-report.txt (133/133).

combined-page-test/ puts Accordion Cards and Custom Width Header together on
realistic SharePoint-style pages and audits whole pages: heading outline, one
H1, unique ids, landmarks, contrast, reflow.
Results: ../combined-page-report.txt (32/32).

The scripts use absolute paths from the machine where they were written, so
update them before running elsewhere. They need Playwright, axe-core, pngjs and
esbuild. Generated files (vendor.js, test images) aren't included.
These tests complement checking in a real tenant; they don't replace it.
