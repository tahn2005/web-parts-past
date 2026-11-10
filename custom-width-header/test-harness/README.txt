Tests for Custom Width Header (v1.0.2).

test.js loads the real production bundle (custom-width-header/dist) in Chromium
inside a stand-in SharePoint host (host.js). It covers the package (names,
version, full-width support), the property pane, uploads, saving, the heading
outline (Automatic level, manual overrides, upgrade from 1.0), security,
layout, keyboard and accessibility. It also measures rendered contrast pixel by
pixel over hostile test images. Results: ../test-report.txt (122/122).

unit.js runs the contrast engine on its own across colors, images and panels.
Results: ../contrast-unit-report.txt (24/24).

combined-page-test/ puts both web parts on realistic SharePoint-style pages and
audits whole pages. Results: ../combined-page-report.txt (32/32).

The scripts use absolute paths from the machine where they were written, so
update them before running elsewhere. They need Playwright, axe-core, pngjs and
esbuild. Generated files (vendor.js, test images) aren't included.
These tests complement checking in a real tenant; they don't replace it.
