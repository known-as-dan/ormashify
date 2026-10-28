# ormashify

A standalone, browser-only formatter for native SmartView Classic Basic Word
reports. No assistant, account, upload service, build step or backend is required.

## Use

1. In SmartView Classic, export a Basic report as Word (.docx), one image record
   per page. Include the thermal and visible-light images, image details and both
   original marker tables. Omit separate cover/contents/summary/histogram pages.
2. Open ormashify in a current browser and choose that file.
3. Enter the full site name, system code and inspection date.
4. Build, inspect the page preview and download the branded PDF.

Original photo timestamps and camera markings remain. No new measurement points,
image blending, temperature calculations or thermal interpretation are performed.
Source notes are retained; long tables/notes continue onto additional branded
pages. Unsupported layouts, external images and substantive headers/footers stop
formatting instead of being silently omitted. The source DOCX is never modified.

## GitHub Pages

Copy this folder's contents to the root of the intended Pages repository. Keep
`assets/`, `vendor/` and `.nojekyll`. In repository Settings → Pages, select
Deploy from a branch, then `main` and `/ (root)`. All paths are relative, so a
project URL such as `https://OWNER.github.io/ormashify/` works unchanged.
Only publish this app folder. The surrounding Thermalite repository contains
separate tools and private local job material and is not the deployment bundle.

For local use, serve this directory over HTTP, e.g. `python -m http.server 8786`.
Opening `index.html` directly as a file is not supported by module/asset loading.

## Dependencies

Vendored locally: pdf-lib 1.17.1, @pdf-lib/fontkit 1.1.1, fflate 0.8.2,
bidi-js 1.0.3, PDF.js 6.4.299, and Noto Sans / Noto Sans Hebrew fonts.
See `vendor/`, `assets/OFL.txt`, `THIRD-PARTY.md` and `vendor-manifest.json`.
No CDN requests, analytics, local-storage copies of reports or external APIs.

This is a separate formatter implementation; it does not execute the Windows
SmartView host or replace the advanced reviewed reporting workflow.
