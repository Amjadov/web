# Amjad Masoud — Complexity into Impact

A redesigned résumé and portfolio, built with semantic HTML, CSS, vanilla JavaScript, and a locally served copy of Three.js r128.

## Preview

Run `node serve.mjs`, then open http://127.0.0.1:4173. The server exposes only the portfolio assets and the linked résumé PDF.

## Build

Run `node build.mjs`. This validates JavaScript syntax, anchors, asset references, and core content, then copies the deployment files into `dist/`.

## Experience

- Three live 3D compositions: Connect, Architect, and Deliver. Drag to rotate; keyboard arrows rotate and Home resets the view. The pause control and system reduced-motion preference are respected.
- Before/after commerce performance comparison using the figures in the original résumé.
- Expandable project and career details, expertise filters, and professional credentials.
- Responsive layout, keyboard focus styles, skip link, accessible control states, and a static image fallback if WebGL is unavailable.
- The content remains readable without JavaScript. Google Fonts is optional; local system fonts act as fallbacks.

## Content and artwork

Career titles, dates, metrics, credentials, and contact details preserve the original `index.html`. TOGAF Foundation is explicitly marked in progress.

`assets/architecture.png` is original artwork generated with the built-in image generation tool. Art direction: an exploded chrome and smoked-glass enterprise architecture sculpture with a copper core on an obsidian background. The interactive hero is rendered locally with Three.js, with no external texture or model downloads.

Three.js is distributed under its MIT license in `vendor/THREE-LICENSE.txt`.
