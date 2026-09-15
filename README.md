# HTML Color Reference (HCR) - Web Edition

A modern, standalone zero-dependency web version of Christopher Fazendin's classic 1995 Visual Basic 4 utility for generating HTML body color codes and palettes.

> **Historical Note**: [Featured in PC Magazine (Sept. 10, 1996)](https://books.google.com/books?id=xrHwrGq70eAC&printsec=frontcover&num=100#v=onepage&q=christopher%20fazendin&f=false).
> The original Visual Basic 4 codebase is preserved in the [`legacy-vb4`](https://github.com/cfazendin/hcr/tree/legacy-vb4) branch.

---

## Features

- 🎨 **Faithful Core Controls**: Full RGB sliders (0–255), numeric text inputs, and native color pickers for all 5 classic HTML targets:
  - `Background` (`bgcolor`)
  - `Normal Text` (`text`)
  - `Link Text` (`link`)
  - `Visited Link` (`vlink`)
  - `Active Link` (`alink`)
- 🖥️ **Interactive Live Preview**: Click any element or background in the sample preview window to directly select and customize its color.
- 📋 **One-Click Tag & CSS Generator**: Generates both classic HTML `<body>` tags and modern CSS rulesets with quick copy buttons.
- 💾 **Dual Themes**:
  - **Windows 95 Retro**: Pixel-accurate bevels, classic window chrome, menus, and status bar.
  - **Modern Developer**: Clean dark-mode slate UI.
- 🌈 **Classic Presets**: 1995 Default, Windows 3.1 Hot Dog Stand, 90s GeoCities Neon, Matrix Terminal, Cyberpunk 1997, Solarized, Paper & Ink, and Midnight Ocean.
- 👁️ **WCAG Contrast Calculator**: Real-time accessibility contrast ratio ratings (AA / AAA) against background color.
- 📥 **Export**: Download a live sample HTML file of your color scheme with a single click.

---

## Running Locally

Since this is a client-side zero-dependency web app, you can simply open `index.html` in any modern web browser:

```bash
# Or start any simple HTTP server:
npx serve .
# or Python:
python -m http.server 8000
```
