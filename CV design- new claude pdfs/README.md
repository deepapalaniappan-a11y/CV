# Deepa Palaniappan, CV

Live: <https://deepapalaniappan-a11y.github.io/CV/>

This repository hosts a long-scroll web CV plus two A4 print-ready CVs. The web CV is the one updated weekly; the print files are snapshots you give to people.

## Files

```
index.html                   The web CV (also served as CV Web.html, identical)
CV Web.html                  Same file, friendlier filename for downloads
print/A-policy-brief.html    3-page A4, institutional / policy / UN / Govt context
print/A-policy-brief.docx    Same content as A in Word format (editable, basic styling)
print/C-modern-quiet.html    3-page A4, university / academic / modern context
print/C-modern-quiet.docx    Same content as C in Word format (editable, basic styling)
CV print explorations.html   Side-by-side viewer for A and C
UPDATING.md                  How to add new activities to the feed
```

The **.html print files** are the primary deliverables; they produce pixel-perfect PDFs via "Save as PDF" (button in the bottom-right corner of each page).
The **.docx files** are a fallback if you ever need to edit in Word or share an editable version.

## Hosting on GitHub Pages

The repository is already configured. Push any change to the default branch and the new version is live within a minute.

1. Settings, Pages.
2. Source: Deploy from a branch.
3. Branch: `main` (or whichever is default), folder: `/ (root)`.

GitHub serves `index.html` automatically.

## Updating

See [`UPDATING.md`](UPDATING.md) for how to add a new activity to the "Now & Next" feed on the web CV, and where to make other common edits.

## Printing the print CVs as PDF

Each print file has a **"Save as PDF"** button in the bottom-right corner.

1. Open `print/A-policy-brief.html` or `print/C-modern-quiet.html` in a browser.
2. Click **"Save as PDF"** (or press `Cmd / Ctrl + P`).
3. Choose "Save as PDF" as the destination. The page styles already set A4, no margins, and clean page breaks. Just hit Save.

The button is hidden when actually printing, so it never appears in the saved PDF.

## Accessibility

The web CV ships with:

- Skip-to-content link
- Semantic landmarks (`<header>`, `<main>`, `<section>`, `<footer>`)
- WCAG AA contrast in both light and dark mode
- Visible focus rings on every link and button
- A dark-mode toggle and three text-size options, both persisted
- `prefers-reduced-motion` honoured (smooth scroll disabled)
- JSON-LD Person schema for search engines

If you spot an accessibility regression, please open an issue.
