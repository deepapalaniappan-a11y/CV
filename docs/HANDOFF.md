# Handoff to Claude Code, 19 May 2026

Use this file to give Claude Code a clear instruction set when pushing the latest web CV to GitHub.

## What changed in this revision

- New portrait photo added to the hero (modest size, 140 px in the sidebar). File: `assets/deepa-photo.jpeg`.
- Lede rewritten:
  - Opens with hands-on experience across grassroots, government, university and international settings.
  - "EquitableDocs" now described as an Accessible Document Initiative, with a link.
  - "Disability Rights Repository" linked and described as a community archive.
  - Bihar paragraph expanded: hands-on work with Self Help Groups for Persons with Disabilities, with the policy draft linked to academia.edu.
  - Snehadhara Foundation advisory line moved to the end of the paragraph, no longer styled as emphasis.
- Hero sidebar: "Code, github.com / CV" replaced with "Online CV, deepapalaniappan-a11y.github.io / CV". Email switched to deepa.palaniappan@graduateinstitute.ch. Phone kept. "Scroll to contact" link removed.
- Footer: same email update applied.
- Section heading changes:
  - "What I am working on this week" replaced with "Current updates."
  - "Currently." replaced with "Active engagements and committee work."
  - "Two projects." replaced with "Two Initiatives."
  - "Trainings and workshops I have led or facilitated." replaced with "Trainings and workshops led and facilitated."
- Highlights tiles: removed "Earlier work" (Bihar) and "Education" tiles. Added two new focus-area tiles: "Higher Education and Disability Inclusion" and "Accessibility and Assistive Technology". Certification tile retitled to "RCI Licence, CPACC, and NISH AT Certification".
- Initiatives section: full content width (no left gutter). Two equal-width cards for the Repository and EquitableDocs, plus a third full-width feature card for the Accessible Documents Community of Practice, with a link to the Google interest form.
- Convening section: added a March 2026 jury entry for the KIHT and AMTZ Hackathon.
- Training and certifications: added RCI Licence (Government of India) and Document remediation training.

## What to push to GitHub

The deployable file is `index.html` at the repo root. `CV Web.html` is the same content under a friendlier filename for downloads.

```
index.html
CV Web.html
assets/deepa-photo.jpeg
README.md
UPDATING.md
HANDOFF.md
print/A-policy-brief.html
print/C-modern-quiet.html
CV print explorations.html
```

The print HTML files in `print/` were not touched in this revision. If you also want to bring them in line with the web CV (photo, email, Bihar paragraph), open a separate task.

## Suggested commit message

```
Refresh web CV, 19 May 2026

- Add portrait photo
- Rewrite lede with hands-on experience and initiative links
- Update email to graduateinstitute.ch
- Rework highlights tiles (drop earlier-work and education tiles,
  add focus-area tiles for Higher Education and for Accessibility
  and Assistive Technology)
- Initiatives section: full-width, add Document Accessibility CoP card
- Convening: add KIHT/AMTZ March 2026 jury entry
- Training: add RCI Licence and document remediation entries
- Replace section headings that read as first-person/casual
```

## Verifying after push

1. Wait one minute for GitHub Pages to redeploy.
2. Open <https://deepapalaniappan-a11y.github.io/CV/> and check:
   - Portrait photo loads (path `assets/deepa-photo.jpeg`).
   - Email link in the sidebar opens `deepa.palaniappan@graduateinstitute.ch`.
   - The Community of Practice card spans the full width below the two main initiative cards.
   - Section headings read "Current updates", "Active engagements and committee work", "Two Initiatives".
3. Test dark mode toggle and the A / A+ / A++ size buttons.
