# Updating the CV

This project has three CV files that can be updated independently. The **web CV** is the one that should change weekly; the two **print CVs** are the snapshot you give to people.

```
CV Web.html                  <- long-scroll site, updated weekly
print/A-policy-brief.html    <- 3-page A4, institutional / policy context
print/C-modern-quiet.html    <- 3-page A4, university / modern context
CV print explorations.html   <- side-by-side viewer for A and C
```

---

## 1. Add a new activity to the web CV

This is the most common update. Open `CV Web.html` and find the block marked:

```
<!-- ============================================================
     FEED START - edit this section weekly to add activities.
     Statuses: upcoming, this-week, recent, past.
     Keep the most recent / soonest on top. See UPDATING.md.
     ============================================================ -->
```

Paste a new `<article>` at the **top** of `<div class="feed">` (so newest is first). Copy this template and fill it in:

```html
<article data-date="YYYY-MM-DD" data-status="upcoming">
  <time><b>Date or range</b>Short qualifier</time>
  <div class="body">
    <h3>Role, Title of the event or programme</h3>
    <p>One or two sentences describing what you are doing.</p>
    <p class="where">Venue, City</p>      <!-- optional -->
  </div>
  <span class="status">Upcoming</span>
</article>
```

### Status values

| `data-status`  | When to use                     | Pill colour     |
|----------------|---------------------------------|-----------------|
| `upcoming`     | scheduled but not yet happened  | warm tan        |
| `this-week`    | happening this week             | green           |
| `recent`       | done in the last few weeks      | blue            |
| `past`         | older, still worth surfacing    | grey            |

When an activity finishes, change its `data-status` from `upcoming` to `recent`, then later to `past`. Old entries can be pruned as the list grows.

### Worked example

You shared this email:

> MQCI is organizing the Accessible Built Environment Certification Course on 28-29 May 2026 at KCC, AMTZ.

That became:

```html
<article data-date="2026-05-28" data-status="upcoming">
  <time><b>28 - 29 May 2026</b>Two days</time>
  <div class="body">
    <h3>Subject Matter Expert, Accessible Built Environment Certification Course</h3>
    <p>MQCI, Andhra Pradesh MedTech Zone. Two-day certification course on accessible built environment, contributing as SME.</p>
    <p class="where">Kalam Convention Centre (KCC), AMTZ, Visakhapatnam</p>
  </div>
  <span class="status">Upcoming</span>
</article>
```

---

## 2. Three ways to actually make the edit

Pick whichever feels easiest. All three end up changing the same HTML file.

### A. Through this chat

Send a message like:

> Add a new activity to the web CV.
> Date: 12 June 2026
> Role: Keynote
> Event: National Conference on Inclusive Higher Education
> Where: Symbiosis, Pune
> Status: upcoming

I will paste the corresponding `<article>` at the top of the feed and re-share the file.

### B. Through Claude Code (on your laptop)

1. Open Claude Code in the folder that has this project.
2. Ask:

   > In `CV Web.html`, add an `<article>` to the feed for the new activity below, at the top of the `<div class="feed">` block. Status: upcoming. <details...>

Claude Code will edit the file in place. Commit and push.

### C. Through GitHub directly

1. Open the repository on github.com.
2. Click `CV Web.html`, then the pencil icon to edit.
3. Find the `<!-- FEED START -->` marker, paste a new `<article>` at the top of `<div class="feed">`, fill in fields.
4. Commit with a message like `Add MQCI cert course, 28-29 May 2026`.

GitHub Pages will pick up the change automatically.

---

## 3. Other common edits

| Change                          | File and roughly where                                          |
|---------------------------------|-----------------------------------------------------------------|
| Email or phone                  | `CV Web.html`, footer block marked `<!-- FOOTER -->`            |
| Photo                           | Replace the `.photo` div in the hero with `<img src="...">`     |
| A new appointment in "Currently"| `CV Web.html`, section `#currently`; print A and C "Currently"  |
| A new publication               | `CV Web.html` section `#publications`, then print A and C       |
| A new fellowship or award       | Same pattern, section `#fellowships`                            |
| Updated lede / summary          | Hero `<p class="lede">` in `CV Web.html`; lede in both prints   |

For the print PDFs: open the print file in a browser and press **Cmd/Ctrl + P**. Choose "Save as PDF". The `@page A4` rules already set the size correctly.

---

## 4. SEO note

The web CV already has:

- `<title>` and `<meta description>` tuned for "Deepa Palaniappan disability inclusion accessibility".
- Open Graph tags so the link previews cleanly when shared.
- JSON-LD `Person` schema so search engines understand the page.

When you make a major change, also bump the footer's `<time datetime>` field so search engines see the page is being maintained.

---

## 5. If you break something

Every change is just an HTML edit. Worst case: revert the file (GitHub Desktop, `git checkout`, or send me the file and I will fix it).
