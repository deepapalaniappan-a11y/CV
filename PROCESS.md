# How I keep my CV up to date

A short guide to my own system. Read this if I have forgotten how it works.

## The one rule

**After every event, the same day, add one line to the INBOX in `MASTER-CV.md`.**

That is the whole discipline. Formatting can wait. Remembering cannot.
Example INBOX line:

```
- 2026-06-12 Keynote, National Conference on Inclusive Education, IIT Madras.
```

## The system in one picture

```
MASTER-CV.md   ->   build.ps1   ->   Word .docx
(I edit only                          PDF
 this file)                           docs/index.html  (public web CV)
                                      docs/resume.json (portable standard)
```

`MASTER-CV.md` is the single source of truth. I never hand-format a CV again.

## Routine: after an event

1. Open `MASTER-CV.md`. Add a dated line under `# INBOX`. Save. Done for now.
2. When I have a few minutes (or before applying anywhere), move the line out of
   INBOX into the right section (`## PROFESSIONAL EXPERIENCE`,
   `## PUBLICATIONS`, `## CONFERENCE ENGAGEMENT (SELECTED)`, etc.), phrased
   like the lines already there. Delete it from INBOX.
3. Double-click `build.ps1` (or run `./build.ps1`). It rebuilds everything.
4. If publishing the web CV: open GitHub Desktop, type a short summary
   ("Add IIT Madras keynote"), click Commit, then Push.

Claude can do steps 2 to 4 for me. I just say which event and Claude tidies the
line, runs the build, and tells me what changed.

## Format reminders for MASTER-CV.md

- `## SECTION TITLE` starts a section. Sections render in the order they appear.
- Inside `## PROFESSIONAL EXPERIENCE`, a role is:
  `### Title — Organisation @@ dates`  then its bullets under it.
- Wrap a lead-in in `**double stars**` to make it bold.
- The `# INBOX` section is never rendered into any CV. It is my scratchpad.

### Visibility markers (this is the clever bit)

One file makes a tight PDF AND a full web CV. The first character of a line
or a role heading decides where it shows:

| Marker | Shows in PDF (3 pages) | Shows on web (full) | Use it for |
|--------|:---:|:---:|---|
| `- `   | yes | yes | normal content, the curated record |
| `+ `   | no  | yes | extra detail, full record only |
| `~ `   | yes | no  | condensed summary lines for the PDF |
| `> `   | yes | no  | small italic note (e.g. "available on request") |

A role can carry the marker too: `### + Title — Org @@ dates` (web only),
`### ~ Title — Org @@ dates` (PDF only). Bullets under a `+` or `~` role
inherit that visibility, so you do not have to mark each bullet.

Rule of thumb when adding an event: if it is strong and you would put it on a
3-page CV, use `- `. If it is real but minor (one more workshop, an older
role), use `+ ` so it enriches the web CV without bloating the PDF.

## Tailoring for a specific job

Do not bend `MASTER-CV.md` toward one job. Keep it complete and honest.
For a specific application, ask Claude to produce a tailored cut from the
master (like the IIMC version). Tailored files are named `*_TAILORED_*` and are
kept out of the public repo automatically.

## GitHub: first-time setup (once only)

1. Install **GitHub Desktop** (desktop.github.com). Sign in.
2. File > Add Local Repository > choose this `CV` folder. It will offer to
   create a repository here. Accept. Keep it **Public** (needed for free
   GitHub Pages) or Private if I do not want the web CV.
3. First commit: summary "Initial CV system", click Commit, then Publish.
4. On github.com, open the repo > Settings > Pages. Under "Build and
   deployment", Source = Deploy from a branch, Branch = `main`, Folder =
   `/docs`. Save.
5. After a minute the web CV is live at:
   `https://<my-github-username>.github.io/<repo-name>/`
   Put that link in email signatures and LinkedIn.

After that, every update is just: build, then Commit and Push in GitHub Desktop.

## Privacy: what is public

If the repo is Public, everything tracked by git is public, including
`MASTER-CV.md` and therefore its INBOX. So:

- Keep INBOX to plain event lines. It is CV material anyway, not private notes.
- Role-tailored CVs, the old source PDF, Word lock files and scratch images are
  excluded automatically by `.gitignore` and never published.
- If I ever want the INBOX private, keep the repo Private (the free web CV
  needs a Public repo, so this is a trade-off to decide once).

## If something breaks

- `build.ps1` prints the error. Usually a stray character in `MASTER-CV.md`.
- The previous `.docx` and `.pdf` are still there until the next successful
  build, and OneDrive keeps file history. Nothing is lost.
- Ask Claude: "the CV build failed, here is the message" and paste it.
