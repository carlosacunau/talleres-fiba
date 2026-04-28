# talleres-fiba

GitHub Pages site for Fiba Labs workshops.

## Domain
- Production: `talleres.fibalabs.com`
- DNS: Cloudflare CNAME `talleres` → `carlosacunau.github.io` (DNS only, proxy OFF — GH Pages handles SSL)

## Structure
- `index.html` — landing page, lists active cohorts
- `<cohort-slug>/` — one folder per cohort (e.g. `chile-1/`, `chile-2/`, `colombia-1/`)
  - `index.html` — application form
  - `gracias.html` — thank-you page
- `assets/` — shared CSS
- `apps-script/Code.gs` — backend endpoint (deployed separately to Apps Script, not from this repo)

## Form backend
- Apps Script web app receives JSON POSTs from the form
- Auto-creates one tab per cohort in the target Sheet
- Emails Carlos on each submission

## Adding a new cohort
1. `cp -r chile-1/ <new-cohort>/`
2. Edit content for the new cohort
3. Add a card to `index.html`
4. The Apps Script endpoint handles new cohorts automatically (uses the `cohort` hidden field as tab name)

## Source spec
Form questions come from `~/OS/fiba-labs/strategy/workshops/chile-business-owners/INTAKE-FORM.md`.
