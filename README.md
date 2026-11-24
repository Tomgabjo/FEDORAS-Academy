# FEDORAS teaching materials — Reading layout

This repository builds a static website from the [approved public Nextcloud share](https://labs.fedoras-academy.eu/index.php/s/yqWgxFNtA2QEFr6). **Edit the Word originals in Nextcloud, not the generated pages.** GitHub holds the build code and handoff instructions. Netlify hosts a built preview of the site; a commit in GitHub does not, by itself, publish a change.

## For researchers: update an activity

1. Update the approved `.docx` in Nextcloud. Keep the standard activity template, update the visible cover date, and keep the filename stable so existing URLs still work.
2. Use Word's real headings, numbering, hyperlinks, figure captions and alternative text where possible. Avoid adding instructions for other editors inside the approved documents. Leave optional fields blank rather than leaving unmarked option lists. If a new URL or a supporting file is needed, add it to the approved share and link it from the activity.
3. Tell the person maintaining the site it is ready for a rebuild. Check the preview against the Word original. Each activity’s Word card names the exact file and opens the public WP2 folder. Download the named file from its **Actions** menu in Nextcloud. The publicly shared folder was verified in the browser; a standalone single-file viewer link was not identified. Raw WebDAV links are used for the build only, not as public site download buttons, after repeat reports of browser failure. Word files are not copied into the deploy.

## For the site maintainer

Install Python 3.10+ and use a virtual environment. Activate it before running these commands (on Windows, activate `.venv\Scripts\Activate.ps1` in PowerShell).

```sh
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python refresh.py
```

`refresh.py` takes a fresh, read-only snapshot of the public share, builds and checks it in a temporary directory, and leaves `site-output/` plus a separate `review-report.json`. It does **not** publish. To run again, use new output and report paths, e.g. `python refresh.py --output preview-new --report preview-new-report.json` (add these to your local excludes or keep them outside the repo).

Preview the output with `python -m http.server 8000 --directory site-output` and open `http://localhost:8000`. The website has no runtime connection to Nextcloud for rendering or search; visiting an original download opens the source on Nextcloud. The two featured PDFs are copied, unmodified, into the Netlify bundle so the browser can download them directly.

If an existing source file is missing or renamed, the build stops against `approved-file-list.json`. Investigate before updating this list. For a legitimate rename, preserve the older URL through the alias logic in `parse.py`, update the baseline and confirm any featured download path in `build.py`. A failed check does not affect the live site. New activities with unique slugs are picked up automatically after a rebuild; update the reviewed inventory baseline for the next refresh.

### Review and deploy

1. Review the separate report and check the changed activities on a phone-sized screen, including lists, links, worksheets and figures. For new template features or equations, check the rendered page alongside Word. The converter is for linear web reading, not pixel-for-pixel Word reproduction.
2. To upload manually, open Netlify Drop at <https://app.netlify.com/drop>. Drop the **unzipped folder that directly contains `index.html`** for a new preview, or open the intended site's **Deploys** page and use its manual upload area when you mean to replace that site. Check the preview URL and keep the last working deploy.
3. Before first public use, confirm the official funding lockup and disclaimer against the signed grant agreement and confirm any remaining content and image rights with the consortium. Do not upload the internal review report, working Word snapshot, or this source repo to Netlify.

### Code and copy

- `site-copy.json`: short public overview and Slidebook link label quoted from the [Academy’s repository announcement](https://www.fedoras-academy.eu/futures-are-in-the-making-the-fedoras-materials-repository/). The About page uses excerpts from the Academy project and Resources pages identified there. Use exact published wording for changes. The grant identification line is the supplied project wording.
- `reader.py`, `equations.py`, `parse.py`: extract and convert Word content. Cover art and empty template prompts are skipped; original links, figures, source dates and native equation content are checked. Text can move into source-labelled sections; the Word document remains definitive, especially for typography, numbering, and Word-specific layouts.
- `facets.py`: explicit map of subject spellings from the Reading pilot. No subject, intent or learning objective is inferred from prose. Unmarked option lists are not tags.
- `site.css`, `site.js`, `build.py`: display and filter behaviour. Search data loads only when a teacher searches. Static HTML remains readable without JavaScript.
- `assets/eu-funding.png`: unmodified official horizontal English RGB Monochrome “Funded by the European Union” lockup, from the [Commission download centre](https://ec.europa.eu/regional_policy/information-sources/logo-download-center_en). The 1.2 cm image height accounts for transparent padding, leaving the flag itself approximately 1.03 CSS cm high. The requested grant line and the [EACEA disclaimer](https://www.eacea.ec.europa.eu/grants/visual-identity/visual-identity-programming-period-2021-2027/european-flag-emblem-and-multilingual-disclaimer_en) are separate text at the end of every page.

### Reproduce the full check

```sh
python -m pip install -r requirements-test.txt
python -m playwright install chromium
python snapshot.py snapshot-local
python build.py snapshot-local test-output test-report.json
python tests/check_site.py test-output snapshot-local test-report.json
```

The checks verify internal paths, activity structure, original download URLs, embedded formulas, source PDF checksums, and browser interaction at phone and desktop widths. Tests tied to specific activities serve as regression checks; review intended changes to those inputs rather than disabling the assertions. The checks do not certify the editorial accuracy of the Word originals or third-party links, or confirm grant compliance on behalf of the consortium.
