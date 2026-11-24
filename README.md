# FEDORAS teaching materials

The FEDORAS teaching materials website gives teachers access to activities and supporting resources for science and sustainability education. It is generated from the [WP2 materials on Nextcloud](https://labs.fedoras-academy.eu/index.php/s/yqWgxFNtA2QEFr6).

The Word documents on Nextcloud are the source materials. This repository contains the code, design and checks used to generate the website. A preview of the first version exists on Netlify.

## Updating an activity

1. Edit the activity’s Word document in the WP2 Nextcloud folder.
2. Keep the existing activity template and update the last-edited date on the cover.
3. Keep the filename unchanged where possible, so existing website addresses remain valid.
4. Use Word headings, numbered lists, hyperlinks, figure captions and alternative text where appropriate. Leave optional fields blank if they are not relevant.
5. Rebuild the website and check the updated page against the Word document.

Each activity page displays the original Word document’s filename and links to the WP2 folder. To download the document, find it in Nextcloud and select **Actions (⋯) → Download**.

## Building the website

The build requires Python 3.10 or later.

Create and activate a virtual environment, install the dependencies, and run the refresh command:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python refresh.py
```

On Windows, activate the environment in PowerShell with:

```powershell
.venv\Scripts\Activate.ps1
```

The refresh command reads the public Nextcloud collection and creates:

- `site-output/` — the generated website
- `review-report.json` — a build report

It does not change files in Nextcloud or publish the website.

To run another build without replacing existing output, choose new paths:

```bash
python refresh.py --output preview-new --report preview-new-report.json
```

To view the generated website locally:

```bash
python -m http.server 8000 --directory site-output
```

Then open `http://localhost:8000`.

The generated website does not need to contact Nextcloud to display pages or search. The two featured PDFs are included in the generated site; Word documents remain on Nextcloud.

## Changes to source files

The build checks the source collection against `approved-file-list.json`. If a previously listed file is missing or renamed, the build stops for review.

For an intentional rename:

1. Check whether the activity’s existing website address needs to be preserved using the alias logic in `parse.py`.
2. Update `approved-file-list.json`.
3. If a featured PDF was renamed, update its path in `build.py`.
4. Rebuild and check the result.

New activities with unique page addresses are included when the site is rebuilt. Update the reviewed file list after adding them.

## Reviewing the generated site

Compare changed pages with their Word originals. Check headings, lists, links, worksheets, figures and equations, including on a phone-sized screen.

The Word documents remain the authoritative source. The website presents their content for continuous reading and does not reproduce Word formatting exactly.

## Repository contents

| File or folder | Purpose |
|---|---|
| `refresh.py` | Retrieves the source materials and builds the site |
| `snapshot.py` | Reads the public Nextcloud collection |
| `build.py` | Generates the website |
| `reader.py`, `parse.py`, `equations.py` | Convert Word content and equations for web reading |
| `facets.py` | Maps subject names used in the activity filters |
| `site-copy.json` | Stores website text sourced from the FEDORAS Academy |
| `site.css`, `site.js` | Define the site’s appearance and interactions |
| `approved-file-list.json` | Lists the expected source files |
| `assets/eu-funding.png` | Official EU funding emblem |
| `tests/` | Contains website checks |

## Running the checks

Install the test dependencies and Chromium:

```bash
python -m pip install -r requirements-test.txt
python -m playwright install chromium
```

Create a snapshot, build the site and run the checks:

```bash
python snapshot.py snapshot-local
python build.py snapshot-local test-output test-report.json
python tests/check_site.py test-output snapshot-local test-report.json
```

The checks cover website paths, activity structure, source links, equations, featured PDFs and browser interactions. They do not replace a content review against the original documents.
