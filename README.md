# Capability Roadmap

An editable, presentation-ready roadmap for planning software capability delivery
(**Use Cases**, abbreviated **UC**) over time. It reads and writes a
**Comma-Separated Values (CSV)** file compatible with Cameo Systems Modeler tables,
groups work into collapsible swimlanes, colors bars by status, and flags special
use cases — all in a single self-contained web page with an optional local server
for saving edits straight to disk.

There are **no dependencies, no build step, and nothing tied to any specific
machine.** Clone the repository, and anyone can run it against a CSV with the
expected fields.

---

## What it does

- **Swimlanes by `lane`** — work is grouped into collapsible horizontal bands
  (e.g. *Requirements Engineering*, *Model Analysis*, *Model Making*). Collapsing
  lanes you are not looking at is the main lever for keeping ~100 use cases readable.
- **Color by `status`** — each status gets a distinct color with an automatically
  chosen readable text color (black or white, picked from the fill's luminance),
  shown in a legend. Click a legend chip to filter that status in or out.
- **Dan-label** — any row with `dan-label = TRUE` gets a gold ring and star, plus a
  legend entry, to mark special use cases.
- **Live editing** — add, remove, rename, re-prioritize (drag), and edit delivery
  dates, effort, status, and lane directly in the page.
- **Two save modes** — served by the Python companion, **Save** writes back to the
  CSV on disk; opened as a plain file, **Download** produces a CSV instead.
- **Scale controls** — density toggle, zoom, fit-to-width, search, and a dan-only
  filter.
- **Present mode** — a clean, read-only view for stakeholders.

Positioning model: each bar **ends on its `deliveryDate`** and extends left by its
`FTE-weeks`. Team-capacity/bandwidth scheduling is intentionally **not** modeled
yet; this release focuses on due-date-driven planning.

---

## The CSV schema

The tool expects these eight columns (header names are matched case-insensitively,
and you can remap them at import time if yours differ):

| Column | Meaning |
| --- | --- |
| `#` | Unique identifier (**UID**) for the use case — an integer. |
| `Name` | Display name. |
| `priority` | Integer rank. Lower is higher priority. On save this is rewritten to a clean `1..N` reflecting the current order. |
| `deliveryDate` | Target delivery date in `M/D/YYYY` (e.g. `9/18/2026`). The bar ends here. |
| `FTE-weeks` | **Full-Time-Equivalent weeks** of effort. Used as the bar length in weeks. |
| `status` | Categorical status. Drives bar color and the legend. |
| `lane` | Discipline / swimlane the work belongs to. |
| `dan-label` | `TRUE` / `FALSE`. `TRUE` adds the gold special-use-case indicator. |

A ready-to-use example is included: **`UC_dummy_data.csv`**.

When you save, the tool writes back **exactly these eight columns** so the file
round-trips cleanly into your existing pipeline. A separate **Download analysis
CSV** button produces the same rows plus computed `DurationWeeks`,
`ScheduledStartDate`, and `ScheduledFinishDate` columns for inspection.

---

## Running it

You have two independent options. Most people want Option A.

### Option A — Local server (edits save to disk)

This runs a tiny **Application Programming Interface (API)** server, in Python,
using only the standard library. It serves the page and saves your edits back to
the CSV you point it at.

**Requirements:** Python 3.7 or newer. Nothing to install.

**Phase 1 — Get the files together**

1. Put `serve.py` and `capability-roadmap.html` in the same folder (they already
   are, in the repository root).
2. Have a CSV with the fields above ready. You can use the included
   `UC_dummy_data.csv`.

**Phase 2 — Start the server**

3. Open a terminal in that folder.
4. Run:
   ```bash
   python3 serve.py UC_dummy_data.csv
   ```
   (On Windows, use `py serve.py UC_dummy_data.csv` if `python3` is not on your path.)
5. Your browser opens automatically at `http://localhost:8000/`. If it does not,
   open that address manually.

> **Checkpoint:** The page loads showing your use cases in swimlanes. The badge in
> the top bar reads **"Local file: UC_dummy_data.csv"** on a green dot. The top-bar
> button says **Save** (not *Download*).

**Phase 3 — Edit and save**

6. Make changes. The status bar shows **"unsaved changes"** in amber when you have
   edits pending.
7. Click **Save** (or press `Ctrl+S` / `Cmd+S`).

> **Checkpoint:** A toast confirms `Saved to <path> (backup written)`. Your CSV on
> disk now reflects the change, and a timestamped `.csv.<timestamp>.bak` backup of
> the previous version sits next to it.

**Useful flags**

```bash
python3 serve.py plan.csv --port 8080      # use a different port
python3 serve.py plan.csv --no-browser     # do not auto-open a browser
python3 serve.py --help                    # all options
```

If the CSV you name does not exist yet, the tool starts empty and creates the file
on your first **Save**.

### Option B — Static file / GitHub Pages (view and download only)

The web page works with no server at all — useful for a zero-setup share or for
hosting read-only on GitHub Pages.

1. Double-click `capability-roadmap.html` (or host it as a static site).
2. Use **Load sample data**, or **Import** your CSV.
3. Edit freely; **Download** writes a CSV to your browser's downloads folder.

> **Checkpoint:** The top-bar badge reads **"Browser only"** and the button says
> **Download**. This mode cannot write to disk — that is by design, since a static
> web page has no safe way to do so.

---

## Putting it on GitHub

**Phase 1 — Create the repository**

1. Create a new repository on GitHub (empty, no README — you have one).
2. From the project folder:
   ```bash
   git init
   git add capability-roadmap.html serve.py README.md UC_dummy_data.csv \
           gantt-core.js gantt-core.test.js smoke.test.js
   git commit -m "Capability Roadmap: web planner + local CSV server"
   git branch -M main
   git remote add origin https://github.com/<you>/<repo>.git
   git push -u origin main
   ```

> **Checkpoint:** The repository shows all files. A colleague can now `git clone` it
> and run Option A immediately, with no further setup.

**Phase 2 — (Optional) Publish the read-only viewer**

3. In the repository's **Settings → Pages**, set the source to `main` / root.
4. Rename or copy `capability-roadmap.html` to `index.html` if you want it served at
   the site root.

> **Checkpoint:** `https://<you>.github.io/<repo>/` loads the viewer in
> "Browser only" mode. (GitHub Pages is static, so it offers view/import/download,
> not server-side save — that is Option A's job.)

---

## Security notes (for reviewers)

The local server is deliberately conservative, because it can write to your disk:

- **Localhost only.** It binds to `127.0.0.1` by default and is not reachable from
  the network. Binding elsewhere requires the explicit `--host` flag and prints a
  warning.
- **One file, no client-supplied paths.** The server only ever reads and writes the
  single CSV path given on the command line. No path is ever accepted from the
  browser, so there is no path-traversal surface.
- **Cross-origin write protection.** Write requests carrying an `Origin` header from
  anywhere other than the server's own address are refused with `403`
  (**Cross-Site Request Forgery**, or **CSRF**, mitigation for a localhost tool).
- **Bounded input.** Request bodies are capped (8 MB), as are row and column counts.
- **Atomic, backed-up saves.** Each save writes to a temporary file, flushes and
  `fsync`s it, then atomically replaces the target with `os.replace`, after copying
  the previous version to a timestamped `.bak` (last 10 kept). A crash mid-write
  cannot corrupt your plan.
- **No `eval`, no shell, no third-party packages.**

The web page itself has **no external dependencies** — no **Content Delivery
Network (CDN)** calls, no trackers, no remote scripts.

---

## Development and tests

The correctness-critical logic (CSV parsing, date/week math, color contrast) lives
in `gantt-core.js`, a dependency-free module. It is the exact code inlined into the
web page (the build verifies they are byte-identical).

Run the tests with Node.js:

```bash
node gantt-core.test.js     # 42 unit assertions (parsing, dates, contrast, round-trip)
npm install jsdom           # only needed for the DOM smoke test
node smoke.test.js          # 30 headless UI assertions (swimlanes, filters, editing, CSV)
```

The Python server has no third-party dependencies; `python3 serve.py --help`
confirms it loads.

---

## Troubleshooting

| Symptom | Cause / Fix |
| --- | --- |
| Badge says "Browser only" when you expected disk saving | You opened the `.html` directly instead of through `serve.py`. Start the server (Option A) and open the address it prints. |
| `Address already in use` on startup | Another process holds the port. Use `--port 8080` (or any free port). |
| `python3: command not found` | Try `python`, or on Windows `py`. Confirm Python 3.7+ with `python3 --version`. |
| Browser did not open | Open the printed `http://localhost:<port>/` manually, or omit `--no-browser`. |
| Dates look wrong after import | Confirm the `deliveryDate` column is `M/D/YYYY`. Remap columns in the Import dialog if your headers differ. |
| Save says "cross-origin request refused" | You are reaching the server from a different address than it serves. Use the URL the server printed. |

---

## References

1. Cameo Systems Modeler — exporting tables to CSV/HTML/Excel (justifies the flexible
   column mapping): https://docs.nomagic.com/display/CSM2022x/Exporting+data
2. Cameo Systems Modeler — importing from CSV and Microsoft Excel (confirms two-way
   round-trip): https://docs.nomagic.com/display/CRMP2022x/Importing+from+CSV+and+MS+Excel+files
3. Python `http.server` (standard library, used by `serve.py`):
   https://docs.python.org/3/library/http.server.html
4. Web Content Accessibility Guidelines (WCAG) 2.1 — relative luminance and contrast
   (basis for the readable-text color choice):
   https://www.w3.org/TR/WCAG21/#dfn-relative-luminance
