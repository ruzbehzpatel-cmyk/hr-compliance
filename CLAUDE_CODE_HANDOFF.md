# HR Compliance Framework — Claude Code Handoff

**Goal:** port the current desktop-bound HR compliance scanner + dashboard to a cloud architecture that runs autonomously every weekday at 03:06 PT, without depending on the user's laptop being on.

**Owner:** Ruz (ruzbeh.z.patel@gmail.com)
**Current state:** working locally inside Cowork as a scheduled task. Production-ready dashboard, link verification, sidecar-aware file handling.
**Target state:** GitHub Actions cron + Cloudflare R2 (xlsx storage) + Cloudflare Pages (hosted dashboard) + Anthropic API (Claude reasoning). Cost ~$30–90/mo (Anthropic API only).

---

## 1. What this thing does

A daily, fully-automated regulatory scanner for US HR compliance. Every weekday at 03:06 AM Pacific it:

1. Reads `HR_Compliance_Framework_COMPLETE.xlsx` (~231 regulation rows across 18 tabs).
2. Scans primary federal and state government sources for changes (DOL, EEOC, NLRB, OSHA, IRS, USCIS, Federal Register, Congress.gov, 50 state legislatures + DC, state AGs, state labor/civil-rights agencies).
3. Uses law-firm trackers (Littler, Seyfarth, Ogletree, Fisher Phillips, Jackson Lewis) as tripwires only — never cited as the source.
4. Updates changed rows in place, adds new regs to the right main tab, splits multi-provision bills by sub-domain.
5. Logs every change in a new `Updates YYYY-MM-DD` tab in the workbook.
6. Reviews the `Survey – AI in Employment` tab and re-tiers states (A/B/C) based on AI bill movement.
7. **HEAD-checks every URL** in the workbook (`Gov Source URL`, `Bill/Statute URL` — ~800 URLs), auto-fixes broken links by re-searching the same agency site, logs fixes as CHANGE entries.
8. Regenerates `dashboard.html` (single self-contained HTML file, ~530 KB).
9. Reports a summary to Ruz.

If the previous run was missed (>24h gap), the next run does a full catch-up scan over the missed window using ACTUAL event dates, not the run date.

---

## 2. Source policy (STRICT — do not deviate)

**Tier 1 — Primary federal (always citeable):**
DOL · EEOC · NLRB · OSHA · IRS · USCIS · Federal Register · Congress.gov

**Tier 2 — Primary state (always citeable, 50 states + DC):**
State legislatures · State Attorney General offices · State labor/workforce agencies · State civil-rights / human-rights commissions · State PFML / paid-leave authorities (where they exist as quasi-public entities like ctpaidleave.org)

**Tier 3 — Tripwires only (NEVER cited):**
Littler · Seyfarth · Ogletree · Fisher Phillips · Jackson Lewis

Workflow: trackers say "X happened" → locate the underlying primary source → only write to the file if the primary verifies. The firm name never appears in `Gov Source URL`.

**OUT of scope (do not add):**
- Local ordinances (city/county) — too volatile
- Advocacy/think-tank trackers (NELP, EPI, NWLC, etc.)
- News media (Bloomberg Law, Law360, HR Dive) — sometimes informal tripwires but never cited
- Wikipedia and user-edited references
- Vendor whitepapers (SHRM commercial, Gartner, etc.)

---

## 3. Operating rules (built up over months — do not break)

### BOTH-TABS rule (added 2026-05-14)
Every confirmed change MUST be recorded in BOTH the date-stamped `Updates YYYY-MM-DD` tab AND the law tab where the regulation should be tracked.
- Existing reg → update in place + log in date-stamped tab.
- New reg → add to appropriate main tab (Federal Laws / State Laws / Additional & Emerging) IN THE SAME RUN as the date-stamped log.
- Multi-provision bill (e.g., CT PA 26-12) → split into multiple rows by sub-domain, each with its own Law ID.

### Edit behavior on existing rows
- Update in place: overwrite changed cell(s), bump `Last Review` to today.
- Do NOT append a dated note to `Notes / Action Items` (the date-stamped tab IS the audit trail).
- Do NOT add "superseded" rows.

### Bundled-emerging rows
Leave bundled rows alone. Split only when one sub-item becomes material (enacted/withdrawn/status change). When splitting, preserve original Law ID with `-A` / `-B` / `-C` suffixes.

### Catch-up rule (added 2026-05-14)
If a scheduled run is missed (weekend, holiday, outage, or any gap >24h since previous run):
- The next run scans the FULL missed window — not just last 24h.
- Catch-up entries get the ACTUAL event date (not the run date) and "CATCH-UP" in the Notes column.
- Applies recursively if multiple runs were missed.

### Survey – AI in Employment tab
Follows the standard 29-column schema. Three tiers per state:
- **Tier A** = enacted state-specific AI-in-employment law
- **Tier B** = active pending bills (overlay baseline + flag in Change Watch)
- **Tier C** = baseline (federal Title VII / ADEA / ADA disparate impact + state EEO)

Daily scan must update this tab whenever an AI-in-employment bill is enacted, signed, vetoed, or materially advances.

### Link verification (added later)
Mandatory every run. HEAD-checks every `Gov Source URL` and `Bill/Statute URL`. Flags:
- 4xx (404 dead, 403 forbidden)
- 5xx (server error)
- Redirect-to-home (regulation-specific URL → bare domain)
- Connection/DNS/SSL errors

For each broken link, search the same agency site to find the current URL, update in place (counts as a CHANGE per BOTH-TABS rule). If no replacement is found after one reasonable search, flag in `Change Watch` with `BROKEN URL since YYYY-MM-DD`.

### Dashboard refresh
Every change goes through the generator — never edit `dashboard.html` directly. The generator auto-picks the freshest of `*.xlsx`, `*_pending_save.xlsx`, `*_NEW.xlsx` (handles main-file-locked-in-Excel case). Dashboard banner at top shows refresh datetime + source filename.

---

## 4. Files in this folder

| File | Purpose |
|---|---|
| `HR_Compliance_Framework_COMPLETE.xlsx` | Canonical source. 18 tabs, ~231 law rows + 14 survey tabs + Reference + Dashboard + QC Findings. 29-column schema on all law/survey tabs. |
| `HR_Compliance_Framework_COMPLETE_pending_save.xlsx` | Sidecar — scheduled scan writes here when main file is locked open in Excel. |
| `HR_Compliance_Framework_COMPLETE_pre-YYYY-MM-DD-backup.xlsx` | Daily rollback snapshots created by the scheduled scan. |
| `generate_dashboard.py` | Builds `dashboard.html` from the xlsx. Self-contained, no external deps beyond `openpyxl`. Auto-picks freshest sidecar. |
| `verify_links.py` | HEAD-checks every URL in the workbook, writes findings to a `Link Check YYYY-MM-DD` tab. Requires `requests`. Sidecar-aware. |
| `dashboard.html` | Generated single-file dashboard, ~530 KB. Embeds D3-projected US outline + all data inline. Works fully offline. |
| `CLAUDE_CODE_HANDOFF.md` | This file. |

### Workbook schema (29 columns, shared across all law and survey tabs)
`Law ID` · `Law / Regulation Name` · `Jurisdiction` · `Core Domain` · `Sub-Domain(s)` · `Category` · `Summary / Key Requirements` · `Applicability` · `Effective Date` · `Status` · `Gov Source URL` · `Bill/Statute URL` · `Littler URL` · `Competitor URL` · `Risk Level` · `Enforcement (1-3)` · `Penalty Exposure` · `Violation Freq (1-3)` · `Clarity (1-3)` · `Multi-State (1-3)` · `Workforce %` · `Primary HR Owner` · `Secondary Owners` · `Compliance Control` · `Audit Frequency` · `Last Review` · `Current Gap` · `Change Watch` · `Notes / Action Items`

### Dashboard features (preserve all of these in the cloud port)
- Top refresh banner (datetime + source filename + file last modified)
- 5 clickable KPI cards with drill-down panels: Total tracked, RED risk, YELLOW risk, Effective ≤ 90 days, Changes last 30 days
- Today's Changes section with executive summary (count breakdown + top 3 highlights with one-line risk framing)
- US state map: real US outline (lat-corrected, aspect ratio 1.71) + transparent pills per state showing reg count, with leader-line callouts for 8 small NE states (VT, NH, MA, RI, CT, NJ, DE, DC) + Maryland
- Category filter dropdown on the map (filters pills + state panel by normalized category)
- Federal button (separate panel for federal regs)
- State drill-down panel: regs grouped by Core Domain, with CoE owner pill (from Primary HR Owner column) next to each
- Upcoming effective dates table (deduped by Law ID, sorted ascending)
- Top jurisdictions bar chart (Chart.js)
- Domain donut chart
- Recent changes feed (last 30 days, deduped)
- Filterable all-regulations table (search + tab/jurisdiction/risk filters)
- Hover tooltips EVERYWHERE: any regulation hover shows a dark floating card with `Summary / Key Requirements`
- Auto dark-mode via `prefers-color-scheme`

---

## 5. Target cloud architecture

```
┌──────────────────────┐
│ GitHub Actions cron  │  ──→ at 03:06 PT weekdays (cron "6 10 * * 1-5" UTC; adjust for DST)
└──────────┬───────────┘
           │ triggers
           ▼
┌──────────────────────┐
│ daily_scan.py        │  ──→ uses Claude Agent SDK + anthropic API
│ (the scheduled task) │       with custom tools: download_xlsx, upload_xlsx,
│                      │       read_xlsx_tab, write_xlsx_row, create_dated_tab,
│                      │       web_search, web_fetch
└──────────┬───────────┘
           │ reads/writes
           ▼
┌──────────────────────┐
│ Cloudflare R2 bucket │  ──→ canonical xlsx + 30-day backup snapshots
│ hr-compliance-data   │
└──────────┬───────────┘
           │ then runs
           ▼
┌──────────────────────┐
│ verify_links.py      │  ──→ HEAD-checks all URLs, writes Link Check tab
└──────────┬───────────┘
           │ then runs
           ▼
┌──────────────────────┐
│ generate_dashboard.py│  ──→ rebuilds dashboard.html from latest xlsx
└──────────┬───────────┘
           │ deploys via Wrangler
           ▼
┌──────────────────────┐
│ Cloudflare Pages     │  ──→ static hosted dashboard
│ + Cloudflare Access  │       gated by email allowlist (free up to 50 users)
└──────────────────────┘
```

### Recommended repo structure
```
hr-compliance-cloud/
├── daily_scan.py              # NEW — Claude Agent SDK wrapper
├── generate_dashboard.py      # COPY as-is from this folder
├── verify_links.py            # COPY as-is from this folder
├── lib/
│   ├── r2_io.py              # download_xlsx / upload_xlsx via boto3
│   ├── xlsx_tools.py         # openpyxl wrappers exposed as Agent SDK tools
│   └── prompts.py            # the daily-scan system prompt (battle-tested)
├── requirements.txt           # anthropic, openpyxl, boto3, requests
├── .github/workflows/
│   └── daily-scan.yml        # cron + run + deploy
├── wrangler.toml              # Cloudflare Pages config
├── README.md
└── CLAUDE_CODE_HANDOFF.md     # this file
```

### Required external accounts/secrets
1. **Anthropic** — console.anthropic.com → API key. Add billing card.
2. **GitHub** — Personal Access Token with `repo` + `workflow` scopes.
3. **Cloudflare** — API token with `Account.R2 Edit`, `Account.Pages Edit`, `Account.Cloudflare Access Edit`.

GitHub Actions secrets:
- `ANTHROPIC_API_KEY`
- `R2_ACCOUNT_ID`, `R2_ACCESS_KEY_ID`, `R2_SECRET_ACCESS_KEY`, `R2_BUCKET`
- `CLOUDFLARE_API_TOKEN`, `CLOUDFLARE_ACCOUNT_ID`

---

## 6. The daily-scan prompt (battle-tested, copy verbatim into `lib/prompts.py`)

```text
Daily HR Compliance regulatory scan. Workflow rules in this prompt are STRICT.

STEP 1 — Download HR_Compliance_Framework_COMPLETE.xlsx from R2 bucket
hr-compliance-data/current.xlsx. Open and load every row from `Federal Laws`,
`State Laws`, and `Additional & Emerging` tabs. Capture Law ID, Law Name,
Jurisdiction, Status, Effective Date, Gov Source URL, Change Watch.

STEP 2 — Catch-up check: compare today's date with the last `Updates YYYY-MM-DD`
tab. If gap > 24h (weekend, holiday, outage), the scan MUST cover the FULL
missed window — not just last 24h. Tag catch-up entries "CATCH-UP" in Notes
and use the ACTUAL event date, not today's date. Apply recursively.

STEP 3 — Scan primary sources:
  Federal: DOL, EEOC, NLRB, OSHA, IRS, USCIS, Federal Register, Congress.gov.
  State: state legislatures + AG offices + labor/civil rights agencies (50 states + DC).
  Tripwires (verify against primary BEFORE using): Littler, Seyfarth, Ogletree,
    Fisher Phillips, Jackson Lewis trackers.
  LOCAL ordinances OUT of scope.

STEP 4 — For every confirmed change, apply BOTH-TABS RULE:
  (a) Update existing law-tab row in place: overwrite changed cell(s),
      set Last Review = today, do NOT append note to Notes / Action Items.
  (b) If NEW reg, add row to appropriate main tab IN THE SAME RUN.
      If multi-provision (e.g., CT PA 26-12), split by sub-domain — each gets
      its own Law ID.
  (c) For previously bundled Additional & Emerging rows where one sub-item
      just became material: split using -A/-B/-C suffix on original Law ID.

STEP 5 — Create a new tab `Updates YYYY-MM-DD` (today's date, user's local TZ).
Schema:
  Type (CHANGE / NEW / SPLIT / CATCH-UP) | Affected Tab | Law ID | Law Name |
  Jurisdiction | Field Changed | Before | After | Primary Source URL |
  Confidence (High/Med/Low) | Notes
EVERY change in STEP 4 must also appear here.

STEP 6 — Review Survey – AI in Employment tab even if no AI changes today.
Re-tier states between Tier A (enacted) / Tier B (pending bills) / Tier C
(baseline) whenever bills are enacted, signed, vetoed, or materially advance.

STEP 7 — Save workbook back to R2 (current.xlsx + dated backup).

STEP 8 — LINK VERIFICATION (mandatory):
  Run verify_links.py. It HEAD-checks every Gov Source URL and Bill/Statute URL
  (~800 URLs). Flags 4xx, 5xx, redirect-to-home, connection errors.
  For each broken link: search the same agency site for the current URL,
  update the row in place (CHANGE per BOTH-TABS rule), log in date-stamped tab
  with Field Changed = Gov Source URL, Before/After. If no replacement found,
  flag in Change Watch as "BROKEN URL since YYYY-MM-DD".

STEP 9 — Run generate_dashboard.py to rebuild dashboard.html. Verify > 100KB
and contains today's date in refresh banner.

STEP 10 — Deploy dashboard.html to Cloudflare Pages via Wrangler.

STEP 11 — Reply with summary:
  - # changes by type (federal vs state, CHANGE/NEW/SPLIT/CATCH-UP)
  - Top 3 highest-impact items with one-line risk framing
  - Link check: X scanned, Y broken, Z auto-fixed, W still broken
  - Live dashboard URL
  If zero changes, say so plainly. Confirm dashboard refreshed.
```

---

## 7. Custom tools needed for the Claude Agent SDK

Define these as Anthropic tool definitions in `lib/xlsx_tools.py` and `lib/r2_io.py`. The Claude Agent SDK will call them autonomously during the daily-scan loop.

| Tool name | Inputs | Returns | Notes |
|---|---|---|---|
| `download_xlsx` | `key` (R2 object key, default `current.xlsx`) | Local path | Caches to /tmp |
| `upload_xlsx` | `local_path`, `key`, `make_backup: bool` | Success | Backup goes to `backups/{date}.xlsx` |
| `read_xlsx_tab` | `tab_name` | List of row dicts | Mirrors generate_dashboard.py's loader |
| `write_xlsx_row` | `tab_name`, `row_index`, `updates: dict` | Success | Used by Step 4(a) |
| `add_xlsx_row` | `tab_name`, `row_dict` | New row index | Used by Step 4(b) |
| `create_dated_tab` | `date_str`, `entries: list` | Success | Step 5 |
| `web_search` | `query`, `num_results: int` | Search results | Anthropic-hosted web search tool |
| `web_fetch` | `url` | Page content | Anthropic-hosted fetch tool |

For the link verification step, `verify_links.py` can run as a standalone Python script — no SDK tool needed. It returns a list of broken URLs to the agent loop, which then does the per-URL search-and-replace using `web_search` + `write_xlsx_row`.

---

## 8. GitHub Actions workflow (drop-in)

```yaml
# .github/workflows/daily-scan.yml
name: Daily HR Compliance Scan
on:
  schedule:
    - cron: "6 10 * * 1-5"   # 03:06 PT (PDT) = 10:06 UTC. For PST, use 11:06 UTC.
  workflow_dispatch: {}        # manual "Run now" button

jobs:
  scan:
    runs-on: ubuntu-latest
    timeout-minutes: 30
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with: { python-version: "3.11" }
      - run: pip install -r requirements.txt
      - name: Run daily scan
        env:
          ANTHROPIC_API_KEY: $
          R2_ACCOUNT_ID: $
          R2_ACCESS_KEY_ID: $
          R2_SECRET_ACCESS_KEY: $
          R2_BUCKET: $
        run: python daily_scan.py
      - name: Regenerate dashboard
        run: python generate_dashboard.py
      - name: Deploy to Cloudflare Pages
        uses: cloudflare/wrangler-action@v3
        with:
          apiToken: $
          accountId: $
          command: pages deploy . --project-name=hr-compliance-dashboard --branch=main
      - name: Notify on failure
        if: failure()
        # email/Slack webhook of choice
        run: echo "scan failed — implement notification"
```

**DST handling:** Pacific time shifts twice a year. Either run the cron twice (once for PST, once for PDT) with appropriate cron guards, or use a single UTC time accepting the 1-hour drift, or use a wrapper script that schedules itself for the next 03:06 PT.

---

## 9. Gotchas learned the hard way (do not re-learn these)

1. **Excel file locking.** If Ruz has the xlsx open in Excel on his Windows machine, scheduled writes fail. Local mitigation: write to `_pending_save.xlsx` sidecar. Cloud version doesn't have this problem since the canonical xlsx lives in R2 — but preserve the sidecar-detection logic in `generate_dashboard.py` for backwards compatibility if anyone manually downloads/edits.

2. **Link rot is ~20%.** A sample check of 17 URLs found 4 broken + 2 redirected. Across 800 URLs, expect 150–200 needing repair. The auto-repair step is non-negotiable.

3. **Survey-tab duplicates.** Many regs are tracked in BOTH a law tab AND a topic-specific survey tab with DIFFERENT Law IDs (e.g., `CA-BRV-001` in State Laws, `BRV-CA-001` in Survey – Bereavement Leave). Same regulation, different IDs. Dedupe by Law ID will NOT catch these. The dashboard's state click-through shows both as separate entries — acceptable.

4. **Bundled emerging-reg rows.** The original workbook had rows like "CA SB 699 + AB 1076" bundled together. Don't split them preemptively — the BOTH-TABS rule says split only when one sub-item becomes material. Splitting too aggressively creates row inflation.

5. **Dashboard generation is the single source of truth.** Never edit `dashboard.html` directly. Always go through `generate_dashboard.py`. The dashboard re-reads the xlsx from scratch every time — this re-validates the data path and surfaces schema drift immediately.

6. **TopoJSON is too big to fetch at runtime.** The US states-10m.json is ~58 KB which exceeded Anthropic's web_fetch cap in one earlier attempt. The current `generate_dashboard.py` embeds a hand-derived simplified US outline (~6.6 KB) inline so the dashboard works fully offline. Don't try to switch back to runtime fetching.

7. **State legislature URLs are the most volatile.** TX `capitol.texas.gov`, CT `cga.ct.gov`, and bill-specific Colorado URLs change parameter formats between sessions. Link verification will find these and the auto-search step should resolve to the new session URL.

8. **EO 14110 → EO 14179.** Anything tied to Biden-era AI executive orders needs to reference the Trump rescission (1/20/2025) and the replacement EO 14179 (1/23/2025). Multiple federal AI guidance items got pulled.

9. **FTC Non-Compete Rule is dead.** Vacated 8/20/2024, FTC abandoned appeal 9/5/2025, formally removed from CFR 2/12/2026. State noncompete bans (CA, MN, ND, OK, MA, TN, UT) are now controlling. Do not re-introduce the FTC rule as enforceable.

10. **DOL OT threshold reverted.** 2024 rule vacated 11/15/2024. EAP exemption is back at $35,568 ($684/week). HCE at $107,432. Do not show the higher numbers as current.

---

## 10. Standing user preferences (carry into the cloud version)

- Always include a `computer://` (or hosted) link to the latest dashboard in every reply.
- Always run a full file refresh (not incremental edits) for the dashboard.
- Daily summary format: count breakdown + top 3 highlights + link-check line + dashboard link.
- No bullet-point padding in summaries. Keep it tight.
- Audit trail lives in the date-stamped `Updates YYYY-MM-DD` tabs — not in row-level notes.

---

## 11. First-run checklist for Claude Code

1. Read this file end to end.
2. Read `generate_dashboard.py` and `verify_links.py` — both work as-is and should be copied verbatim into the cloud repo with one tweak: replace local filesystem path resolution with R2 download/upload at the top and bottom of each.
3. Read the daily-scan prompt above. It is battle-tested. Don't rewrite it without strong reason.
4. Pick a single Anthropic model for `daily_scan.py` — recommend `claude-sonnet-4-6` (good cost/capability balance for this volume). For the final summary generation, the same model is fine.
5. Bootstrap script (`bootstrap.py`) is optional — Ruz declined the auto-provisioning approach. Instead, document the manual setup steps in the README.
6. Test the workflow once with `workflow_dispatch` before relying on the cron. Approve any first-run API permission prompts.
7. Verify the deployed dashboard URL renders identically to the local one. Compare KPI numbers, state pill counts, today's-changes section, and tooltip behavior.
8. Set up Cloudflare Access on the Pages URL before sharing the link with anyone.
9. Set up Slack/email failure notifications on the GitHub Actions workflow.
10. Delete the local Cowork scheduled task once the cloud version has run successfully twice in a row.

---

## 12. Open questions for Ruz to confirm before launch

- **Custom domain?** Default Cloudflare Pages URL is `hr-compliance-dashboard.pages.dev`. If you want `compliance.yourdomain.com`, add a CNAME and update the Cloudflare Access policy.
- **Backup retention?** Recommend 30 daily snapshots in R2. After 30 days, oldest gets deleted by lifecycle policy.
- **Failure notification channel?** Email is simplest. Slack webhook if you want it in a channel.
- **API spend cap?** Recommend setting an Anthropic API spend alert at $100/mo as a tripwire.
- **Who else gets access?** Cloudflare Access free tier allows 50 emails. Add team members in the Zero Trust dashboard.

---

That's everything. The local files in this folder are the source of truth for the port. If something in this doc conflicts with the actual behavior of `generate_dashboard.py` or `verify_links.py`, the code wins — update this doc to match.
