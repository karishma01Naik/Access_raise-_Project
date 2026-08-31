# GIAM Bulk Access Request Automation

**Status: implemented and verified end-to-end against production** with a
single real self-test submission (2026-08-20) — install, login, validation,
submission, and reporting all confirmed working on a real operator machine.
**Bulk/multi-recipient production use is still pending the on-behalf-of usage
sign-off — see Known Blockers.**

> ⚠️ **Do not run a real *bulk* batch (multiple recipients) against production
> until the bulk/on-behalf-of usage sign-off (design doc Section 9.1) is
> resolved.** Individual self-service test submissions (like the one that
> verified this tool) are lower-risk and already proven to work; submitting
> on behalf of a whole team is the specific thing that still needs sign-off.

Automates bulk GIAM access provisioning for team onboarding: instead of
raising one request per user per access group through the GIAM portal
(`giam.web.allianz`), Cloud Ops provides a list of users + required GIAM
groups and gets back a consolidated status report (success / warning /
error), with failed rows retriable without re-entering everything.

It's a script that orchestrates the existing internal **`lrp` CLI** (a
first-party wrapper around the LRP/GIAM API — see [`lrp-cli`](../lrp-cli)) —
no browser automation, and the decision-making pipeline has **no AI/LLM
involved** (an optional plain-English summary is the one exception — see
[Is this AI?](#is-this-ai)).

**Contents:** [Docs](#docs) · [Known Blockers](#known-blockers-before-a-real-bulk-batch-run) · [Prerequisites](#prerequisites) · [Input Excel Format](#input-excel-format) · [Running Locally](#running-this-locally-per-operator) · [Verifying a Submission](#verifying-a-submission-actually-landed-in-giam) · [AI-Generated Summary](#ai-generated-summary-optional) · [Is This AI?](#is-this-ai) · [Development](#development)

## Docs

- Full design: [`docs/superpowers/specs/2026-08-17-giam-bulk-access-automation-design.md`](docs/superpowers/specs/2026-08-17-giam-bulk-access-automation-design.md)

## Known blockers before a real *bulk* batch run

1. **Bulk/on-behalf-of usage sign-off** — the `lrp` CLI itself is sanctioned
   for on-behalf-of ordering, and a real self-test submission confirmed the
   mechanics work, but bulk whole-team automated submissions to *different*
   recipients aren't yet confirmed as within accepted use. **This is the one
   hard blocker** on running a real multi-recipient batch.
2. No confirmed test/staging LRP instance for safe development.
3. ~~GIAM group catalogue lives on Confluence; no defined sync process to
   the Excel copy this tool validates against.~~ **Resolved (2026-08-20):**
   there's no local catalogue copy anymore — `lrp search-products` itself
   validates each requested group live, so there's nothing to keep in
   sync. Trade-off: any group `lrp` can find is now requestable, not just
   ones your team pre-approved for bulk use — see [Input Excel Format](#input-excel-format).
4. Each operator's own SSO session (`lrp login`) lasts ~1 hour and is
   untested against a full multi-recipient batch's duration — if it expires
   mid-run, the remaining rows come back as `error` and can be retried.
5. `lrp login`'s automated browser can fail with a certificate error even
   when your regular browser works fine (observed directly) — see
   [Prerequisites](#prerequisites) for why, and the automatic workaround
   already built in.

See Section 9 of the design doc for the full list.

## Prerequisites

Every operator runs this on their **own machine**, using their **own**
corporate SSO identity — there is no shared service/TU account, so every
request's audit trail points at a real person. Before running anything:

- **Python 3.9+ and `pip`** — the tool auto-installs its one dependency
  (`openpyxl`) if missing.
- **Network access** to `lrp.giam.allianz` and the internal Artifactory host
  that serves the `lrp` binary. Behind a corporate proxy? `lrp` respects
  `LRP_PROXY_URL`/`HTTP_PROXY`/`HTTPS_PROXY` — see `lrp-cli`'s own README.
- **Your own regular browser, already logged into corporate SSO day-to-day**
  — you'll need it either way (see next point), which is also why this tool
  must run on your own laptop, never a shared/sandboxed environment.
- **Your own corporate SSO credentials + MFA**, completed interactively;
  nothing in this tool handles or stores your credentials.
- **You do *not* need `lrp` installed ahead of time** — it auto-installs via
  the official install script on first run.

**About the login step specifically:** `lrp login` normally opens its own
automated browser (an isolated Chromium instance). In practice, that
automated browser can fail with a certificate error **even when your own
regular browser and `curl` both work fine** — observed directly, not
theoretical. Root cause: that isolated Chromium doesn't share your
OS/regular browser's certificate trust store, so a corporate root CA
trusted everywhere else isn't trusted there.

**You don't need to work around this yourself** — `bulkprovision` already
falls back to `lrp login --manual` automatically when the automated browser
fails: it prints the login URL, you open it in *your own regular browser*,
log in, copy the `Cookie` header from DevTools (Network tab → any request to
`lrp.giam.allianz` → Headers), and paste it back into the terminal. This is
why "your own regular browser, already working for SSO" is a prerequisite
even though `lrp login` starts by trying an automated one.

**Never paste a session cookie anywhere other than that terminal prompt** —
not into a chat, ticket, or document. If one is ever exposed, treat it as
compromised and get a fresh one (see the "force a fresh login" option in
the interactive menu below).

Optional: running this from a **Claude Code** session (e.g. asking it to run
`./azgiam-onboarding.sh` and interpret the report) needs no separate LLM/API key setup —
the pipeline is fully deterministic and behaves identically either way.

**Before running a real bulk batch, confirm the bulk/on-behalf-of usage
sign-off (design doc Section 9.1) has been resolved.**

## Input Excel Format

**Just one `.xlsx` file is required** — the exact column names below (row 1)
are mandatory; column order doesn't matter, and matching is case-insensitive.
One row per requested access item — a recipient needing multiple groups gets
multiple rows.

| Column | Required | Description |
|---|---|---|
| `recipient` | Yes | Whatever `lrp find-user` accepts to identify the person: email, name, or user ID. Must resolve to exactly one real user. |
| `giam_group` | Yes | The exact GIAM group/role name to request (e.g. `AWS-ReadOnly-Subscription-662403251356`). |
| `justification` | Yes | Business reason for the request — shown to GIAM approvers. Each row is submitted as its own independent `lrp order-product` call (since 2026-08-20 — see [Is this AI?](#is-this-ai)), so a recipient with multiple rows can give each group its own justification. |

Example (`samples/input.xlsx` — kept short, just one row, since it's only
illustrating the format; add as many rows as you need for a real batch):

| recipient | giam_group | justification |
|---|---|---|
| jane.doe@allianz.com | APP-EXAMPLE-GROUP | Onboarding test |

**There is no separate catalogue file.** Instead, `lrp search-products` is
called live for every `giam_group` in your input — if it can't find an
exact match, that row becomes `error` (with the reason, e.g. "not found")
and nothing is submitted for it, but it never blocks that recipient's
*other* requested groups. This means the group just has to be real and
findable in LRP, not on a separately-maintained pre-approved list — worth
knowing if your team wants to restrict bulk requests to a curated subset of
groups rather than "anything `lrp` can find."

A row with a blank `recipient` or `giam_group` isn't silently dropped — it's
reported as `rejected` with a reason, so every input row is accounted for in
the final report.

## Running this locally (per operator)

### Interactive menu (no raw `lrp` commands needed)

Run with no arguments for a menu covering everything this tool (and `lrp`
itself) can do — you never need to type a raw `lrp` command directly:

```sh
./azgiam-onboarding.sh
```

```
GIAM Bulk Provisioning - Interactive Menu
1) Run a bulk batch (input Excel)
2) Check authentication status
3) Log in (opens browser, falls back to --manual if needed)
4) Find a user
5) Search for a GIAM group/product
6) List your recent orders
7) Exit
```

Option 3 ("Log in") always forces a genuinely fresh authentication — it
never silently trusts a cookie that's already cached, even if that cookie
is technically still valid. Use it whenever you want to be certain you're
not reusing an old session (e.g. after a cookie was ever exposed anywhere
it shouldn't have been).

### Direct batch run (no menu)

The wrapper script auto-installs dependencies, auto-generates a run ID, and
handles `lrp login` itself (with the automatic `--manual` fallback described
in [Prerequisites](#prerequisites)) if you're not already authenticated:

```sh
./azgiam-onboarding.sh <input.xlsx>
```

1. Prepare your input Excel file (see
   [Input Excel Format](#input-excel-format); `samples/input.xlsx` has a
   placeholder example). For your own real per-operator data (real
   recipients/justifications), put it under `local/` instead of editing
   `samples/` — `local/` is gitignored and never committed.
2. Run it:
   ```sh
   ./azgiam-onboarding.sh local/input.xlsx
   ```
   Complete SSO if prompted (automated browser, or the `--manual` fallback
   if that fails) — the run continues automatically once authenticated.
   Each row prints its own status line as soon as it's done, e.g.:
   ```
   [1/3] jane.doe@allianz.com | APP-EXAMPLE-GROUP | submitted | Order submitted successfully.
   [2/3] john.smith@allianz.com | APP-EXAMPLE-GROUP | warning | This request was already submitted/assigned previously for this user. (lrp: ...)
   [3/3] row 4 (missing recipient) | row 4 (missing recipient) | rejected | Input row 4 could not be processed: missing recipient
   ```
   so you're never watching a silent terminal during a multi-row batch —
   the same final summary line (and AI summary, if enabled) still prints
   once everything's done.
3. Check the printed `report.csv` for per-row status (`submitted` /
   `warning` / `rejected` / `error`), and `retry.json` for just the rows
   that errored. A `warning` row's message always leads with a plain
   statement like *"This request was already submitted/assigned previously
   for this user"* — that recipient already had (or already requested) that
   group before, so nothing new was submitted for that row; it's not a
   failure. If the local `claude` CLI is installed, a plain-English summary
   of the same results also prints automatically (see
   [AI-Generated Summary](#ai-generated-summary-optional) below) — pass
   `--no-ai-summary` to skip it.
4. To retry only the rows that errored, read `retry.json` and re-run with
   just those rows in a new input Excel — the CLI doesn't yet read
   `retry.json` directly as input (a possible v1.1 improvement; see design
   doc Open Questions).

Prefer to call the underlying CLI directly (e.g. to control the report/retry
file paths yourself)? `azgiam-onboarding.sh` is a thin wrapper around:

```sh
python3 -m bulkprovision.cli \
  --input local/input.xlsx \
  --report report.csv \
  --retry-out retry.json \
  --run-id "$(date +%Y%m%d-%H%M%S)"
```

**Exit codes:** `0` clean run · `1` some rows errored · `2` not authenticated
(both login attempts failed or were declined — run `lrp login` manually and
retry) · `3` `lrp` couldn't be installed automatically · `4` failed to
read/parse the input Excel · `5` unexpected error during
processing/reporting (a bug — the run may be partially submitted; check
`lrp list-orders` and any partial report before re-running) · `6` `--input`
was given without all of `--report`/`--retry-out`/`--run-id`
(omit `--input` entirely for the interactive menu instead).

## Verifying a submission actually landed in GIAM

`report.csv` reflects `lrp`'s immediate API response to the submission call
— for full confirmation that GIAM registered (and, eventually, approved) the
request, check independently:

1. **`lrp list-orders`** (or `-o json` for full detail) — find the submitted
   group/product name and check its `status`: `Pending` (awaiting
   approval), `Approved`, or `Rejected`.
2. **The `message` column in `report.csv`** — carries `lrp`'s own outcome
   text (e.g. "Order submitted successfully.") per row.
3. **The GIAM portal itself** (`giam.web.allianz`), under your own request
   history — the human-facing source of truth, independent of the CLI.

## AI-Generated Summary (optional)

After every run, `bulkprovision` attempts a plain-English summary of the
*already-computed* results — via the operator's own local **Claude Code**
CLI (`claude -p ...`), under their existing profile, no separate API key.
This is a pure presentation layer: the AI only phrases results that
`orchestrator.py` already classified deterministically — it never
re-judges, re-classifies, or decides anything. If `claude` isn't installed,
times out, or fails for any reason, it's skipped and the deterministic
report/exit code are completely unaffected — but **every skip reason now
prints a one-line `AI summary skipped: ...` message to stderr** (added
2026-08-21, after a real run had `claude` installed and on `PATH` yet the
summary never appeared with zero diagnostic output anywhere). If you don't
see the summary, check your terminal for that line — it tells you exactly
which of "not on PATH" / timed out / non-zero exit (with the actual stderr
tail) / empty output happened. Pass `--no-ai-summary` (or add it to your
own `azgiam-onboarding.sh` invocation) to skip it outright. Example output:

```
--- AI-generated summary (via local Claude Code) ---
**3 requests: 1 submitted, 1 warning, 0 rejected, 1 error.**

| Recipient | Group | Status | Reason |
|---|---|---|---|
| jane.doe@allianz.com | APP-GITHUB-ADMIN | submitted | Order submitted successfully. |
| jane.doe@allianz.com | APP-ALREADY-ASSIGNED | warning | Already has this access — not a failure. |
| jane.doe@allianz.com | APP-RETIRED-ROLE | error | lrp could not find this group; nothing submitted. |

Error rows are in the retry file and can be re-run once corrected. Warning
rows are never in the retry file.
```

No boilerplate "what these statuses mean" glossary repeats every run — just
a headline, a table covering every row (including the ones that succeeded),
and a closing note only when there's something to retry.

## Is this AI?

**Mostly no.** Every decision-making step — parsing Excel, validating
each group live via `lrp search-products`, resolving users, classifying
results, submitting each row's order — is deterministic Python calling the
real `lrp` CLI. No model is
involved in any of that, even though the originating request framed this
as an "AI Agent" (see design doc Section 4 for the reasoning). The **one**
optional exception is the summary above: a model rephrases already-decided
results in plain English, purely for readability — it has no say in what
actually happened. You can also optionally ask a **Claude Code** session to
run the tool's commands for you, which is just an assistant typing on your
behalf, not AI inside the pipeline.

## Development

Run the test suite (120 tests, no network/`lrp`/`claude` binary required) from the
repository root:

```sh
python3 -m pytest -v
```
