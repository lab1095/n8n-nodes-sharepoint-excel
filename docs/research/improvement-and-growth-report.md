# Improvement & Growth Report — `@lab1095/n8n-nodes-sharepoint-excel`

**Date:** 2026-06-14
**Scope:** Product purpose, technical/UX improvement opportunities, competitive landscape, and a promotion / go-to-market plan.
**Method:** Multi-agent analysis — a code & architecture audit (source read + `bun run test`/`lint`), plus web research across the n8n ecosystem, native nodes, user workarounds, rival automation platforms, direct library alternatives, and the n8n verified-nodes program. Claims are sourced inline.

---

## 1. Executive summary

This node solves a real, well-documented, and still-unfixed problem: n8n's native **Microsoft Excel 365** node reads/writes cell data through Microsoft Graph's `/workbook/` (WAC-token-backed) API, which **fails with `403 "Could not obtain a WAC access token"` on Excel files stored in SharePoint document libraries** (it only works reliably for personal OneDrive). The bug is tracked at [n8n-io/n8n#20040](https://github.com/n8n-io/n8n/issues/20040) — **closed as stale, no code fix** — and recurs across many forum threads. This node sidesteps WAC entirely with a **download → edit (exceljs) → upload** round-trip against the `/content` endpoint. The codebase is clean, tested (153 tests passing, strict TS, lint green), and unusually well-documented for a community node.

**Three findings that should drive the roadmap:**

1. **The market moved.** When this was first published (2026-02-04), nothing solved SharePoint Excel read/write in n8n. As of mid-2026 there is now a **more feature-complete, more-downloaded direct competitor**: `n8n-nodes-msexcel-sharepoint` (~990 downloads/month vs this node's ~195/month). First-mover advantage is eroding; differentiation and distribution now matter.
2. **Verification is the single biggest growth lever — and it's currently blocked.** n8n Verified Community Nodes appear in the in-app node panel and are installable on n8n Cloud. They **forbid runtime dependencies**, but this package ships `exceljs` as a runtime dependency. Bundling exceljs into `dist/` (empty `dependencies`) is the highest-ROI work item in this report.
3. **The biggest correctness gap is multi-item handling.** For most operations the node reads parameters only at item index `0`, so a workflow feeding N items silently processes only the first. This will surprise users and undercuts trust.

---

## 2. Purpose & how it works (confirmed)

- **Problem:** Native Excel 365 node depends on WAC/Office-Online-Server; the `/workbook/` endpoints don't support app-only auth and fail on SharePoint-hosted files (403 WAC, empty sheet/table dropdowns, driveItem-ID mismatch). Sources: [#20040](https://github.com/n8n-io/n8n/issues/20040); [forum: WAC token on corporate SharePoint](https://community.n8n.io/t/critical-n8n-http-request-fails-with-403-forbidden-could-not-obtain-a-wac-access-token-when-accessing-workbook-endpoint-on-corporate-sharepoint/236137); [forum: SharePoint support request](https://community.n8n.io/t/support-for-sharepoint-in-microsoft-excel-node/57538).
- **The native SharePoint & OneDrive nodes** can move whole `.xlsx` files but **cannot touch cell/row data** — confirming the gap. Sources: [SharePoint node docs](https://docs.n8n.io/integrations/builtin/app-nodes/n8n-nodes-base.microsoftsharepoint/), [OneDrive node docs](https://docs.n8n.io/integrations/builtin/app-nodes/n8n-nodes-base.microsoftonedrive/).
- **This node's approach** (`GET /content` → `exceljs` → `PUT /content`) is _exactly_ the workaround the bug reporter and community recommend, packaged into one node. This is the correct core design decision.

---

## 3. Competitive landscape & alternatives

### 3.1 Inside n8n — who actually solves SharePoint Excel _data_ R/W

| Package                                          | What it does                                                                                                                | Solves SP Excel R/W? | Downloads (30d) | Last updated | Verified |
| ------------------------------------------------ | --------------------------------------------------------------------------------------------------------------------------- | -------------------- | --------------- | ------------ | -------- |
| **`@lab1095/n8n-nodes-sharepoint-excel`** (this) | Download-edit-upload via exceljs; sheet/table/workbook ops                                                                  | **Yes**              | ~195            | 2026-06-13   | No       |
| **`n8n-nodes-msexcel-sharepoint`** (K. Malcolm)  | Extends Excel to SharePoint **+ OneDrive**; tables + ranges, add/update/delete/clear rows, search rows, cascading selectors | **Yes (broader)**    | **~990**        | 2026-05-21   | No       |
| `wtyeung/n8n-nodes-ms-onedrive-business`         | OneDrive Business + SP drives; download + append/update worksheet rows                                                      | Partial/Yes          | (n/v)           | 2026-05-20   | No       |
| `@bitovi/n8n-nodes-excel`                        | Add/Delete/List sheets on **binary** Excel via `xlsx`; no SharePoint                                                        | No (pair w/ SP node) | ~342            | 2025-06-30   | No       |
| `n8n-nodes-microsoft-sharepoint` (Savjee)        | Site/File/Folder ops                                                                                                        | No                   | ~117            | 2024-11-08   | No       |
| `n8n-nodes-community-sharepoint` (arisechurch)   | General SharePoint/file access                                                                                              | No                   | ~109            | 2023-10-03   | No       |

Source: npm registry/downloads APIs (`registry.npmjs.org`, `api.npmjs.org`, 30-day point May 15–Jun 13 2026) and GitHub. The n8n verified registry (`api.n8n.io/api/community-nodes`) returns **zero** SharePoint or Excel packages — **none of the above are verified**.

**Verdict:** The author's "I found nothing" was true at build time. It no longer is. `n8n-nodes-msexcel-sharepoint` is now a stronger, more-installed direct competitor (OneDrive support, table writes, delete/clear, ~5× the downloads). This node is still differentiated on **code quality, documentation honesty, and the clean WAC-bypass narrative**, but it is being out-shipped on features and distribution.

### 3.2 Outside n8n — rival platforms hit the _same_ Graph root cause

| Platform       | SharePoint doc libraries                                      | Table required for row ops            | Notes                                                                                                                                        |
| -------------- | ------------------------------------------------------------- | ------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------- |
| n8n native     | Broken (WAC)                                                  | n/a                                   | `/workbook/` session                                                                                                                         |
| Make.com       | Partial (Site/Group drives, 403-prone)                        | No                                    | [apps.make.com/microsoft-excel](https://apps.make.com/microsoft-excel)                                                                       |
| Zapier         | Worst — OneDrive-for-Business; default Documents library only | No (needs headers/row 2)              | [Zapier help](https://help.zapier.com/hc/en-us/articles/8496029352973-How-to-get-started-with-Microsoft-Excel-on-Zapier)                     |
| Power Automate | Best location support                                         | **Yes (mandatory)**                   | 25 MB cap, no arbitrary-cell read except Office Scripts; [connector docs](https://learn.microsoft.com/en-us/connectors/excelonlinebusiness/) |
| Workato        | Yes                                                           | **Yes for rows** ("Get cells" exempt) | [Workato Excel](https://docs.workato.com/connectors/excel.html)                                                                              |

**Strategic takeaway:** No mainstream platform offers **clean, table-free read/write on arbitrary cells in any SharePoint document library** — they all inherit Graph's location limits _or_ mandatory-table constraint. The download-modify-upload approach is the one design that escapes both. **This is the marketing wedge**: "edit any cell in any SharePoint Excel file — no tables required, no WAC errors, no OneDrive-only limitation."

### 3.3 Direct library alternatives (DIY)

The pattern exists as recipes, not packaged tools: **exceljs + Graph** (Node), **Office365-REST-Python-Client + openpyxl** (Python), pandas+requests. Microsoft-native paths (Graph workbook API, Office Scripts + Power Automate, Office-JS add-ins) all require the WAC/session runtime this node avoids. None is a drop-in n8n equivalent.

---

## 4. Technical & product improvement opportunities

Current state verified: `bun run test` → **153 tests pass**; `bun run lint` → **clean**; `tsconfig` strict. Architecture (layered `api`/`descriptions`/`listSearch`/`actions`/`router`) is well above the community-node bar, and resource-locator UX (searchable Site/Drive/File/Sheet/Table + manual entry with regex validation) is excellent.

### High impact

1. **Fix multi-item handling.** `actions/router.ts:13-38` builds context from index `0`, and nearly every action calls `getNodeParameter(..., 0)` (`readRows.ts`, `updateRange.ts`, `clearSheet.ts`, `deleteSheet.ts`, `addSheet.ts`, table ops, `getWorkbooks`). With N input items most ops run **once** and silently ignore items 1..n; `pairedItem` lineage is mostly absent; `continueOnFail` only catches at the top level. → Loop per item `i`, read per-item params, set `pairedItem`, per-item error handling. _(M–L)_
2. **Optimistic concurrency (ETag/If-Match) on writes.** Every write is whole-file `PUT /content` with no `If-Match`; two concurrent runs = silent last-writer-wins data loss (the repo's own WAC research warns against parallel writes). Read the eTag on download, send `If-Match` on upload, surface 412 as a clear conflict. `api.ts:128-160`. _(M)_
3. **429/503 throttling + `Retry-After`.** `graphRequest` (`api.ts:19-115`) only special-cases 423; Graph throttling on `/content` is routine and currently fails hard. Add backoff honoring `Retry-After`. _(M)_
4. **Runtime Excel-Table detection before writes.** The exceljs table-corruption risk (issue #2585) is documented but **not prevented at runtime**. The fix is cheap and already spec'd in `docs/research/wac-tokens-research.md` (check `worksheet.tables`): warn/abort, or route table writes through Graph `/workbook/tables`. _(S–M)_
5. **App-only / client-credentials auth for `/content` ops.** Research confirms `/content` works app-only (unlike `/workbook/`), so unattended automation is achievable — a core use case competitors can't cleanly offer. _(M)_

### Medium impact

6. **Paginate Graph lists/rows** (`@odata.nextLink`) — `getWorkbooks.ts`, `listSearch.ts`, table `getRows`/`lookup`/`getColumns` read `value` once; large drives/sites/tables are silently truncated. _(M)_
7. **Cell value (de)serialization.** exceljs returns `Date`, `{formula,result}`, `{richText}`, `{hyperlink,text}`, `{error}` objects; reads pass them through raw (`readRows.ts:69,107`) and `updateRange` writes everything as text (`updateRange.ts:17-25`). Fixes "my date is a weird object / my number is a string." _(M)_
8. **Add Delete Row(s) and Write-to-Table operations** — closes the most obvious CRUD gaps; table-write avoids corruption. (Competitor `msexcel-sharepoint` already has delete/clear.) _(M each)_
9. **Large-file guard** — `$select=size` pre-check + warning threshold; whole file is loaded in memory via exceljs. _(S)_
10. **Range read/write** beyond a single cell. _(M)_

### Low impact / cleanup

11. De-duplicate `getResourceValue` (in both `listSearch.ts` and `resourceMapping.ts`), the resource-locator unwrap repeated ~12×, the copy-pasted `getRowData` (`appendRows.ts`/`upsertRows.ts`), and the `GRAPH_BASE_URL`/`CREDENTIAL_NAME` constants redeclared in 3 files. _(S)_
12. Remove dead generality (`Source`/`context.source` always `'sharepoint'`; unused `OperationHandler`); fix `download()` return type (`Promise<ArrayBuffer>` but actually `Buffer`). _(S)_
13. Add tests for `api.ts` transport branches, throttling, and multi-item paths; consider `n8n.strict: true` in `package.json`. _(S–M)_
14. Implement **OneDrive support** (docs already mention it; code is SharePoint-only) — also closes a competitor gap. _(M)_

### Security (already reasonable)

`Sites.Read.All` + `Files.ReadWrite.All` is the minimal practical delegated scope (no narrower tenant-wide write scope exists); the documented `Sites.Selected` hardening path is a nice touch. Credential scope is `hidden` so users can't broaden it. Minor: only table _names_ are `encodeURIComponent`'d — low risk since Graph IDs are URL-safe.

---

## 5. Promotion / go-to-market plan

### The #1 lever: get Verified

Verified nodes show in the in-app node panel and install on n8n Cloud — the difference between organic discovery and "you must already know the npm name." Requirements (2026):

- **Zero runtime dependencies** → **bundle exceljs into `dist/`** (tsup/esbuild, exceljs bundled not external; ship empty `dependencies`). _This is the gating task._ Confirmed resolution path in [forum: dependencies rejection](https://community.n8n.io/t/cant-submit-node-invalid-reason-dependencies/205807).
- **No `fs` / `process.env` access** (verify — current in-memory buffer flow should comply).
- One third-party service per package (compliant), must not duplicate a core node (lead with the WAC-bypass distinction), English-only UI/docs, MIT (compliant).
- **GitHub Actions publish with npm provenance — mandatory from May 1, 2026** (`@n8n/node-cli ≥ 0.23.0`, npm Trusted Publishers / OIDC; copy [`n8n-nodes-starter/.github/workflows/publish.yml`](https://github.com/n8n-io/n8n-nodes-starter/blob/master/.github/workflows/publish.yml)).
- Submit via the n8n Creator Portal with npm name, repo URL, descriptions, example workflows, screenshots, and a short demo video. Review is iterative (~1 month). Docs: [submit community nodes](https://docs.n8n.io/integrations/creating-nodes/deploy/submit-community-nodes/), [verification guidelines](https://docs.n8n.io/integrations/creating-nodes/build/reference/verification-guidelines/).

### Do this week (quick wins)

1. Add a **demo GIF** to the README; rewrite the top paragraph around "n8n SharePoint Excel WAC token error."
2. Add npm keywords (`wac-token`, `microsoft-excel-365`, `office365`, `xlsx`, `n8n-node`) and GitHub repo topics.
3. Add README badges: npm version, **npm downloads**, license, CI/tests.
4. Post a genuinely helpful root-cause + workaround comment on [#20040](https://github.com/n8n-io/n8n/issues/20040), with a soft node link.
5. Reply helpfully in [forum #57538](https://community.n8n.io/t/support-for-sharepoint-in-microsoft-excel-node/57538) and related threads.
6. PR the node into `restyler/awesome-n8n` and similar lists.

### Do this month

1. **Bundle exceljs** (unblocks verification) + validate with `@n8n/node-cli`.
2. Set up GitHub Actions publish + provenance.
3. Record the demo video; **submit for verification.**
4. Publish one cornerstone tutorial (dev.to/Medium) titled for the exact pain query; cross-link README.
5. "Show & Tell" post on r/n8n (problem-story framing, disclose authorship).
6. Upload 1–2 workflow templates to n8n.io/workflows (most valuable post-verification).

### SEO target phrases (verbatim, high intent)

`n8n sharepoint excel` · `could not obtain a WAC access token n8n` · `n8n microsoft excel 365 sharepoint not working` · `n8n read/write excel from sharepoint` · `n8n excel sheet dropdown empty sharepoint` · `n8n bypass WAC token excel`.

### Spam-risk guardrails

Never paste identical promo across threads; answer each thread's actual question first and link as a footnote. Don't bump unrelated threads. On Reddit, disclose authorship and lead with the problem story. Don't fabricate download/star counts — let the badge speak.

---

## 6. Suggested priority order

1. **Bundle exceljs** → unblocks verification (biggest distribution win) and is also good hygiene.
2. **Fix multi-item handling** → biggest correctness/trust win.
3. **Submit for verification** (needs #1 + demo video + provenance publish).
4. **Concurrency (ETag) + throttling (429)** → reliability under real load.
5. **Add Delete Row + Table write + value serialization** → feature parity with the new competitor.
6. **Marketing quick wins** (README GIF, badges, #20040 comment, awesome-list PRs) — can run in parallel from day one.

---

## 7. Open items / caveats

- npm **download counts are 30-day points**; weekly/lifetime totals weren't exposed by the APIs used.
- Exact reaction/comment counts on #20040 couldn't be retrieved (GitHub access was scoped to this repo); state read as **closed/stale** — likely but treat as such.
- No r/n8n thread on this topic surfaced in searches (forum + GitHub dominate); the repo's npm scope `@lab1095` and GitHub org `lab1095` match (some search results show a `lab1085` typo — canonical is **lab1095**).
