# GLSenseFinanceTemplate (VBA) vs GLSense (C#) — Implementation Understanding

This document captures a first-pass read of three repositories, built to prepare for a full implementation walkthrough:

- **GLSenseFinanceTemplate** — a single macro-enabled workbook, `finance_report_macro_template.xlsm`, containing ~20 VBA modules/forms/classes.
- **GLSense** — the VSTO Excel add-in (C#/WPF) previously worked on in this session; the template's VBA is, per the developer's own description, "majorily derived from the drilldowns and jobs monitor sections" of this project.
- **XLEdge** — a sibling VSTO add-in (also C#/WPF, same architectural patterns: `WpfAppManager`, `DpiAwareWindow`, `AppOverlay`), currently mid-migration from an older VB.NET codebase (see its own `MIGRATION_STATUS.md`). It integrates with GLSense (GLLogin cascades a login into XLEdge via COM) but is otherwise a separate reporting product. Referenced here only where it's directly relevant (e.g. `AttachmentLinkHelper.cs`).

The overall conclusion, reached independently by every angle of analysis below: **the C# GLSense add-in is the mature, original implementation. The VBA workbook is a later, deliberately trimmed-down, standalone distillation of it** — built so a workbook can keep supporting login, the 8 drilldown types, and a jobs monitor without the add-in installed. Evidence for this direction is collected in its own section at the end.

---

## 1. Repository/module map

### VBA modules (extracted from `xl/vbaProject.bin` via `oletools`)

| Module | Lines | Role |
|---|---|---|
| `ThisWorkbook.cls` | 389 | Workbook lifecycle; hyperlink-click gate (`Workbook_SheetFollowHyperlink`); custom-drilldown and journal-attachment handlers |
| `PublicSubs.bas` | 1216 | CommandBar (toolbar) construction; login/logout; drilldown dispatch; Excel state toggling |
| `PublicFunctions.bas` | 717 | General-purpose helpers (HTTP client, JSON validation, formula-argument parsing, ledger lookups) — despite the name, **not** worksheet UDFs |
| `GLParser.bas` | 127 | Extracts `GLSENSE_GETBALANCE(...)` call text out of a formula string (paren/quote-aware scanner) |
| `DrilldataToWorksheet.bas` | 808 | Core drilldown-result writer: metadata → array → worksheet → formatting → finalize |
| `FrmMonitor.frm` | 684 | Jobs monitor UserForm (ListView of tracked drilldown jobs) |
| `FrmProcessing.frm` | 8 | Blank marquee/progress UserForm (no code) |
| `ProgressBar.bas` | 38 | Show/Hide/Close wrapper around `FrmProcessing` |
| `FormLogin.frm` | 87 | Login UserForm — embeds an IE `WebBrowser` control, scrapes the auth cookie after server-side login |
| `FormHyperlink.frm` | 102 | Journal-attachment picker dialog |
| `JSONBuilder.bas` | 419 | Domain-specific REST payload builders (balance/journal/subledger drilldown requests) |
| `JsonConverter.bas` | 2301 | Third-party "VBA-JSON" library (Tim Hall, VBA-tools) — generic parse/serialize engine, not application code |
| `modFileLogger.bas` | 223 | Hand-rolled rotating file logger |
| `modCustomXmlHelpers.bas` | 65 | Reads/cleans `<DRILLDOWNSHEET>` Custom XML Parts (drilldown-sheet provenance) |
| `mod_GlobalDataManager.bas` | 270 | In-memory ledger/segment metadata cache, built from a hidden metadata worksheet |
| `clsLedgerStructure.cls` | 85 | Ledger metadata value object |
| `clsArrayList.cls` | 88 | Hand-rolled ArrayList (VBA has no generics) |
| `Declarations.bas` | 15 | Global app-state variables (**not** Win32 API declarations, despite the name) |
| `Sheet4.cls` / `Sheet5.cls` | 8 each | Empty worksheet code-behind stubs |

### Corresponding GLSense C# areas

| Concern | Files |
|---|---|
| Drilldown entry points | `Drilldowns/DD_BL.cs`, `DD_JL.cs`, `DD_SL.cs`, `DD_ExcelPrecedents.cs`, `DrillCellHighlighter.cs` |
| Drilldown worksheet writer | `Drilldowns/DDDatatoWorksheet.cs` (1829 lines) |
| Bulk/background refresh (VBA has no equivalent) | `Drilldowns/BulkRefreshProcess.cs`, `BalanceRefresh.cs`, `BalanceNormalizer.cs`, `ExcelFormulaGenerator.cs`, `DataTableBuilder.cs` |
| Formula parsing/repair (VBA has no equivalent) | `Helpers/ClsFormulaParser.cs` |
| Jobs monitor | `Views/GLJobsMonitor.xaml(.cs)`, `ViewModels/GLSubmittedJobsViewModel.cs`, `Models/GLJobModel.cs` |
| Single-operation wait dialog | `Views/GLWaitWindow.xaml(.cs)` |
| Worksheet UDFs | `GLSenseExcelFunctions.cs` |
| Ribbon entry points | `AddinModule.cs` |
| Login | `Views/GLLogin.xaml(.cs)` |
| Logging | `Utilities/LogUtility.cs`, `Helpers/LogHelper.cs` |
| Global session state | `AppState.cs` |
| Ledger/segment metadata (persisted) | `Repositories/LedgerDataRepository.cs`, local SQLite DB |
| User preferences (VBA has no equivalent) | `Utilities/UserConfig.cs`, SQLite-backed |
| JSON response parsing | `Helpers/JsonHelper.cs`, `Helpers/JsonGlobals.cs` |
| Attachment dialog | `Views/AttachmentsDialog.xaml.cs` |

---

## 2. Drilldown pipeline

### Trigger → fetch (VBA)

A floating CommandBar ("ORBIT" menu), built in `PublicSubs.Auto_Open`, wires 8 buttons to thin one-line Subs (`Balance_Drilldown`, `Balance_Journals_Drilldown`, `Balance_SubLedger_Drilldown`, `Journal_Drilldown`, `Balance_Drilldown_SubLedger_Drilldown`, `Balance_Drilldown_Unified_Drilldown`, `Subledger_Drilldown`, `Unified_Drilldown`), each calling the private dispatcher `ExecuteDrilldown(DDType)` with a short type code (`BL`, `JL`, `SL`, `BL_JL`, `BLDD_SL`, `BLDD_UF`, `UF`). `ExecuteDrilldown`:
1. Gates on `CanProceed` (metadata sheet present, logged in, CubeID set).
2. Builds a request via `BuildBalanceRequest`/`BuildJournalRequest`/`BuildSubledgerRequest` (which in turn call into `JSONBuilder.bas`'s `BalanceJson`/`JournalJson`/`SubLedgerJson` for the actual payload construction).
3. POSTs synchronously via `gethttp` (`MSXML2.ServerXMLHTTP60`, `Open(..., Async:=False)` — **this blocks the entire Excel UI thread** for the duration of the call) to `rest/secure/finance/{balance-drilldown|balance-journal-drilldown|balance-subledger-drilldown|...}`.
4. On a synchronous success response, hands the parsed dictionary straight to `DrilldataToWorksheet.DrilldownDataToSheet`. On a "processing in background" response, registers the job's process ID into the `XLDrillJobs` named range and the user later retrieves it via the jobs monitor.

### Trigger → fetch (C#)

`DD_BL.cs`'s `DrilldownBl.ProcessBLDrilldown()` (and parallel `DD_JL.cs`/`DD_SL.cs` classes) do the equivalent of `ExecuteDrilldown`, but fully asynchronously: resolve the clicked range (`ExcelExternalRef.ResolveRangeWithContext`) → `GuardValidInputs` (formula-presence check) → `CommonMethods.DisableExcelSettings()` → show `GLWaitWindow` progress dialog → optional formula auto-repair (`ValidateAndFixFormulasAsync`, no VBA equivalent) → build payload (`CreateBalancePayload`, `System.Text.Json`) → `BuildDrilldownInfo` (produces the same `"WorksheetName*Address*Timestamp*DDType*Title"` encoded string VBA uses) → `BuildApiUrl` → `FetchDrilldownDataAsync` (async `ApiHelper.ServerAPI`) → `ProcessDrilldownResponseAsync` (parses JSON; on background-job response registers an Excel named range `GLSense_DD_{processId}`; on success constructs `DDDatatoWorksheet` and awaits `DD_DatetoWorksheet()`).

**Same REST endpoints, same encoded string, same Bearer-token auth, same background-job/named-range registration pattern on both sides** — this is a shared backend contract, not two independently-designed APIs.

### Worksheet write — near line-for-line port

`DrilldataToWorksheet.DrilldownDataToSheet` (VBA) and `DDDatatoWorksheet.DD_DatetoWorksheet` (C#) run the **identical seven-step pipeline, in the same order**, with matching private-method names:

| Step | VBA | C# | What it does |
|---|---|---|---|
| 1 | `ExtractMetadata` | `ExtractMetadata`/`BuildMetadataDictionary`/`FillColumnAndTypeInfo` | Build displayName-keyed metadata dict, DataType/Format dicts, backfill missing record keys |
| 2 | `PrepareDataArray` + `JavaDateConv` | `PrepareDataArray`/`FillDataRows` + `JavaDateConv` | Header + data rows into a 2D array; epoch-ms → date; neutralize leading `=` (formula-injection guard) |
| 3 | `PrepareWorksheet` | `PrepareWorksheet`/`FindWorksheetForDrilldown` | Parse `ObjStr` on `"*"`, reuse existing sheet/table name, clear cells |
| 4 | `ApplyDataFormats`/`ApplyingFormatsToRange` | `ApplyDataFormats`/`ApplyColumnFormat` | Identical default number formats per TEXT/DECIMAL/INTEGER/DATE |
| 5 | `PopulateSheet` | `PopulateSheet` | Paste array from row 5; internal column names in row 4, white font (hidden) |
| 6 | `ApplyFormatting` | `ApplyFormatting`/`ApplySubledgerLinkStyling`/`ApplyAttachmentHyperlinks`/`ApplyColumnSubTotals` | `ListObject` table, subledger blue-underline styling, ATTACHMENT hyperlinks, `customFormula`, `customDrilldownConfig` hyperlinks, `SUBTOTAL()` via same 101–111 code table, hide internal columns |
| 7 | `FinalizeSheet` | `FinalizeSheet` | Header row coloring, A1 title by DDType, C1 timestamp, C2 timezone, A2/A3 reference text, tab color by DDType |

Both also persist per-column drilldown config into the workbook's `CustomXMLParts` via a matching `CreateCustomDrilldownXMLPart`/`EscapeXml` routine, using the same `<DRILLDOWNSHEET><COLUMNNAME Name="...">` XML shape.

**⚠️ One naming discrepancy worth resolving in the walkthrough:** VBA's hyperlink-click gate (`ThisWorkbook.Workbook_SheetFollowHyperlink`) checks the sheet's `ListObject` name **contains** `"ORB_XLDD_"`; the C# gate (`AddinModule.cs`) checks it **starts with** `"ORB_DD_"`. Different prefix, different match semantics — either an intentional "XL" (Excel-template) naming convention to distinguish template-generated tables from add-in-generated ones, or genuine drift. Worth confirming which is intended.

### What C# adds that VBA has no equivalent of

- **Bulk refresh subsystem** (`BulkRefreshProcess.cs`, `BalanceRefresh.cs`, `BalanceNormalizer.cs`, `ExcelFormulaGenerator.cs`, `DataTableBuilder.cs`): extracts every `GLSENSE_GETBALANCE` UDF call in a workbook, batch-resolves values, and rewrites formulas (baking in literal values for cells with no cross-references). VBA's `GLParser.ExtractGLSenseCalls` only *extracts* call text — it never rewrites formulas.
- **Cell highlighting / precedent tracing** (`DrillCellHighlighter.cs`, `DD_ExcelPrecedents.cs`) — full precedent-tree walking, no VBA equivalent found.
- **Formula auto-repair** (`ClsFormulaParser.cs`'s `Formula_Correction`, used via `ValidateAndFixFormulasAsync`) — VBA's closest cousin, `PublicFunctions.TrapParameters`, only *parses* arguments, it doesn't fix/pad malformed ones.
- Excel-state snapshot/restore (ScreenUpdating/DisplayAlerts/EnableEvents/Calculation) and an explicit 1,048,576-row guard before writing — both absent from the VBA writer.
- Async/await + `CancellationToken` throughout, plus a live progress dialog with dozens of granular `SetProcessMessage(...)` step updates — VBA's `ExecuteDrilldown` is one blocking call with only a bare marquee (`ShowProgress`/`CloseProgress`) shown around it.

### `GLParser.bas` vs `ClsFormulaParser.cs` vs `ExcelFormulaGenerator.cs` — not the same thing

These three are easy to conflate but solve different sub-problems:
- **`GLParser.ExtractGLSenseCalls`** — narrowest: find balanced `glsense_getbalance(...)` call text in a formula string. Nothing else.
- **`ExcelFormulaGenerator.ExtractAllUdfCalls`** — uses the *same* paren-depth/quote-state scanning algorithm as `GLParser` (closest cousin), but goes further: pairs each call with a returned balance value, substitutes it in, and decides static-value-vs-live-formula. This is the "smart bulk refresh" logic — no VBA equivalent.
- **`ClsFormulaParser.cs`** — broadest: general argument extraction/resolution (`ExtractArguments_WithValues`, cell-reference resolution, `INDIRECT()` unwrapping, formula padding/repair). VBA's rough equivalent is `PublicFunctions.TrapParameters`/`CleanParameter`, but without the repair/resolution capability.

---

## 3. Jobs Monitor

### VBA `FrmMonitor.frm`

A `ListView`-based UserForm with **no typed model at all** — everything lives in two private fields: `JobsStr` (raw comma-joined process IDs from the workbook-level named range `XLDrillJobs`) and `MonitorReports` (`Scripting.Dictionary`, ProcessID → a `~!~`-delimited string re-split with `Split()` on every render). Columns: checkbox, Process ID, Sheet Name, Phase, Status, Date, synthetic "Download" text.

Data flow: `InitFormLoad` reads `XLDrillJobs`, makes a **synchronous, blocking** `gethttp` GET to `.../rest/secure/finance/drilldown-processes?limit=100&page=1` (freezes Excel for up to the 30s XHR timeout), parses with `JsonConverter.ParseJson`, filters to jobs whose description contains `*` and whose process ID is tracked, rebuilds the dictionary and ListView from scratch. Search/filter (`Populate_List`) is pure client-side `Like`-pattern matching over the cached dictionary (3 modes: Contains/Starts with/Ends with).

Actions: `CmdRefresh_Click` (full re-fetch), `CmdDelete_Click`/`CmdDeleteAll_Click` (silent — rewrites/deletes the `XLDrillJobs` named range, no server call, no confirmation), `CmdDrilldowns_Click` (downloads selected Complete+Success jobs via `DrillData` → second `gethttp` call → `DrilldownDataToSheet`), `CmdLogs_Click` (separate raw XHR to a log-download endpoint, saves a `.txt` to `%LOCALAPPDATA%\ORBIT\Excel_Logs`), `CmdCancel_Click` (just closes the window — not a job cancel). No threading exists; `DoEvents` around the modeless `ShowProgress`/`CloseProgress` marquee (`ProgressBar.bas` + blank `FrmProcessing.frm`) is the only concession to UI responsiveness. There's dead/unwired instrumentation code (`ProcessDrilldown`, `LogPerformanceAfterProcessing`) referencing `JsonConverter.ParseJsonWithTiming` performance experiments.

### C# `GLJobsMonitor` / `GLSubmittedJobsViewModel`

Full MVVM: `GLJobModel` (typed, `INotifyPropertyChanged`, with computed UI properties like `DownloadIconKind`/`StatusColor`/`PhaseColor`) backing an `ObservableCollection` bound to a `DataGrid` via `ICollectionView`. `LoadJobsAsync` hits the **identical** `drilldown-processes?limit=100&page=1` endpoint, but asynchronously (`ApiHelper.ServerAPI` with a 300s `CancellationTokenSource`), deserializes into typed `DrilldownJobsResponse`/`JobRecord` via `System.Text.Json`, and marshals the parse work onto a background thread before hopping back to the UI dispatcher (`ParseAndDisplayJobs`). Excel-side job tracking uses **one named range per job** (`GLSense_DD_{processId}`) rather than VBA's single CSV named range. Search offers 7 match modes vs VBA's 3. Delete flows show a confirmation dialog first (VBA's are silent). Download-outputs branches by job type, including a `"FinanceSnapshotJob"` type VBA's monitor doesn't appear to handle at all. Logs download via `HttpClient` streaming a server-named zip (vs. VBA's raw `.txt` write). The single-operation wait dialog (`GLWaitWindow`) is the structural analog of `FrmProcessing`/`ProgressBar.bas`, but adds a real, wired-up **Cancel button** (`CancellationHelper.Cancel()`) — VBA's `FrmProcessing` has no cancel affordance at all, only the fixed 30s XHR timeout.

**Neither side persists jobs in a database** — both use Excel named ranges as the durability mechanism and re-fetch full status from the same REST endpoint on demand. The C# side didn't add real persistence here, just async and richer typing.

---

## 4. UDF / formula layer

The real worksheet UDF producing the formula text both codebases operate on is C#'s `GLSenseExcelFunctions.cs` (`GLSense_GetBalance`, an AddinExpress XLL-module function), which emits `GLSENSE_GETBALANCE(...)` cell formulas — the exact literal name `GLParser.bas` hardcodes and searches for. The VBA argument order in `JSONBuilder.BalanceJson`/`PublicFunctions.TrapParameters` matches `GLSense_GetBalance`'s parameter list position-for-position (ChangeSign, LedgerName, Activity, Period, BalanceType, CurrencyCode, TranslatedFlag, ActualFlag, BudorEncName, JESource, JECategory, then 9 segments) — an unusually precise match for something built independently.

VBA's `PublicFunctions.bas`, despite its name, contains **no actual worksheet UDFs** — it's a general helper library (HTTP client `gethttp`, JSON validator `ReturnJsonDict`, formula-argument parser `TrapParameters`, ledger lookups `GetLedgerArrays`, timezone lookup, XML escaping, registry read for the Downloads folder). C# has a much larger family of real UDFs beyond `GLSense_GetBalance`: period lookups (`GLSense_GetPeriod*` — 8 variants), segment lookups (`GLSense_GetSegmentDesc`, etc.), account type/DFF lookups, and daily rate conversion — none of which have a VBA UDF counterpart (VBA's `ConvertPeriod` is a crude, private echo of the `GLSense_GetPeriod*` family).

`Declarations.bas` is **not** a Win32 API layer despite its name — it's just global app-state variables (auth token, login URL, cube ID, journal dictionary, debug flag). The real Win32 `Declare` statements in the VBA project live inside the third-party `JsonConverter.bas` (timezone conversion boilerplate, not application code). The honest C# counterpart of `Declarations.bas` is `AppState.cs`'s properties (same concepts, consolidated into one observable, loggable singleton instead of scattered globals) — not `DpiAwarenessHelper.cs` (which is a WPF-hosting-specific concern with no VBA equivalent, since Excel already handles DPI scaling for native UserForms).

---

## 5. Login, workbook lifecycle, logging, global state

### Login

Both use the same fundamental strategy — **embed a browser, let the server's own login page authenticate, then scrape the resulting session cookie** — but with a generational tooling gap: VBA's `FormLogin.frm` uses a legacy IE `WebBrowser` ActiveX control and manually string-splits `document.cookie`; C#'s `GLLogin.xaml` uses WebView2 with `CoreWebView2.CookieManager.GetCookiesAsync` (a real API), adds a server picker, SSO support (`AllowSingleSignOnUsingOSPrimaryAccount`), reads additional cookies (`X-ORB-USERNAME`, `XLEDGE-USER-ACCESS`), and cascades a successful login into the sibling XLEdge add-in via COM reflection — none of which VBA needs or has, being self-contained by design. Both hit the same `finance-cubes` endpoint after login; VBA does it synchronously via `gethttp`, C# asynchronously via `ApiHelper.ServerAPI`.

### Workbook lifecycle

There's no `Workbook_Open` handler in the VBA template at all — login happens on-demand, and `LoginURL`/`CubeID` can be pre-seeded directly in metadata-sheet cells for an "auto-login" shortcut (consistent with distributing an already-configured exported workbook). The real lifecycle logic is in `Workbook_SheetFollowHyperlink`, which gates on the active sheet being a recognized drilldown table and branches by hyperlink `ScreenTip` text (`"CUSTOM DRILLDOWN"` → `GLSenseCustomDrillDown`, `"ATTACHMENTS LINK"` → `HandleJournalAttachments`). `GLSenseCustomDrillDown` explicitly comments `' Remove 'id' if it exists (per .NET logic)` — a direct textual reference to the C# implementation as the behavior's source of truth. `Workbook_BeforeClose` just tears down the custom CommandBar menu.

### Logging

Both loggers share the same root log folder (`%LOCALAPPDATA%\ORBIT\Excel_Logs\...`, differing only in subfolder name — `GLFinanceLogs` for VBA vs `GLSense_Logs` for C#) and near-identical header formatting conventions — a strong signal one was modeled directly on the other. VBA's `modFileLogger.bas` is a competent hand-rolled rotating file logger (4 levels, 10MB rotation, 30-day cleanup) with no dedup/scoping. C#'s `LogUtility.cs` (backed by NLog) adds exception deduplication (5s window), scoped/nested logging (`LogScope`, BEGIN/END indentation), pre-init debug buffering, and delegates rotation/retention to NLog itself.

### Global state

VBA's `mod_GlobalDataManager.bas` is a narrow, session-rebuilt ledger/segment metadata cache (scanned from a hidden worksheet each time) — not general app state. Its closer C# analog for *persistence* is actually `Repositories/LedgerDataRepository.cs` (SQLite-backed, cross-session), a real architectural upgrade over VBA's in-memory-only dictionaries. VBA's loose `Declarations.bas` globals map most directly onto `AppState.cs`'s properties (session/login/selection state), while `GlobalStateViewModel.cs` is unrelated (just a reference-text string) and `mod_GlobalDataManager.bas`'s "3 public APIs" (`GetLedgerInfo`, `GetSegmentNamesAndIDs`, `GetSegmentValuesBySetID`) have no single 1:1 file — they're spread across `LedgerDataRepository`/`DataRepository`/`ServiceLocator.SegmentDataService` in C#.

### JSON

`JSONBuilder.bas` is a thin, domain-specific **request-payload builder** on top of the third-party `JsonConverter.bas` (3 entry points: `SubLedgerJson`, `JournalJson`, `BalanceJson`, the last including a small DSL for segment-token parsing: `~` force-non-summary, `--` NOT, `|` BETWEEN, `%` LIKE, comma-separated IN). C#'s `JsonHelper.cs`/`JsonGlobals.cs` solve the *opposite* problem — defensive, case-insensitive **response reading** (`TryGetProperty`/`TryGetDouble`/`TryGetString` etc. over `System.Text.Json`), with no payload-builder equivalent needed because C# just serializes typed POCOs directly. The domain-specific *request-building* logic in C# lives inside the `Drilldowns` namespace itself, not in a JSON helper file.

### Attachments

`FormHyperlink.frm` (VBA attachment picker) and `Views/AttachmentsDialog.xaml.cs` (C#) are a near line-for-line port of each other, including a matching `KeyData()` reverse-lookup function — C# just adds logging around identical core logic. Note: this is a different mechanism from `XLEdge/Helpers/AttachmentLinkHelper.cs` (parses Oracle EBS/Fusion HTML anchor snippets into ScreenTip strings) — that belongs to the sibling XLEdge product's own attachment-link flavor, not this journal-attachment flow.

---

## 6. Direction of derivation — consolidated evidence

Every one of the four independent analysis passes above reached the same conclusion from different angles:

1. **Feature subtraction is the only pattern found.** Every divergence is "C# has X, VBA lacks it" (bulk refresh, precedent tracing, cell highlighting, formula auto-repair, async/cancellation, granular progress messaging, row-limit guard, user preferences, SQLite persistence, confirmation dialogs, richer search, snapshot-job support, SSO). Never the reverse.
2. **VBA explicitly references the C# implementation as ground truth** — `ThisWorkbook.cls`'s `' Remove 'id' if it exists (per .NET logic)` comment.
3. **VBA knows exact, oddly-specific C# implementation details it would have no reason to invent independently** — the literal UDF name `glsense_getbalance` and its precise parameter ordering (`GLParser.bas`, `JSONBuilder.BalanceJson`) match `GLSenseExcelFunctions.GLSense_GetBalance`/`ExcelFormulaGenerator.cs`'s `FunctionName = "GLSENSE_GETBALANCE"` constant exactly.
4. **Both sides share a non-obvious, specific convention** — the `<DRILLDOWNSHEET><COLUMNNAME Name="...">` Custom XML Parts scheme for tracking drilldown-sheet provenance appears in both `modCustomXmlHelpers.bas`/`ThisWorkbook.cls`/`DrilldataToWorksheet.bas` and `AddinModule.cs`/`DDDatatoWorksheet.cs`, down to the same tag names.
5. **Logger branding is explicit** — VBA's log header literally says *"GLSense Finance Template Logs"*, self-identifying as a GLSense derivative.
6. **Structural decomposition matches almost function-for-function** between `DrilldataToWorksheet.bas` and `DDDatatoWorksheet.cs` (same private-method names/order) — unusual for an independent from-scratch VBA build, consistent with translating an already-decomposed C# method structure back into VBA one sub at a time.
7. **Both hit the identical REST paths and cookie names**, confirming one already-existing server API that both are clients of — with the C# side's far greater sophistication marking it as the reference implementation the API was designed alongside.
8. **VBA's login form auto-seeds from cells that look pre-populated by an export step** (`LoginToOrbit` checks a token already sitting in `OrbitSchMetadata!Z2`), consistent with "take a workbook a user already built with the full add-in, strip the add-in dependency, keep just enough VBA alive to preserve drilldown/login/monitor behavior."

One caveat surfaced by the Jobs Monitor analysis: comments *inside* `GLSubmittedJobsViewModel.cs` itself (`// Helper methods from VB.NET`, `// Implement from your VB.NET GLSense_DownloadProcessesc method`) suggest an even earlier VB.NET/VBA-era implementation existed before *either* of the two artifacts examined here — i.e. the likely full lineage is an original VB.NET/VBA prototype → matured into the current C# GLSense add-in → later distilled back down into this standalone VBA workbook template. Worth confirming directly with you in the walkthrough.

---

## 7. Open questions for the walkthrough

- Confirm the `ORB_XLDD_` (VBA, contains-match) vs `ORB_DD_` (C#, starts-with-match) table-prefix discrepancy — intentional or drift?
- Is there an even earlier VB.NET codebase (per the `GLSubmittedJobsViewModel.cs` comments) worth pulling in as a fourth reference point?
- Which of the "VBA lacks this" gaps (bulk refresh, precedent tracing, formula auto-repair, snapshot jobs, user preferences) are things you actually want ported into the template, versus deliberately out of scope for a lightweight standalone workbook?
- Confirm whether the template is meant to keep tracking a *separate* backend contract long-term, or whether it should stay strictly in lockstep with whatever `GLSenseExcelFunctions.cs`/`ExcelFormulaGenerator.cs` emit as the C# side evolves.
- Priority order for the implementation walkthrough itself — suggest starting with the drilldown pipeline (section 2) since it's the largest, most directly-ported piece, then Jobs Monitor, then login/lifecycle.
