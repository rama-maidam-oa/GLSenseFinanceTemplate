# GLSenseFinanceTemplate (VBA) — Fixes ported from GLSense (C#)

**Status (2026-10-05): every fix in this document, #1–#26, is applied** in
`finance_report_macro_template.xlsm`. #1–#7 were applied by hand before 2026-09-02.
#8–#26 were applied by Excel automation, then the template was re-signed with
`SignTemplate.bat` and pushed as commit **`0040acf`** (OISR-22571 / OISR-22575). See
[*Applied 2026-10-05*](#applied-2026-10-05--commit-0040acf-oisr-22571--oisr-22575) at the end for
how the fixes were applied and validated, and for #22–#26.

Sources: drilldown-parity audit against `GLSense\FinalWorkingCode\GLSense` (2026-08-17), a
delta audit of later C# fixes (#8–#13), and a VBA-vs-C# HTTP request audit (#14–#21), both
on 2026-10-05.

Each entry gives the **module**, what to **find**, and what to **replace it with**, as it was
applied. The find/replace text is kept so the changes can be reviewed, or re-applied to an output
file generated from an older template.

---

## 1. `JSONBuilder.bas` — `isFunctionalCurrency` missing from Balance payload

**Root cause:** VBA's `BalanceJson` never sent an `isFunctionalCurrency` field at all. C#'s
`BalanceDtoModel.CreateFromXllParameters` computes it by checking every ledger named in the
formula (a formula can name multiple, comma-separated) — true if **any** named ledger's
functional currency matches the balance's currency code, false only if none do.

### 1a. Add new function

Insert immediately after `GetLedgerInformation` ends (i.e. right before
`Private Function NewBuildSegJSON(...)`):

```vba
''' <summary>
''' Determines whether ANY of the ledger(s) named in this formula's own LedgerName
''' parameter has a functional (base) currency equal to the balance's currency code.
''' Mirrors GLSense (C#) BalanceDtoModel.CreateFromXllParameters: a formula can name
''' multiple ledgers (comma-separated), and each can have a different functional
''' currency, so every named ledger must be checked individually - True if ANY of
''' them matches, False only if at least one resolves but none match. Defaults to
''' True only when no named ledger can be resolved at all, matching the C# fallback.
''' </summary>
Private Function IsFunctionalCurrency(ByVal LedgStr As String, ByVal balanceCurrencyCode As String) As Boolean
    Dim LedgSplit   As Variant
    Dim i           As Long
    Dim currentName As String
    Dim oLedger     As clsLedgerStructure
    Dim anyResolved As Boolean

    anyResolved = False
    IsFunctionalCurrency = True ' Default to True if no matching ledger found (matches C#)

    LedgSplit = Split(LedgStr, ",")

    For i = LBound(LedgSplit) To UBound(LedgSplit)
        currentName = Trim(LedgSplit(i))
        If currentName <> "" Then
            Set oLedger = GetLedgerInfo(currentName)
            If Not oLedger Is Nothing Then
                If Not anyResolved Then
                    anyResolved = True
                    IsFunctionalCurrency = False ' At least one ledger resolved; stop defaulting to True
                End If
                If oLedger.CurrencyCode = balanceCurrencyCode Then
                    IsFunctionalCurrency = True
                    Exit Function ' Short-circuit, same as C#'s .Any()
                End If
            End If
        End If
    Next i
End Function
```

### 1b. Wire it into `BalanceJson`

**Find:**
```vba
                    D1("balanceType") = FuncParam(4)
                    D1("currencyCode") = FuncParam(5)
                    D1("translatedFlag") = FuncParam(6)
```

**Replace with:**
```vba
                    D1("balanceType") = FuncParam(4)
                    D1("currencyCode") = FuncParam(5)
                    D1("isFunctionalCurrency") = IsFunctionalCurrency(Trim(FuncParam(1)), FuncParam(5))
                    D1("translatedFlag") = FuncParam(6)
```

---

## 2. `JSONBuilder.bas` — `encumbranceTypeId` missing from Journal payload

**Root cause:** `JournalJson` reads indexes 0,1,4–13 out of the `~~`-delimited cell value but
skips index 3 entirely. C#'s `DD_JL.cs` always sends `encumbranceTypeId` (that same index 3)
for `JL`, `BLDD_SL`, and `BLDD_UF` drilldowns.

**Function:** `JournalJson`

**Find:**
```vba
            strValue = Split(ActiveSheet.Cells(CR.Row, ColFound).value, "~~")
            D1("actualFlag") = strValue(1)
            D1("balanceType") = strValue(6)
            D1("codeCombinationId") = strValue(4)
            D1("ledgerId") = strValue(5)
            D1("periodName") = strValue(0)
```

**Replace with:**
```vba
            strValue = Split(ActiveSheet.Cells(CR.Row, ColFound).value, "~~")
            D1("actualFlag") = strValue(1)
            D1("balanceType") = strValue(6)
            D1("codeCombinationId") = strValue(4)
            D1("ledgerId") = strValue(5)
            D1("periodName") = strValue(0)
            If Len(Trim(strValue(3))) > 0 And IsNumeric(strValue(3)) Then
                D1("encumbranceTypeId") = CLng(strValue(3))
            Else
                D1("encumbranceTypeId") = 0
            End If
```

---

## 3. `PublicFunctions.bas` — CTD/JED/JEDP/JEDU period corrupted

**Root cause:** CTD/JED/JEDP/JEDU period arguments are built as an Excel concatenation literal
(`"Jan-08"&"~"&"Feb-08"`). C#'s formula parser (`ClsFormulaParser.cs`) has a special case that
strips the `&` operators before quote-cleanup for any argument containing both `&` and `~`.
VBA's `CleanParameter` has no equivalent — it only strips one outer quote pair — so it leaves
stray embedded quotes in the result, corrupting `periodName` for any Balance Drilldown on these
4 balance types.

**Function:** `CleanParameter`

**Find:**
```vba
    ' Trim spaces
    param = Trim(param)
    
    ' Remove surrounding quotes
    If Len(param) >= 2 Then
```

**Replace with:**
```vba
    ' Trim spaces
    param = Trim(param)
    
    ' CTD/JED/JEDP/JEDU period args are built as Period&"~"&EndPeriod (see
    ' GLConfiguratorViewModel.CombinePeriod in the add-in) - strip the "&"
    ' operators and every quote character directly, rather than the generic
    ' single-outer-quote-strip below, which leaves stray embedded quotes.
    If InStr(param, "&") > 0 And InStr(param, "~") > 0 Then
        CleanParameter = Trim(Replace(Replace(param, "&", ""), """", ""))
        Exit Function
    End If
    
    ' Remove surrounding quotes
    If Len(param) >= 2 Then
```

---

## 4. `DrilldataToWorksheet.bas` — table-name reuse on refresh never fires

**Root cause:** Drilldown tables are generated as `"ORB_XLDD_" & Format(Now, "ddMMyyyyhhmm")`
(a 9-character prefix). The "reuse existing table name on refresh" check compares only the
first **7** characters, so the comparison can never be true — every refresh of an
already-populated drilldown sheet creates a brand-new timestamped table instead of keeping the
old one, silently breaking any named range/formula/bookmark pointing at the old table name.

**Function:** `PrepareWorksheet`

**Find:**
```vba
            If Left(ws.ListObjects(1).Name, 7) = "ORB_XLDD_" Then tblName = ws.ListObjects(1).Name
```

**Replace with:**
```vba
            If Left(ws.ListObjects(1).Name, 9) = "ORB_XLDD_" Then tblName = ws.ListObjects(1).Name
```

---

## 5. Unguarded `EnableActions` restore — crash/hang risk (3 call sites)

**Root cause:** Same bug class already fixed in C#'s `CommonMethods.cs`/`AddinModule.cs`
(see this project's `CLAUDE.md`) — a restore-Excel-state call sitting bare inside an
error-recovery label, with no protection if it itself throws (e.g. Excel mid-close, COM in a
bad state):
- 2 sites use `Resume`, which **re-arms** the handler — a second failure there loops the same
  GoTo/Resume forever (Excel hangs, no dialog).
- 1 site uses a plain `GoTo`, which leaves the handler **disarmed** — a second failure there
  escapes as a raw uncaught VBA runtime error, leaving `ScreenUpdating`/`DisplayAlerts`/
  `EnableEvents` stuck off.

### 5a. Add a non-throwing wrapper

**Module:** `PublicSubs.bas` — add right after the existing `EnableActions` Sub:

```vba
Public Sub SafeEnableActions()
    ' Non-throwing counterpart to EnableActions, for use ONLY at CleanUp/error-recovery
    ' labels that already run under an active "On Error GoTo" handler for the same
    ' procedure. A second failure restoring Excel state there (e.g. Excel already
    ' closing) must not itself throw - depending on whether that handler was reached
    ' via Resume (handler re-armed - a repeat failure here would GoTo/Resume in an
    ' infinite loop) or a plain GoTo (handler left disarmed - a repeat failure here
    ' would escape uncaught as a raw runtime-error dialog, leaving ScreenUpdating/
    ' DisplayAlerts/EnableEvents stuck off). Mirrors GLSense (C#)
    ' CommonMethods.TryEnableExcelSettings.
    On Error Resume Next
    Call EnableActions
    On Error GoTo 0
End Sub
```

### 5b. Swap the 3 call sites

**Module:** `PublicSubs.bas`, Sub `ExecuteDrilldown`

**Find:**
```vba
CleanUp:
    Call EnableActions
    Call CloseProgress
    If Len(errorString) > 0 Then MsgBox errorString, vbCritical, "ORBIT"
```

**Replace with:**
```vba
CleanUp:
    Call SafeEnableActions
    Call CloseProgress
    If Len(errorString) > 0 Then MsgBox errorString, vbCritical, "ORBIT"
```

---

**Module:** `ThisWorkbook.cls`, Sub `Workbook_SheetFollowHyperlink`

**Find:**
```vba
CleanExit:
    Call EnableActions
    DoEvents
    
    Call CloseProgress
```

**Replace with:**
```vba
CleanExit:
    Call SafeEnableActions
    DoEvents
    
    Call CloseProgress
```

---

**Module:** `FrmMonitor.frm`, Sub `CmdDrilldowns_Click`

**Find:**
```vba
CleanUp:
    Call CloseProgress
    DoEvents
    Call EnableActions
    If Len(errorString) > 0 Then MsgBox errorString, vbCritical, "ORBIT"
```

**Replace with:**
```vba
CleanUp:
    Call CloseProgress
    DoEvents
    Call SafeEnableActions
    If Len(errorString) > 0 Then MsgBox errorString, vbCritical, "ORBIT"
```

---

## 6. `JSONBuilder.bas` — `ledgerIdList` sent as one comma-joined string instead of an array of IDs

**Root cause:** Found by comparing a live payload capture from the two sides (`ExcelTemplate.json`
vs `FromMSI.json`, `%LOCALAPPDATA%\ORBIT\Excel_Logs\GLFinanceLogs\`, 2026-08-21) for the same `BL`
drilldown, which reproduced the `HTTP 400` on `balance-drilldown` seen in that day's log.

`GetLedgerInformation` already builds a correct comma-joined string of ledger IDs when a formula
names multiple ledgers (e.g. `"300000046988965,300000046975971"`) — same pattern `IsFunctionalCurrency`
uses. But `BalanceJson` never splits that string back apart before adding it to the JSON array; it
adds the whole string as a **single** collection item, so `ledgerIdList` serializes as a one-element
array containing a quoted comma-joined string (`["300000046988965,300000046975971"]`) instead of
C#'s array of numeric IDs (`[300000046988965, 300000046975971]`, one per ledger, unquoted). Every
other field in the two payloads matched exactly.

**Function:** `BalanceJson`

**Find:**
```vba
                    GetLedg = GetLedgerInformation(Trim(FuncParam(1)), Coaid)
                    
                    If Len(GetLedg) > 0 And GetLedg <> "" Then
                        C2.Add GetLedg
                        D1.Add "ledgerIdList", C2
                        D1("coaid") = Coaid
                    Else
```

**Replace with:**
```vba
                    GetLedg = GetLedgerInformation(Trim(FuncParam(1)), Coaid)
                    
                    If Len(GetLedg) > 0 And GetLedg <> "" Then
                        Dim LedgIdSplit As Variant
                        Dim ledgIdx     As Long
                        LedgIdSplit = Split(GetLedg, ",")
                        For ledgIdx = LBound(LedgIdSplit) To UBound(LedgIdSplit)
                            C2.Add CDec(Trim(LedgIdSplit(ledgIdx)))
                        Next ledgIdx
                        D1.Add "ledgerIdList", C2
                        D1("coaid") = Coaid
                    Else
```

`CDec` (not `CLng`) because ledger IDs like `300000046988965` exceed `Long`'s ~2.1 billion range.
`ConvertToJson` (same module, ~line 3525) emits `vbDecimal` values unquoted with no scientific
notation, matching the C# output exactly.

---

## 7. `JSONBuilder.bas` — `JournalJson` throws "Subscript out of range" when the helper cell has fewer than 14 `~~`-delimited tokens

**Root cause:** `JournalJson` reads `strValue(0)` through `strValue(13)` unconditionally out of the
`~~`-delimited helper cell (`Split(ActiveSheet.Cells(CR.Row, ColFound).value, "~~")`). For
period-based balance types (YTD/PTD/QTD — no date range), the helper string only carries 12
tokens (indices 0-11); `startDate`/`endDate` (indices 12-13) are never appended upstream. Indexing
past `UBound` throws VBA runtime error 9, caught by `JournalJson`'s own `On Error GoTo GlobalErr`,
which logs the error and returns `""` — surfacing to the user as "Selected range does not have
values to perform journals drilldowns." / "Error building journal request" from
`BuildJournalRequest`. Confirmed against a live download (`PeriodParam_BL_Q32!Q6:S11` in a
user-provided workbook) where every helper cell split to exactly 12 parts. Same root-cause class
as fix #2 above (parser assumes a fixed token count; helper-string shape actually varies by
balance/drilldown type).

Verified against the C# reference (`GLSense\FinalWorkingCode\GLSense\Drilldowns\DD_JL.cs`), which
builds the equivalent payload from the same `~~`-delimited string
(`BuildDrilldownList`, `DD_JL.cs:371-415`) and defines the exact wire behavior to match:

- **String fields** (`periodName`, `actualFlag`, `balanceType`, `currencyCode`, `translatedFlag`,
  `jeSourceName`, `jeCategoryName`, `status`, `startDate`, `endDate`): C# uses
  `parts.ElementAtOrDefault(n)`, which returns `null` for a missing index. `JsonGlobals.Options`
  sets `DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull` (`JsonHelper.cs:27`), so a
  missing token **omits the key from the JSON entirely** — it is not sent as `""` or `null`. A
  token that exists but is genuinely empty (e.g. `jeSourceName`/`jeCategoryName` in the sample
  data) still serializes as `"jeSourceName": ""`, since empty string isn't null.
- **Numeric fields** (`codeCombinationId`, `ledgerId`, `encumbranceTypeId`): C# always emits these,
  unquoted, via `ToLongSafe(parts.ElementAtOrDefault(n))`, which returns `0` when the token is
  missing or non-numeric (`long` is a non-nullable value type, so `WhenWritingNull` never drops
  it). VBA's existing code assigns the raw split **string** for `codeCombinationId`/`ledgerId`
  (never converted to a number), and VBA-JSON only auto-unquotes numeric-looking strings that are
  16–100 characters long (`json_StringIsLargeNumber`) — so typical short CCIDs/ledger IDs like
  `"12831"` or `"1"` were being sent **quoted**, unlike C#'s unquoted `12831`. `encumbranceTypeId`
  was previously omitted entirely (not defaulted to 0) when its token was blank/non-numeric,
  also unlike C#.

**Function:** `JournalJson`

### 7a. Add bounds-safe accessors

Insert immediately before `Public Function JournalJson(...)`:

```vba
''' <summary>
''' Bounds-safe *presence* check for a ~~-delimited helper-cell Split() result,
''' mirroring C#'s ElementAtOrDefault + DefaultIgnoreCondition.WhenWritingNull
''' (DD_JL.cs BuildDrilldownList + JsonHelper.cs JsonGlobals): a missing index
''' means the JSON key must be omitted entirely, not sent as "" or null. Returns
''' the element via outVal and True only when idx actually exists in arr.
''' </summary>
Private Function TryElem(ByVal arr As Variant, ByVal idx As Long, ByRef outVal As String) As Boolean
    If idx >= LBound(arr) And idx <= UBound(arr) Then
        outVal = arr(idx)
        TryElem = True
    Else
        TryElem = False
    End If
End Function

''' <summary>
''' Bounds-safe numeric element access, mirroring C#'s ToLongSafe(parts.ElementAtOrDefault(n))
''' (DD_JL.cs): always returns a value (0 when the token is missing/blank/non-numeric) since
''' these fields are always present, unquoted, in the C# payload - never omitted like the
''' string fields. CDec (not CLng) matches the existing ledgerIdList fix (#6) so large IDs
''' beyond Long's ~2.1 billion range still serialize correctly via ConvertToJson.
''' </summary>
Private Function SafeNumElem(ByVal arr As Variant, ByVal idx As Long) As Variant
    Dim s As String
    s = ""
    If idx >= LBound(arr) And idx <= UBound(arr) Then s = Trim(arr(idx))
    If Len(s) > 0 And IsNumeric(s) Then
        SafeNumElem = CDec(s)
    Else
        SafeNumElem = CDec(0)
    End If
End Function
```

### 7b. Use them for every indexed access

**Find:**
```vba
            strValue = Split(ActiveSheet.Cells(CR.Row, ColFound).value, "~~")
            D1("actualFlag") = strValue(1)
            D1("balanceType") = strValue(6)
            D1("codeCombinationId") = strValue(4)
            D1("ledgerId") = strValue(5)
            D1("periodName") = strValue(0)
            If Len(Trim(strValue(3))) > 0 And IsNumeric(strValue(3)) Then
                D1("encumbranceTypeId") = CLng(strValue(3))
            End If
            D1("currencyCode") = strValue(7)
            D1("translatedFlag") = strValue(8)
            D1("jeSourceName") = strValue(9)
            D1("jeCategoryName") = strValue(10)
            D1("status") = strValue(11)
            D1("startDate") = strValue(12)
            D1("endDate") = strValue(13)
```

**Replace with:**
```vba
            strValue = Split(ActiveSheet.Cells(CR.Row, ColFound).value, "~~")
            Dim elemVal As String

            D1("codeCombinationId") = SafeNumElem(strValue, 4)
            D1("ledgerId") = SafeNumElem(strValue, 5)
            D1("encumbranceTypeId") = SafeNumElem(strValue, 3)

            If TryElem(strValue, 1, elemVal) Then D1("actualFlag") = elemVal
            If TryElem(strValue, 6, elemVal) Then D1("balanceType") = elemVal
            If TryElem(strValue, 0, elemVal) Then D1("periodName") = elemVal
            If TryElem(strValue, 7, elemVal) Then D1("currencyCode") = elemVal
            If TryElem(strValue, 8, elemVal) Then D1("translatedFlag") = elemVal
            If TryElem(strValue, 9, elemVal) Then D1("jeSourceName") = elemVal
            If TryElem(strValue, 10, elemVal) Then D1("jeCategoryName") = elemVal
            If TryElem(strValue, 11, elemVal) Then D1("status") = elemVal
            If TryElem(strValue, 12, elemVal) Then D1("startDate") = elemVal
            If TryElem(strValue, 13, elemVal) Then D1("endDate") = elemVal
```

**Why the numeric fields moved above the `TryElem` block:** `encumbranceTypeId`, `codeCombinationId`,
and `ledgerId` are unconditional now (always present, defaulting to 0), matching C#'s non-nullable
`long` properties — grouping them together makes that "always sent" behavior visually distinct
from the `TryElem`-gated string fields that can be omitted.

---

## Delta audit — C# fixes made after 2026-08-17 (audited 2026-10-05)

Entries #1–#7 above were confirmed **already applied** in the current
`finance_report_macro_template.xlsm`. Entries #8–#13 below port the C# drilldown fixes that
landed in `GLSense\FinalWorkingCode` *after* the first audit. The new code for #8, #11 and #12 was
compiled and run in a scratch Excel workbook (#8: 15 token cases; #11: normal / empty /
missing / null `metadata`), not by editing the template.

> **Note on `FSG_Test (8).xlsm`:** that sample carries an *older* VBA build with **none** of
> #1–#7 applied. Outputs generated before the template fixes keep the old VBA, so they
> won't show any of these fixes until they're regenerated from the current template.

---

## 8. `JSONBuilder.bas` — `~` / `--` segment markers stripped anywhere, corrupting real segment codes

**C# fix:** `8398ec3` (2026-09-23), `BalanceDtoModel.CreateSingleSegmentValue`.

**Root cause:** `ParseSegmentToken` uses `Replace(token, "~", "")` and `Replace(token, "--", "")`,
which strip the markers **anywhere** in the value. The markers are positional: `~` is only ever
a *trailing* "summary checkbox modified" marker (`SegmentSelectorViewModel.GetEffectiveValue`),
and `--` is only ever a *leading* exclude prefix. Real segment codes can contain `~`.

**Real-data repro (`FSG_Test (8).xlsm`, `OrbitSchMetadata!X2`):** the formula's segment value is
`"~118"`, a real code in value set `1002472` ("Test Special Character", summary flag `N`).
This is the root cause of **OISR-22571** ("Balance Drilldown does not exist although the balances
are available").

| | `values` | `summaryEnabled` |
|---|---|---|
| VBA (before fix) | `["118"]`, a code that doesn't exist | `false` |
| VBA (after fix) | `["~118"]` | `false` (summary flag `N`) |
| C# (11.1.0–11.1.2 after `7b460a4` / `62ce35c` / `a4d4847`) | `["~118"]` | `false` |

Confirmed by running the workbook's own `BalanceJson` on cell F3 of the real sample, before and
after (see *Applied 2026-10-05*).

### 8a. Add helper

Insert immediately **above** `Private Function ParseSegmentToken(...)`:

```vba
Private Function StripTrailingTilde(ByVal s As String) As String
    s = Trim(s)
    Do While Right(s, 1) = "~"
        s = Left(s, Len(s) - 1)
    Loop
    StripTrailingTilde = Trim(s)
End Function
```

### 8b. Replace the whole `ParseSegmentToken` function

Replace everything from `Private Function ParseSegmentToken(ByVal token As String, _` through
its `End Function` with:

```vba
Private Function ParseSegmentToken(ByVal token As String, _
    ByVal segSetID As String) As Scripting.Dictionary
    Dim D           As Scripting.Dictionary
    Dim values      As Collection
    Dim parts       As Variant
    Dim SegOverride As Boolean
    Dim isNot       As Boolean
    Dim bound2      As String

    Set D = New Scripting.Dictionary
    Set values = New Collection

    token = Trim(token)

    ' Markers are POSITIONAL, matching GLSense (C#) BalanceDtoModel.CreateSingleSegmentValue:
    '   "~"  = trailing only (the "summary checkbox was modified" marker)
    '   "--" = leading only  (the exclude / NOT marker)
    ' Real segment codes can legitimately start with or contain "~" (e.g. "~118", "~GTS"),
    ' so a blanket Replace() of either marker corrupts them into a different code.

    If InStr(token, "%") > 0 Then
        ' LIKE pattern - only a trailing "~" is stripped
        D("operator") = "LIKE"
        D("summaryEnabled") = False
        values.Add StripTrailingTilde(token)

    ElseIf InStr(token, "|") > 0 Then
        ' BETWEEN / NOTBETWEEN - each bound carries its own markers ("X~|Y~").
        ' The Segment Values window writes NOT BETWEEN as "--X|--Y" (both bounds prefixed),
        ' so the leading "--" of the whole token decides the operator and is also stripped
        ' from the second bound.
        isNot = (Left(token, 2) = "--")
        If isNot Then token = Mid(token, 3)
        parts = Split(token, "|")

        If isNot Then
            D("operator") = "NOTBETWEEN"
        Else
            D("operator") = "BETWEEN"
        End If
        D("summaryEnabled") = False

        values.Add StripTrailingTilde(parts(0))
        If UBound(parts) >= 1 Then
            bound2 = StripTrailingTilde(parts(1))
            If isNot And Left(bound2, 2) = "--" Then bound2 = Trim(Mid(bound2, 3))
            values.Add bound2
        Else
            values.Add ""
        End If

    Else
        ' IN / NOTIN - "--X~" : strip the trailing marker first, then the leading prefix
        SegOverride = (Right(token, 1) = "~")
        token = StripTrailingTilde(token)

        isNot = (Left(token, 2) = "--")
        If isNot Then token = Trim(Mid(token, 3))

        If isNot Then
            D("operator") = "NOTIN"
        Else
            D("operator") = "IN"
        End If

        If SegOverride Then
            D("summaryEnabled") = False
        Else
            D("summaryEnabled") = IsSummary(segSetID, token)
        End If

        values.Add token
    End If

    D.Add "values", values
    Set ParseSegmentToken = D
End Function
```

**Behavior kept on purpose:** VBA's existing rule "trailing `~` ⇒ `summaryEnabled = False`" is
preserved. It's correct: a trailing `~` means a summary account converted to non-summary
(user-confirmed). C# has now been fixed to match. See *C#-side issues* below.

**Verified results** (scratch workbook, stub `IsSummary` where `~118`/`1000` are summary accounts):

```text
~118            => IN         summary=True  values=[~118]
~118~           => IN         summary=False values=[~118]
1000~           => IN         summary=False values=[1000]
--1000~         => NOTIN      summary=False values=[1000]
--~118          => NOTIN      summary=True  values=[~118]
A-B--C          => IN         summary=False values=[A-B--C]
1000~|2000~     => BETWEEN    summary=False values=[1000][2000]
--1000|--2000   => NOTBETWEEN summary=False values=[1000][2000]
~118|~200       => BETWEEN    summary=False values=[~118][~200]
~1%~            => LIKE       summary=False values=[~1%]
```

---

## 9. `PublicSubs.bas` — Journal drilldown rejects headers written with spaces ("PTD NET")

**C# fix:** `cc7b08f` (2026-09-17), `DD_JL.GetDrilldownType`.

**Root cause:** `BuildJournalRequest` matches the row-5 header with exact `LCase` equality, so
`PTD NET` / `ACCOUNTED CR` hit "Invalid selection!". C# now normalizes spaces to underscores.

**Find** (in `BuildJournalRequest`):
```vba
    Else
        val = LCase(val)
    End If
```

**Replace with:**
```vba
    Else
        ' Normalize spaces to underscores so "PTD NET" matches the same as "PTD_NET" (C# DD_JL)
        val = LCase(Replace(Trim(val), " ", "_"))
    End If
```

---

## 10. `PublicSubs.bas` — Background-job path leaves Calculation=Manual and **EnableEvents=False**

**C# fix (same bug class):** `e0a7a46` (2026-09-17), `DD_SL.ProcessSLDrilldown`. An early exit
skipped restoring Excel settings, so Excel looked hung.

**Root cause:** In `ExecuteDrilldown`, when the server queues the drilldown as a background job
(`MessageString <> ""`), the block ends with a bare `Exit Sub`. That skips `CleanUp:` and
therefore `SafeEnableActions`. `DisableActions` had set `Calculation = xlCalculationManual`,
`DisplayAlerts = False` and `EnableEvents = False`. Unlike `ScreenUpdating`, those **don't**
reset when the macro ends. With events off, `Workbook_SheetFollowHyperlink` never fires again,
so **drilldown-sheet hyperlinks (custom drilldowns / attachments) stop working**. Formulas also
stop recalculating. Both last until Excel is restarted.

**Find** (in `ExecuteDrilldown`):
```vba
        If FrmMonitor.Visible = True Then
            Call FrmMonitor.InitFormLoad
        Else
            FrmMonitor.Show vbModeless
        End If
        Exit Sub
    End If
```

**Replace with:**
```vba
        If FrmMonitor.Visible = True Then
            Call FrmMonitor.InitFormLoad
        Else
            FrmMonitor.Show vbModeless
        End If
        Call SafeEnableActions   ' Exit Sub here bypasses CleanUp - restore Excel state explicitly
        Exit Sub
    End If
```

---

## 11. `DrilldataToWorksheet.bas` — Response without a `metadata` node silently writes nothing

**C# fix:** `6739550` (2026-08-11), `DDDatatoWorksheet.ExtractMetadata`.

**Root cause:** VBA already handled `"metadata":[]`. But when `metadata` is **missing** (or JSON
`null`), `ExtractMetadata` exits *before* deriving columns from the records' keys. Then
`DrilldownDataToSheet` sees `LastColumn = 0` and does a bare `Exit Sub`: no sheet, no message,
and no log line saying why. C# always derives columns from record keys and warns if none are found.

### 11a. Replace the whole `ExtractMetadata` function

```vba
Private Function ExtractMetadata(DrillsData As Object, DataTypeDict As Scripting.Dictionary, _
                                FormatDict As Scripting.Dictionary, ActualColumnName As Collection, _
                                DisplayColumnName As Collection) As Scripting.Dictionary
    On Error GoTo ErrHandler
    Dim metaDict As New Scripting.Dictionary
    metaDict.CompareMode = TextCompare

    Call modFileLogger.FileLogDebug("Extracting metadata from JSON object", "ExtractMetadata")

    ' A missing/null/empty "metadata" node is a valid response shape - columns must then
    ' be derived from the records' own keys below, so never bail out of the function here.
    ' Mirrors GLSense (C#) DDDatatoWorksheet.ExtractMetadata (fix 6739550).
    Dim metaItem As Object
    If Not DrillsData.exists("metadata") Then
        Call modFileLogger.FileLogWarn("No metadata node found in DrillsData - deriving columns from record keys", "ExtractMetadata")
    ElseIf TypeName(DrillsData("metadata")) <> "Collection" Then
        Call modFileLogger.FileLogWarn("Metadata node is not a list (" & TypeName(DrillsData("metadata")) & ") - deriving columns from record keys", "ExtractMetadata")
    Else
        For Each metaItem In DrillsData("metadata")
            If TypeName(metaItem) = "Dictionary" Then
                If metaItem.exists("displayName") Then
                    Set metaDict(metaItem("displayName")) = metaItem
                    DisplayColumnName.Add metaItem("displayName")
                    If metaItem.exists("columnName") Then ActualColumnName.Add metaItem("columnName")
                    If metaItem.exists("dataType") Then DataTypeDict(metaItem("displayName")) = metaItem("dataType")
                    If metaItem.exists("format") Then FormatDict(metaItem("displayName")) = metaItem("format")
                End If
            End If
        Next
    End If

    ' Check for any keys in records missing from metadata
    Dim record As Object
    Dim key As Variant
    For Each record In DrillsData("records")
        If TypeName(record) = "Dictionary" Then
            For Each key In record.Keys
                If Not ExistsInCollection(DisplayColumnName, CStr(key)) Then
                    Call modFileLogger.FileLogDebug("Found record key missing from metadata: " & key, "ExtractMetadata")
                    DisplayColumnName.Add CStr(key)
                    ActualColumnName.Add CStr(key)
                End If
            Next
        End If
    Next

    Set ExtractMetadata = metaDict
    Exit Function

ErrHandler:
    Call modFileLogger.FileLogError("Metadata Extraction Error: " & Err.Description, "ExtractMetadata")
    Set ExtractMetadata = metaDict
End Function
```

### 11b. Don't fail silently when no columns are found

**Find** (in `DrilldownDataToSheet`):
```vba
    If LastColumn = 0 Then Exit Sub
```

**Replace with:**
```vba
    If LastColumn = 0 Then
        Call modFileLogger.FileLogWarn("No columns could be determined from metadata or record keys (DDType=" & DDType & ")", "DrilldownDataToSheet")
        errorString = "Unable to determine columns for this drilldown. Refer Excel logs for more information."
        GoTo CleanUp
    End If
```

---

## 12. `PublicSubs.bas` — Single-row Journal drilldown shows the raw `~~` helper string as its description

**C# behavior:** `DD_JL.BuildMultiStringSafe` / `BuildMultiString`.

**Root cause:** For a single-row journal drilldown, VBA uses the helper cell's raw value
(`Jan-07~~A~~~~~~12345~~1~~PTD~~USD~~...`) as the description. That text appears in cell **A3** of
the drilldown sheet and is sent as `jobDescription`. C# formats it as
`Period_Name:= Jan-07, Actual_Flag:= A, ..., Ccid:= 12345, ...`.

### 12a. Add function

Add at the end of `PublicSubs.bas`:

```vba
Private Function BuildJournalMultiString(ByVal cellValue As String) As String
    Dim headers As Variant
    Dim parts   As Variant
    Dim i       As Long
    Dim n       As Long
    Dim s       As String

    If InStr(cellValue, "~") = 0 Then
        BuildJournalMultiString = cellValue
        Exit Function
    End If

    headers = Array("Period_Name", "Actual_Flag", "Budget_Version_Id", "Encumbrance_Type_Id", _
                    "Ccid", "Ledger_Id", "Balance_Type", "Currency_Code", "Translated_Flag", _
                    "Source_Name", "Category_Name", "Status", "StartDate", "EndDate")
    parts = Split(cellValue, "~~")

    n = UBound(parts)
    If UBound(headers) < n Then n = UBound(headers)

    For i = 0 To n
        If i > 0 Then s = s & ", "
        s = s & headers(i) & ":= " & parts(i)
    Next i

    BuildJournalMultiString = s
End Function
```

### 12b. Use it

**Find** (in `BuildJournalRequest`):
```vba
        MultiString = ActiveSheet.Cells(rng.Row, FoundCol).value
```

**Replace with:**
```vba
        MultiString = BuildJournalMultiString(CStr(ActiveSheet.Cells(rng.Row, FoundCol).value))
```

---

## 13. `ThisWorkbook.cls` — Journal attachment IDs overflow `CLng`

**C# contract:** `fileIds` is `long[]` (`AllModels.cs`). Same bug class as #6.

**Root cause:** `HandleJournalAttachments` converts each attachment ID with `CLng` (max
2,147,483,647). A larger ID raises "Overflow" and the download never starts.

**Find:**
```vba
                JrAttachIDs.Add CLng(Trim(SpltStr(i)))
```
**Replace with:**
```vba
                JrAttachIDs.Add CDec(Trim(SpltStr(i)))
```

**Find:**
```vba
            JrAttachIDs.Add CLng(AttachIDs)
```
**Replace with:**
```vba
            JrAttachIDs.Add CDec(Trim(AttachIDs))
```

---

## C#-side issues found during the delta audit (not VBA ports)

- **C# regression from `8398ec3`: NOT BETWEEN's second bound keeps its `--`.** The Segment
  Values window writes NOT BETWEEN as `--X|--Y` (`SegmentSelectorViewModel.AddNotBetweenSelection`
  prefixes **both** bounds). The new `CreateRangeSegmentValue` strips only the leading `--` of
  the whole string, so C# now sends `values: ["X", "--Y"]`. The old blanket `Replace("--","")`
  handled this. VBA fix #8 strips both. **Fixed in C# 2026-10-05** (`CreateRangeSegmentValue`, both
  FinalWorkingCode and AIPowered): when `isNotBetween`, also strips a leading `--` from the
  second bound. Build-verified, checked via reflection on the compiled DLL, and pushed: FinalWorkingCode
  11.1.0 `7b460a4`, 11.1.1 `62ce35c`, 11.1.2 `a4d4847` (AIPowered on 11.1.2 only).
- **`~` vs `summaryEnabled`: resolved, VBA was right.** User-confirmed (2026-10-05): a trailing
  `~` means a summary account converted to non-summary, so it must send `summaryEnabled = false`.
  VBA (#8) already did this. C# never had (it always used the DB summary flag). **Fixed in C#
  2026-10-05** in both codebases (`BalanceDtoModel.CreateSingleSegmentValue` →
  `CreateComparisonSegmentValue`). Build- and reflection-verified, and pushed in the same commits as the
  NOT BETWEEN fix above.

## Request-parity audit: VBA vs C# HTTP requests (2026-10-05)

Compared every server call the template makes (URL, query string, method, headers/auth,
timeout, JSON body field by field, and how the response is read back) against
`GLSense\FinalWorkingCode` (11.1.2). Payload types were checked against the C# models
and the 2026-08-21 live captures (`ExcelTemplate.json` / `FromMSI.json`).
Fixes for these are entries #14–#21 below. Confirmed vs. inferred is marked in the table.

### Matches C# (no action)
- Endpoints, query string (`cubeId`, `jobName`, `jobDescription`) and method for all 8 drilldown
  types, plus custom-drilldown, journal-attachment-files, journal-attachments, finance-cubes,
  drilldown-processes, drilldown-data and applogout. `jobDescription` URL-encoding differs in style only
  (`+` vs `%20`).
- `Authorization: Bearer` and `Content-Type` headers.
- Journal body (`journalDrilldowns`) field for field, after #2/#7.
- Subledger body (`subledgerDrilldowns`) field for field.

### Differences, highest impact first

| # | Area | VBA | C# | Effect | Status |
|---|---|---|---|---|---|
| R1 | Response cleanup (`gethttp`) | `Replace(resp,"null","")` **case-insensitive, everywhere** | Case-sensitive replace, then every empty `"key":` becomes `"key":""` | `"x":null` → VBA `Null` → **`Invalid use of Null`** in `ApplyFormatting`. That skips the rest of that column, so its **custom-drilldown hyperlinks and subtotals are lost**. Data containing "null" is corrupted (`Annulled invoice` → `Aned invoice`, `NULL value` → ` value`). | **Confirmed**: reproduced with the template's own `ParseJsonFast`. Matches every balance drilldown in the user logs. |
| R2 | Timeout (`gethttp`) | 30 s connect/send/receive, no retry | 5 min, with retry on transient failures | Large drilldowns fail with `The operation timed out` in VBA but complete in C#. | **Confirmed** in user logs (2026-08-28). |
| R3 | Error responses (`gethttp`) | Non-200: body discarded, only `HTTP 400: Bad Request` shown | Body passed through, so the server's message is shown | Users see a generic HTTP error with no reason. | Code-confirmed |
| R4 | `segmentValueSetId` (Balance) | String `"50742"` | Number `50742` | Type mismatch. The server may coerce it; not verified against the server. | **Confirmed** in the 08-21 capture, still the case in the current code |
| R5 | `cubeId` (attachment list + download bodies) | String `"123"` (`CubeID As String`) | Number | Same as R4 | Code-confirmed |
| R6 | `encumbranceTypeIdList` (Balance, `E`/`A+E`) | Never sent | IDs looked up by encumbrance name | Encumbrance balance drilldowns may return different or no data. VBA's metadata sheet has no encumbrance table, so a fix needs that data first. | Code-confirmed; server impact unverified |
| R7 | Job **Logs** button | `POST /web/secure/monitoringViewLogDownload?processId=`, read as text | `GET /web/secure/schedule/{id}/log-zip`, a zip download | Different endpoint. The VBA one is likely legacy. | Code-confirmed; whether the old endpoint still exists is unverified |
| R8 | Attachment / zip download save path | `%HOMEDRIVE%%HOMEPATH%\Downloads` | `%USERPROFILE%\Downloads` | On domain PCs with a network home drive, the VBA path doesn't exist, so the download fails with a "Stream Error". A missing `Content-Disposition` also crashes VBA, where C# falls back to `DownloadedFile.zip`. | Code-confirmed |
| R9 | Custom-drilldown `COLUMN` values | Cell `.Value` (dates as locale text, e.g. `10/5/2026`), quotes stripped | Cell `Value2` (dates as serial numbers, e.g. `45570`) | Date parameters reach the server in a different form. | Code-confirmed |
| R10 | Actual flag long forms | Only `A`/`B`/`E`/`A+E` handled | Also `BUDGET`/`ENCUMBRANCE`/`ACTUAL+ENCUMBRANCE` | With a long form, VBA sends no `budgetName`/`encumbranceName`. The Balance Configurator only writes short forms, so this only affects hand-typed formulas. | Code-confirmed |
| R11 | ~~Combined segments~~ | ~~A first segment argument containing `;` is sent as one value~~ | | **False finding.** `CleanParameter` already expands it. See #16. | Withdrawn |
| R12 | Unmatched ledger name | Whole request aborted ("Error building balance request") | `ledgerIdList` omitted, request still sent | Different failure mode | Code-confirmed |
| R13 | Journal row selection | Skips a row when the *selected amount cell* is blank | Skips a row when the *helper* cell is blank | Edge-case row differences | Code-confirmed |

Minor, no action: VBA always sends `segments` (possibly `[]`) and `budgetName`/`encumbranceName`
as `""`, where C# omits nulls. The single-row journal description is cosmetic (#12).

**Also a C# issue:** C#'s own `Replace("null","")` (R1) is case-sensitive but still runs inside
string values, so lowercase "null" in data (e.g. "annulled") is corrupted there too.

---

## Request-parity fixes (2026-10-05): entries #14–#21

These close R1–R5, R8–R10 and R13 from the request-parity audit above. R11 was a false
finding (see #16). R6 is closed with no action (the formula is the source of truth). R7 and R12 are
**not** ported. See *Not ported* at the end of this section.

**How these were tested:** every new function and every multi-line replacement below was
compiled and run in a scratch Excel workbook, together with the template's own
`JsonConverter.bas` (extracted unchanged), before anything was applied to the template.
- `gethttp` was run against a local 127.0.0.1 test server: 200 with nulls, 400 with a JSON
  error body, 500 with an HTML body, a 40-second response, and a refused port.
- The other functions ran through 31 unit cases, all passing.

The one-line edits (#18, #21's call site) were checked for exact-once matches against the
current VBA.

---

## 14. `PublicFunctions.bas` — `gethttp`: corrupted responses, 30 s timeout, no retry, server errors hidden (R1, R2, R3)

**Root causes:**
- **R1:** `gethttp = Replace(gethttp, "null", "")` uses the project's own `Replace`, which is
  case-insensitive, and runs over the whole response. `"x":null` becomes `"x":`, so the parser
  produces a VBA `Null`. Assigning that to a String fails with **"Invalid use of Null"** in
  `ApplyFormatting`, which skips the rest of that column (custom-drilldown hyperlinks and the
  subtotal row are lost). It also corrupts data: `Annulled invoice` becomes `Aned invoice`, and
  `NULL value` becomes ` value`. Reproduced with the template's own `ParseJsonFast`.
- **R2:** `setTimeouts 30000, 30000, 30000, 30000` means any drilldown taking longer than 30 s fails
  with "The operation timed out" (seen in the user logs). C# waits 5 minutes and retries connection
  failures up to 3 times.
- **R3:** on a non-200 response the body is discarded, so the user never sees the server's message.

**Fix:** replace the **whole** `gethttp` function (from `Public Function gethttp(` through its
`End Function`) with the block below. It includes three new private helpers. The new null
cleanup matches C# `ApiHelper.CleanResponse` after this session's C# fix (commits `e912b5e` /
`58aedb3` / `e012f5c`): nulls become `""`, and string data is never touched.

```vba
Public Function gethttp(serviceURL As Variant, ByVal StrContentType As String, Optional PostData As String = "", Optional MethodType As String = "POST") As String
    ' Mirrors GLSense (C#) ApiHelper.ServerAPI:
    '  - 5 minute receive timeout (C#: HttpClient.Timeout = 5 min) instead of 30 seconds
    '  - up to 3 attempts, 1s/2s backoff, ONLY for connection-level failures (C#:
    '    ApiOperationHelper.ExecuteWithRetry / IsTransientError) - a timeout is not retried
    '  - JSON null literals cleaned without touching string data (C#: CleanResponse)
    '  - the server's own error text is shown for a non-200 response
    Const MAX_ATTEMPTS As Long = 3
    Dim hReq    As MSXML2.ServerXMLHTTP60
    Dim attempt As Long
    Dim errNum  As Long
    Dim errDesc As String
    Dim body    As String

    Call modFileLogger.FileLogDebug(serviceURL, "API URL")
    Call modFileLogger.FileLogDebug(MethodType, "Method")
    Call modFileLogger.FileLogDebug(PostData, "Payload")

    RespStr = ""

    For attempt = 1 To MAX_ATTEMPTS
        Set hReq = New MSXML2.ServerXMLHTTP60
        ' resolve, connect, send, receive (ms)
        hReq.setTimeouts 30000, 60000, 60000, 300000

        On Error Resume Next
        hReq.Open MethodType, serviceURL, False
        If StrContentType = "JSON" Then
            hReq.setRequestHeader "Content-Type", "application/json"
        Else
            hReq.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        End If
        hReq.setRequestHeader "Authorization", "Bearer " & OrbitAuthToken
        If Len(PostData) = 0 Then
            hReq.send
        Else
            hReq.send PostData
        End If
        errNum = Err.Number
        errDesc = Err.Description
        On Error GoTo 0

        If errNum = 0 Then Exit For

        Call modFileLogger.FileLogError("Err " & errNum & ": " & errDesc & " (attempt " & attempt & " of " & MAX_ATTEMPTS & ")", "gethttp")
        Set hReq = Nothing

        If attempt = MAX_ATTEMPTS Or Not IsTransientHttpError(errNum) Then
            errorString = errorString & vbCrLf & "- gethttp System Error: Err " & errNum & ": " & errDesc
            Exit Function
        End If

        Application.Wait Now + TimeSerial(0, 0, attempt)
    Next attempt

    On Error GoTo ErrHandler
    body = hReq.responseText

    If hReq.status = 200 Then
        gethttp = CleanJsonNulls(body)
    Else
        Dim httpErr As String
        Dim serverMsg As String
        httpErr = "HTTP " & hReq.status & ": " & hReq.StatusText
        serverMsg = ServerErrorText(body)
        Call modFileLogger.FileLogError(httpErr & " for URL: " & serviceURL & " | Response: " & Left$(body, 2000), "gethttp")
        If Len(serverMsg) > 0 Then httpErr = httpErr & " - " & serverMsg
        errorString = errorString & vbCrLf & "- Web Error: " & httpErr
    End If

    Set hReq = Nothing
    Exit Function

ErrHandler:
    Set hReq = Nothing
    Dim vbaErr As String: vbaErr = "Err " & Err.Number & ": " & Err.Description
    Call modFileLogger.FileLogError(vbaErr, "gethttp")
    errorString = errorString & vbCrLf & "- gethttp System Error: " & vbaErr
End Function

''' <summary>
''' Replaces JSON null literals the way GLSense (C#) ApiHelper.CleanResponse does
''' ("x":null -> "x":"", [null] -> []), but ONLY outside string values, so data that
''' merely contains the text "null" (e.g. "Annulled", "NULL") is never changed.
''' Uses native RegExp passes so multi-MB responses stay fast. ChrW(&HFFFF) is a Unicode
''' noncharacter used as a temporary marker - it never occurs in real data.
''' </summary>
Private Function CleanJsonNulls(ByVal json As String) As String
    Dim re   As Object
    Dim mark As String
    Dim s    As String

    If Len(json) = 0 Or InStr(1, json, "null", vbBinaryCompare) = 0 Then
        CleanJsonNulls = json
        Exit Function
    End If

    mark = ChrW(&HFFFF&)
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True

    ' 1. Every string literal is kept exactly ($1); a null literal outside strings
    '    (where $1 is empty) is replaced by the marker alone.
    re.Pattern = "(""[^""\\]*(?:\\.[^""\\]*)*"")|null"
    s = re.Replace(json, "$1" & mark)

    ' 2. A null that was an object VALUE becomes an empty string, like C#'s colon pass.
    re.Pattern = ":(\s*)" & mark
    s = re.Replace(s, ":$1""""")

    ' 3. Drop the remaining markers (after strings, and nulls inside arrays).
    CleanJsonNulls = VBA.Replace(s, mark, "")
End Function

''' Connection-level failures worth retrying (WinHTTP codes as returned by MSXML).
Private Function IsTransientHttpError(ByVal errNum As Long) As Boolean
    Select Case errNum
        Case -2147012867, _
             -2147012866, _
             -2147012865, _
             -2147012744
            ' 12029 cannot connect, 12030 connection aborted, 12031 connection reset,
            ' 12152 invalid server response
            IsTransientHttpError = True
        Case Else
            IsTransientHttpError = False
    End Select
End Function

''' Best-effort extraction of the server's own error text from a non-200 body.
Private Function ServerErrorText(ByVal body As String) As String
    On Error GoTo Fail
    Dim d As Object
    Dim k As Variant

    If InStr(body, "{") = 0 Then Exit Function
    Set d = JsonConverter.ParseJsonFast(CleanJsonNulls(body))
    If d Is Nothing Then Exit Function
    If TypeName(d) <> "Dictionary" Then Exit Function

    For Each k In Array("msg", "msg1", "message", "error")
        If d.exists(k) Then
            If Not IsObject(d(k)) Then
                If Len(CStr(d(k))) > 0 And CStr(d(k)) <> "Failed to parse JSON" Then
                    ServerErrorText = CStr(d(k))
                    Exit Function
                End If
            End If
        End If
    Next k
Fail:
End Function
```

**Verified:**

```text
CleanJsonNulls  11/11 PASS  ("annulled"/"NULL x"/"null" in strings untouched, escaped quotes,
                            key named "null", [null] -> [], "x":null -> "x":"",
                            parsed customFormula is now a String, record data intact)
                5.4 MB / 20,000-record response cleaned in 0.11 s
gethttp         200 -> cleaned body                                          PASS
                400 {"status":"failed","msg":"Ledger 'X' not found"}
                    -> "- Web Error: HTTP 400: Bad Request - Ledger 'X' not found" PASS
                500 HTML body -> "- Web Error: HTTP 500: Internal Server Error"  PASS
                40 s response -> succeeds (old code: timed out at 30 s)       PASS
                refused port -> 3 attempts logged, then reported (8.7 s)      PASS
```

**Difference from C#, on purpose:** C# passes a non-200 JSON body back to the caller. VBA
puts the server's message in `errorString`, which every caller already shows (e.g. "Balances
Drilldown Error: Empty response / - Web Error: HTTP 400: Bad Request - Ledger 'X' not found").
That avoids the same message being shown twice through `ReturnJsonDict`.

---

## 15. `JSONBuilder.bas` — `segmentValueSetId` sent as a string (R4)

**Root cause:** `NewBuildSegJSON` assigns the value-set ID as text, so the payload has
`"segmentValueSetId": "50742"`. C# sends the number `50742` (confirmed in the 2026-08-21 captures).

### 15a. Add function (anywhere in `JSONBuilder.bas`)

```vba
Private Function SegSetIdValue(ByVal SegValID As String) As Variant
    ' segmentValueSetId is a number in the C# payload (long, 0 when unknown)
    If Len(Trim(SegValID)) > 0 And IsNumeric(SegValID) Then
        SegSetIdValue = CDec(SegValID)
    Else
        SegSetIdValue = CDec(0)
    End If
End Function
```

### 15b. Use it

**Find** (in `NewBuildSegJSON`):
```vba
    NewJSON("segmentValueSetId") = SegValID
```
**Replace with:**
```vba
    NewJSON("segmentValueSetId") = SegSetIdValue(SegValID)
```

Verified: `ConvertToJson` output is `"segmentValueSetId":50742`, and an unknown ID gives `0` (the C# default).

---

## 16. `JSONBuilder.bas` — segment argument read without a bounds check

**Root cause:** `BalanceJson` read `FuncParam(10 + idx + 1)` for every segment the ledger has,
without a bounds check. A formula with fewer segment arguments than the ledger has segments threw
"Subscript out of range" (reported as "Error building balance request"). C# processes only
`Min(segment count, argument count)`.

**Correction:** this entry first also claimed (as R11) that VBA didn't split a `;`-combined
segment argument (`"101;10;11200;000"`). That was wrong: `PublicFunctions.CleanParameter`
already expands it into separate parameters before `BalanceJson` sees it. The fix applied is only
the bounds check.

### 16a. Add function (anywhere in `JSONBuilder.bas`)

```vba
''' <summary>
''' Returns the formula's segment argument for segment position idx (0-based), or "" when the
''' formula has fewer segment arguments than the ledger has segments. Mirrors GLSense (C#)
''' BalanceDtoModel.ProcessSegments, which only processes Min(segment count, argument count).
''' (A ";"-combined segment argument is already expanded into separate parameters by
''' TrapParameters/CleanParameter before BalanceJson sees it.)
''' </summary>
Private Function SegmentArg(ByVal FuncParam As Variant, ByVal idx As Long) As String
    If 11 + idx <= UBound(FuncParam) Then
        SegmentArg = VBA.Replace(FuncParam(11 + idx), Chr(34), "")
    End If
End Function
```

### 16b. Use it

**Find** (in `BalanceJson`):
```vba
                            If Len(Replace(FuncParam(10 + (idx + 1)), Chr(34), "")) > 0 Then
                               Set D2 = NewBuildSegJSON(Replace(FuncParam(10 + (idx + 1)), Chr(34), ""), (idx + 1), IDsArray(idx))
                               C2.Add D2
                            End If
```
**Replace with:**
```vba
                            Dim segArg As String
                            segArg = SegmentArg(FuncParam, idx)
                            If Len(segArg) > 0 Then
                               Set D2 = NewBuildSegJSON(segArg, (idx + 1), IDsArray(idx))
                               C2.Add D2
                            End If
```

Verified: individual arguments give `101|10|11200|` (a missing 4th argument is empty, with no
error); a quoted `"~118"` with no later arguments gives `~118||`; a formula with no segment
arguments gives empty, with no error.

---

## 17. `JSONBuilder.bas` — long-form actual flags (`BUDGET` / `ENCUMBRANCE` / `ACTUAL+ENCUMBRANCE`) (R10)

**Root cause:** VBA only recognized `A`/`B`/`E`/`A+E`. With a long form, `budgetName` and
`encumbranceName` were never set, so they were missing from the payload. C# accepts both forms.

**Find** (in `BalanceJson`):
```vba
                    If Len(FuncParam(7)) > 0 Then
                        If FuncParam(7) = "A" Then
                            D1("budgetName") = ""
                            D1("encumbranceName") = ""
                        ElseIf FuncParam(7) = "B" Then
                            D1("budgetName") = FuncParam(8)
                            D1("encumbranceName") = ""
                        ElseIf FuncParam(7) = "E" Or FuncParam(7) = "A+E" Then
                            D1("budgetName") = ""
                            D1("encumbranceName") = FuncParam(8)
                        End If
                    Else
                        D1("budgetName") = ""
                        D1("encumbranceName") = ""
                    End If
```
**Replace with:**
```vba
                    Select Case FuncParam(7)
                        Case "B", "BUDGET"
                            D1("budgetName") = FuncParam(8)
                            D1("encumbranceName") = ""
                        Case "E", "ENCUMBRANCE", "A+E", "ACTUAL+ENCUMBRANCE"
                            D1("budgetName") = ""
                            D1("encumbranceName") = FuncParam(8)
                        Case Else
                            D1("budgetName") = ""
                            D1("encumbranceName") = ""
                    End Select
```

Verified: `A`, `B`, `BUDGET`, `E`, `ACTUAL+ENCUMBRANCE` and blank all map correctly (6/6).

---

## 18. `ThisWorkbook.cls` — `cubeId` sent as a string in both attachment requests (R5)

`CubeID` is declared `As String`, so both bodies have `"cubeId":"123"`. C# sends a number.

**Find** (in `HandleJournalAttachments`):
```vba
    JsonObj.Add "cubeId", CubeID
```
**Replace with:**
```vba
    JsonObj.Add "cubeId", CDec(CubeID)
```

**Find** (same procedure):
```vba
        DownloadObj.Add "cubeId", CubeID
```
**Replace with:**
```vba
        DownloadObj.Add "cubeId", CDec(CubeID)
```

Verified: serializes as `"cubeId":123`.

---

## 19. `ThisWorkbook.cls` — attachment download saved to the wrong folder / crashes without a file name (R8)

**Root cause:** `DownloadZipWeb` saves to `%HOMEDRIVE%%HOMEPATH%\Downloads`. On domain PCs with a
network home drive that folder usually doesn't exist, so the download fails with "Stream Error". C#
uses `%USERPROFILE%\Downloads`. A response without a `Content-Disposition` filename also threw
"Subscript out of range", where C# falls back to `DownloadedFile.zip`. (The old
`If filePath = ""` fallback could never fire, because `"\Downloads"` was always appended first.)

### 19a. Add function (anywhere in `ThisWorkbook.cls`)

```vba
''' File name from a Content-Disposition header; "DownloadedFile.zip" when the header or its
''' filename is missing (same fallback as GLSense (C#) AddinModule.GLSense_DownloadFile).
Private Function DownloadFileName(ByVal contentDisposition As String) As String
    Dim cd  As String
    Dim pos As Long

    cd = VBA.Replace(contentDisposition, Chr(34), "")
    pos = InStr(1, cd, "filename=", vbTextCompare)
    If pos > 0 Then
        cd = Mid$(cd, pos + Len("filename="))
        If InStr(cd, ";") > 0 Then cd = Left$(cd, InStr(cd, ";") - 1)
        DownloadFileName = Trim$(cd)
    End If
    If Len(DownloadFileName) = 0 Then DownloadFileName = "DownloadedFile.zip"
End Function
```

### 19b. Use it

**Find** (in `DownloadZipWeb`):
```vba
   fileName = Split(Replace(hReq.getResponseHeader("Content-Disposition"), Chr(34), ""), "filename=")(1)
   filePath = Environ("HOMEDRIVE") & Environ("HOMEPATH") & "\Downloads"
   If filePath = "" Or Len(filePath) = 0 Then
      filePath = AppDataLoc
   End If
```
**Replace with:**
```vba
   Dim contentDisp As String
   On Error Resume Next
   contentDisp = hReq.getResponseHeader("Content-Disposition")
   On Error GoTo ErrHandler
   fileName = DownloadFileName(contentDisp)
   filePath = Environ("USERPROFILE") & "\Downloads"
   If Len(Dir(filePath, vbDirectory)) = 0 Then
      filePath = AppDataLoc
   End If
```

Verified `DownloadFileName` (4/4): quoted name, `FileName=` in any case with a trailing
`; size=…`, missing header, and header without a filename.

---

## 20. `ThisWorkbook.cls` — custom-drilldown parameter values read with `.Value` instead of `.Value2` (R9)

**Root cause:** C# (`UpdateColumnParameter`) sends `Value2`, so a date goes as its serial number
and a number goes without its display format. VBA used `.Value`, which sends a date as
locale text such as `10/5/2026`, and also stripped `"` from the value. The server therefore got
a differently formatted date than C# sends.

**Find** (in `GLSenseCustomDrillDown`):
```vba
            If colIdx > 0 Then
                p("value") = Replace(CStr(ActiveSheet.Cells(rng.Row, colIdx).value), Chr(34), "")
            End If
```
**Replace with:**
```vba
            If colIdx > 0 Then
                Dim colCellVal As Variant
                ' Value2, same as GLSense (C#) UpdateColumnParameter: dates go as their serial
                ' number and values are not reformatted by the user's locale/number format
                colCellVal = ActiveSheet.Cells(rng.Row, colIdx).Value2
                If IsError(colCellVal) Then
                    p("value") = ""
                Else
                    p("value") = CStr(colCellVal)
                End If
            End If
```

Verified (4/4): a date becomes its serial, `1234.5` formatted `#,##0.00` is sent as `1234.5`,
`#N/A` becomes `""`, and text with quotes is kept as-is (ConvertToJson escapes it). The sheet
header (cell A3) built from these values now shows the same serial-number form C# shows.

---

## 21. `JSONBuilder.bas` — Journal rows chosen by the amount cell instead of the helper cell (R13)

**Root cause:** `JournalJson` includes a row when the *selected amount cell* is non-empty. C#
includes it when the row's `~~` helper cell is non-empty. So a blank amount on a valid row was
skipped, and a row with an amount but no helper produced an empty journal record.

### 21a. Add function (anywhere in `JSONBuilder.bas`)

```vba
''' Cell value as text; "" for an error value (#N/A etc.) instead of a type-mismatch error.
Private Function CellText(ByVal v As Variant) As String
    If IsError(v) Then
        CellText = ""
    Else
        CellText = CStr(v)
    End If
End Function
```

### 21b. Use it

**Find** (in `JournalJson`):
```vba
        If Len(CR.value) > 0 Then
```
**Replace with:**
```vba
        If Len(CellText(ActiveSheet.Cells(CR.Row, ColFound).value)) > 0 Then
```

Verified `CellText` (2/2): an error value gives `""` instead of a type mismatch, and text is returned unchanged.

---

### Not ported (need a decision or data first)

- **R6 `encumbranceTypeIdList`: closed, no action (user decision, 2026-10-05).** The template's
  metadata sheet has no encumbrance or journal metadata, so the formula is the source of truth:
  VBA sends the formula's `encumbranceName` (balances), and journal drilldowns take
  `encumbranceTypeId` straight from the formula's `~~` helper cell (#2/#7). VBA has no ID lookup to do.
- **R7 Logs button endpoint:** switching to C#'s `GET /web/secure/schedule/{id}/log-zip` changes
  the feature from "show text" to "download a zip". Do this only after confirming the old
  `monitoringViewLogDownload` endpoint is gone.
- **R12 unmatched ledger:** VBA stops with "Error building balance request". C# sends the
  request without `ledgerIdList`. VBA's clear local error is arguably better than a server
  error, so it's left as is.

---

## Applied 2026-10-05 — commit `0040acf` (OISR-22571 / OISR-22575)

### How it was applied

- Every fix (#8–#26) was applied by Excel automation (pywin32) as exact text edits on the live
  module code. Each edit had to match exactly once, or the run aborted without saving.
- After writing, every module was read back and checked: each expected new line present, each
  old buggy line gone.
- The whole project was then compiled (*Debug > Compile VBAProject*, with a watchdog that captures
  any compile-error dialog). The file was saved only after a clean compile.
- An independent re-extraction (olevba) diffed against the original showed changes only in the
  intended modules.
- The template was re-signed with `SignTemplate.bat`: "Successfully signed", 0 warnings,
  0 errors. `signtool verify /pa` gives "Successfully verified" (DigiCert RFC 3161 timestamp),
  Excel reports `VBASigned = True`, and the signed file's VBA is identical to the validated version.

### #22–#26: found while compiling and validating

The template **had never compiled as a whole project**. Excel compiles on demand, which hid
three errors (#22–#24) until a full compile was run. Any user with "Compile On Demand" turned off,
or any code path that reached the broken procedure, would get a VBA compile error.

| # | Module | Problem | Fix |
|---|---|---|---|
| 22 | `PublicFunctions.bas` | `GetLedgerArrays` uses `LedgerRegistry`, which is declared nowhere ("Variable not defined"). It has no callers. | Removed the dead function |
| 23 | `JsonConverter.bas` | `closeBrackets()` declares a local `closeBrackets As Long` with its own name ("Duplicate declaration"), so the `CleanJsonAggressive` JSON-repair path could never compile | Renamed the local counter to `closeBracketCount` (inside that function only) |
| 24 | `clsArrayList.cls` | A stray line containing only two backticks (a pasted markdown fence) after `Remove` caused a syntax error | Removed the line |
| 25 | `modFileLogger.bas` | `AppendLine` runs `On Error GoTo`, which resets the global `Err`. About 17 handlers log first and then build the user's message from `Err.Description`, so users saw an **empty reason**, e.g. `"Error building balance request: "` or `"- DrilldownDataToSheet Error: "` | Routed all four `FileLog*` calls through a new `WriteLog` that saves and restores `Err` |
| 26 | `DrilldataToWorksheet.bas` | **Root cause of OISR-22575.** Output files can be saved with sheets **grouped** (the hidden `OrbitSchMetadata` sheet selected together with the report sheet). `Worksheets.Add` then fails with "This action won't work on multiple selections", and the user can't ungroup by clicking a tab when the report sheet is the only visible one. | `PrepareWorksheet` now selects only the active sheet first when more than one sheet is selected |

#25, the new logger entry points:

```vba
Public Sub FileLogDebug(ByVal msg As String, Optional ByVal ctx As String = "")
    If Not DebugLogsEnabled Then Exit Sub
    WriteLog "DEBUG", msg, ctx
End Sub

Public Sub FileLogInfo(ByVal msg As String, Optional ByVal ctx As String = "")
    WriteLog "INFO ", msg, ctx
End Sub

Public Sub FileLogWarn(ByVal msg As String, Optional ByVal ctx As String = "")
    WriteLog "WARN ", msg, ctx
End Sub

Public Sub FileLogError(ByVal msg As String, Optional ByVal ctx As String = "")
    WriteLog "ERROR", msg, ctx
End Sub

' Writing a line runs On Error statements (AppendLine and friends), and executing an On Error
' statement resets the global Err object. Error handlers all over the project log FIRST and then
' build the user's message from Err.Description, so without this the user saw an empty reason
' (e.g. "Error building balance request: "). The caller's Err is saved and restored here.
Private Sub WriteLog(ByVal level As String, ByVal msg As String, ByVal ctx As String)
    Dim errNum  As Long
    Dim errDesc As String
    Dim errSrc  As String
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    EnsureLogger
    AppendLine Stamp(level, msg, ctx)

    If errNum <> 0 Then
        Err.Number = errNum
        Err.Description = errDesc
        Err.Source = errSrc
    End If
End Sub
```

#26, inserted in `PrepareWorksheet` right after its `FileLogDebug("Initializing worksheet: ...")` line:

```vba
    ' Ungroup first: with several sheets selected (e.g. the hidden OrbitSchMetadata sheet grouped
    ' with the report sheet), adding a sheet or table fails with "This action won't work on
    ' multiple selections". Keep only the active sheet selected.
    If ActWB.Windows(1).SelectedSheets.Count > 1 Then
        Call modFileLogger.FileLogWarn("Sheets were grouped (" & ActWB.Windows(1).SelectedSheets.Count & ") - selecting only the active sheet", "PrepareWorksheet")
        ActWB.ActiveSheet.Select True
    End If
```

#25 was proven with the logger's real code: after logging inside an error handler, the original
left `Err.Number=0, Description=[]`; the fixed version keeps `11, [Division by zero]`.

### Validation on the real sample (`FSG_Test (8).xlsm`)

The sample carried a much older VBA build: none of #1–#7, and no `clsArrayList`. Its VBA was
brought in line with the fixed template: 12 modules synced, `clsArrayList` imported, and its own
`Sheet1` module kept. It compiles cleanly. The workbook's own code was then run on throwaway
copies, before and after:

| Check | Original sample | Updated sample |
|---|---|---|
| F3 (`~118`) balance payload | `values ["118"]` (OISR-22571) | `values ["~118"]` |
| F4 (`11$9`) / F5 (`!@#2`) | correct values | correct values |
| `ledgerIdList` / `segmentValueSetId` | `["1"]` / `"1002472"` (strings) | `[1]` / `1002472` (numbers) |
| `isFunctionalCurrency` | missing | `true` |
| Drilldown response with nulls, "Annulled", no `metadata` (local test server) | response corrupted (`Aned invoice`, `"NOTE":`); nothing written, no message | clean response; drilldown sheet created with `JE_NAME / AMOUNT / NOTE / FLAG` and the data intact |
| File opened with sheets grouped | (not reached) | handled by #26 |

The user then tested the updated sample against the real server and confirmed it works. The
updated sample is in `Downloads\FSG_Test (8).xlsm` (unsigned, because `SignTemplate.bat` only
signs the template). The original is next to it as `FSG_Test (8) - ORIGINAL backup.xlsm`.
Output files generated from the old template keep their old VBA until they're regenerated.

### Still open

- **R7:** the jobs-monitor Logs button calls `POST /web/secure/monitoringViewLogDownload`, while C#
  uses `GET /web/secure/schedule/{id}/log-zip`. Left unchanged pending a decision.
- **R12:** an unmatched ledger name stops VBA locally. That's left as is on purpose.

---

## Checked — not a bug, no action needed

- **Activity ShortName normalization**: C# resolves the Activity formula argument against a
  DisplayName/ShortName lookup before sending it; VBA sends the raw argument untouched, with
  no equivalent lookup available (VBA's metadata engine has no Activity table at all).
  **Confirmed by the user**: the hidden metadata/config sheet is prepopulated with
  already-correct Activity values, so no VBA-side resolution is needed. Not a gap.
- All 8 DDType dispatch entries (`BL, BL_JL, BL_SL, JL, BLDD_SL, BLDD_UF, SL, UF`) are wired
  identically on both sides.
- `SubLedgerJson` matches the C# SubLedger payload field-for-field.
- ~~Segment-token parsing matches exactly between both sides.~~ This was true on 2026-08-17 but is superseded by C# `8398ec3` (2026-09-23). See #8.
- The 7-step worksheet-writer pipeline (metadata → array → sheet → formats → populate →
  formatting → finalize) matches step-for-step, including subledger styling, attachment
  hyperlinks, subtotal codes, hidden columns, and tab coloring.
- The `ORB_XLDD_` (VBA) vs `ORB_DD_` (C#) table-name-prefix difference is not a bug — each side
  generates and checks its own consistent prefix.
- Jobs Monitor's apparent "FinanceSnapshotJob" gap is a non-issue — VBA structurally can never
  queue that job type in the first place.
- `CanProceed` (VBA) vs `GuardValidInputs` (C#) cover equivalent ground, just split differently.

## Separately fixed in C# (not a VBA port — the add-in was the buggy side here)

`DDDatatoWorksheet.cs`'s formula-injection guard (`ProtectExcelFormulaLikeText`) only quoted a
value when it contained **more than one** `=` character, missing the actual injection pattern
(a single leading `=`, e.g. `=HYPERLINK(...)`). VBA's own guard was already correct
(`Left(strVal,1) = "="`). Fixed in `FinalWorkingCode`'s `DDDatatoWorksheet.cs` by checking
`text.TrimStart().StartsWith("=")` instead — already built and verified, no VBA action needed.
