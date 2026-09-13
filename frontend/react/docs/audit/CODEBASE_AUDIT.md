# WORD — Full Codebase Audit Report

**Scope:** every line of `frontend/react/src` (~11,300 lines TS/TSX), CSS, tests, config.
**Verification:** `npm run build` ✅ · `vitest` 54/54 ✅ · `eslint --max-warnings 0` ✅ — so **all bugs below are runtime/logic bugs**, invisible to the compiler.

Legend: 🔴 P0 breaks documents / data loss · 🟠 P1 feature silently does the wrong thing · 🟡 P2 minor / cosmetic.

---

## A. Engine — text editing & selection (`engine/DocumentEngine.ts`)

| # | Sev | Bug | Detail |
|---|-----|-----|--------|
| A1 | 🔴 | **Replace All infinite loop** | `replaceAllText` loops `find → replaceText → find`. If `find` matches the new `replace` text (replace `"x"` with `"xx"`, or `"the"` with `"the "`), every pass finds the fresh replacements again → the browser hangs and the document is lost. |
| A2 | 🔴 | **Replace replaces the wrong occurrence** | `replaceText` ignores the dialog's `currentResult`; it replaces the *first* match in the document, so “Find Next → Replace” walks the first occurrence repeatedly. |
| A3 | 🔴 | **Formatting / clear / change-case / applyStyle ignore the selection** | `applyFormattingToSelection`, `clearFormatting`, `changeCase`, `applyStyle` iterate **every paragraph in the document**. Select one word → Bold → *entire document* turns bold. Same for font, size, color, highlight, case, style, superscript/subscript, spacing. |
| A4 | 🔴 | **AutoCorrect corrupts text across runs** | In `insertText`, the trigger length is subtracted from `cursor.offset` of the *current run*, but the trigger text may live in a previous run. Characters from the wrong position get deleted. Also it fires even when typing mid-word. |
| A5 | 🔴 | **Delete-selection leaves empty paragraphs with stale cursor** | `deleteSelectionInternal` wipes text but never merges paragraphs; the caret can end up on a paragraph that no longer contains the cursor run (rendered caret drifts/ghosts). |
| A6 | 🟠 | **Hyperlink applies to the whole document** | `insertHyperlink` with a selection loops *all sections/blocks/runs* and stamps the URL on everything, not just the selected range. |
| A7 | 🔴 | **Comments tag every run in the document** | `addComment` with a selection sets `commentIds` on **every run in every paragraph**; the highlight then covers the entire document, not the commented range. |
| A8 | 🟠 | **selectLine / selectParagraph produce wrong selection** | They select `runIndex: 0, offset: wordStart` … `runIndex: lastRun, offset: textLength` — but `textLength` is the *whole paragraph length*, so the end offset is past the last run's end; selectWord uses paragraph-relative offsets but passes them as run-0 offsets. |
| A9 | 🟠 | **`tOGGLEcASE` is fake** | It uppercases every even character — real toggle-case flips the existing case of each letter. |
| A10 | 🟠 | **`sentenceCase` regex is broken** | The pattern `/^\s*[.!?]\s*[a-z]|[a-z]/` doesn't capture what the callback assumes; results are wrong for common sentences. |
| A11 | 🟠 | **setOrientation double-swaps** | It swaps `pageSetup` dims, then inside the section loop checks `sp.orientation` *before* assigning the new value — after the first iteration sections that were already landscape swap twice. |
| A12 | 🟠 | **insertTableWithData picks the wrong table** | After `insertTable`, it scans for *any* empty table with matching dims and fills the first one found — not necessarily the one just inserted at the cursor. |
| A13 | 🟠 | **updateTableOfContents duplicates** | It calls `insertTableOfContents`, pushing a new TOC entry + a second set of TOC paragraphs every time (no replace). |
| A14 | 🟠 | **Track Changes is a shell** | Nothing records changes on edit; `acceptChange`/`rejectChange` just delete the change row and do nothing to the document (both branches are empty). Accept All = Reject All = `changes = []`. |
| A15 | 🟠 | **Format Painter never paints** | `startFormatPainter` sets a flag that nothing consumes; there is no apply-on-click, and no way to stop it except the unused `stopFormatPainter`. |
| A16 | 🟠 | **insertParagraph ignores the selection-merge case** | after `deleteSelectionInternal`, `this.cursor.runIndex` may point past the end of remaining runs; `para.textRuns[runIndex]` is undefined → no text inserted until you click again. |
| A17 | 🟡 | **`PageSize` includes ` executive'` with a leading space** and `' executive'` key — dead/typo size. |
| A18 | 🟡 | **Undo stack is heavy** | Every keystroke pushes a full document deep-clone (200 deep). Large docs → serious memory + GC churn. (Fix: coalesce typing bursts.) |
| A19 | 🟡 | **`getSelectedText` ignores selection direction** correctly, but `deleteSelectionInternal` always sets the cursor to `selection.start` even when selection was made right-to-left. |
| A20 | 🟡 | **findText `useWildcard` is not wildcards** — it just runs the query as a raw regex with no Word-style `*`/`?` translation. |

## B. Engine — feature modules

| # | Sev | Bug | Detail |
|---|-----|-----|--------|
| B1 | 🟠 | **Numbered lists all show "1."** | `ParagraphRenderer` calls `engine.getNumberCharacter(level, 0)` — index is hardcoded 0, so every numbered paragraph is “1.”. No numbering context exists at all. |
| B2 | 🟠 | **Page count is fake** | Status bar, word count dialog and File > Info all estimate pages as `words/320` (word count dialog uses `/250`). Two different magic numbers, both wrong. |
| B3 | 🟡 | **Equation/Chart/SmartArt are placeholders** — equation renders raw LaTeX, chart always renders bars regardless of `chartType`, SmartArt ignores its layout. |

## C. Editor canvas (`DocumentCanvas.tsx`)

| # | Sev | Bug | Detail |
|---|-----|-----|--------|
| C1 | 🟠 | **Home/End jump to document start/end** — mapped to `moveCursorToStart/ToEnd` instead of line start/end (the real ones exist but are unused). |
| C2 | 🟠 | **Word-based shortcuts missing** — no Ctrl+E/L/R/J (align), Ctrl+1..5 (styles), Ctrl+Shift+arrow. Ribbon shows “(Ctrl+E)” on the button but the key does nothing. |
| C3 | 🟠 | **Ctrl+] / Ctrl+[ read the wrong base size** — uses `activeFormatting.fontSize ?? 11` even when the selection carries an explicit size, so increasing a selected 24pt text drops it to 13. |
| C4 | 🟡 | **Print CSS missing** — `window.print()` prints the app chrome; there is no `@media print` for the canvas (FileMenu.css has one only for the file menu view). |
| C5 | 🟡 | **Mouse wheel over canvas doesn't zoom with Ctrl** (Word behavior), minor. |
| C6 | 🟡 | **`canvasRef.style.cursor='none'`** hides the mouse cursor over the whole canvas while the caret is visible — intentional (custom caret) but it makes text hover/click UX odd since there's no I-beam fallback. |

## D. App shell / session (`App.tsx`)

| # | Sev | Bug | Detail |
|---|-----|-----|--------|
| D1 | 🔴 | **File > New (FileMenu) nukes the open document with no tab** | `engine.newDocument()` replaces the live engine document directly; tabs list still shows the old tab pointing at the old snapshot; unsaved work is only reachable via Ctrl+Z. Same for template picks and GeneratorDialog “Create document”. |
| D2 | 🟠 | **`closeTab` doesn't save to recents** — closing a dirty unsaved tab silently loses it (no prompt, no recents entry). |
| D3 | 🟠 | **Autosave timer writes only when status is exactly 'unsaved'** — a document left open but never dirty-again stays out of recents; also autosave doesn't create a snapshot label, mixing manual & auto saves. |
| D4 | 🟡 | **beforeunload listener registered only when 'unsaved'** — a save in flight (status 'saving') at close time is not guarded. |
| D5 | 🟡 | **`applyTemplate` sets style AFTER inserting text, then applies to the whole doc** (engine bug A3) — templates visually “work” by accident but the style application also restyles everything else when body text exists. |

## E. Dialogs / panels

| # | Sev | Bug | Detail |
|---|-----|-----|--------|
| E1 | 🟠 | **PageSetupDialog is half-fake** — tabs (Margins/Paper/Layout) do nothing, gutter/header/footer distances are captured but never applied, “wide” preset == “normal”. |
| E2 | 🟠 | **PropertiesPanel shows hardcoded zeros** — right indent / first-line / spacing inputs always display 0 and line-spacing select is locked to “1.15” (`value="1.15"`); changing them writes invisible values to the cursor paragraph only. |
| E3 | 🟠 | **Ruler margin math is wrong** — `x * (1/(96*zoom/100)) * 20 * 10` mixes CSS px, twips and a mystery ×10; dragging margins moves them by ~10× the mouse. Indent handles aren't implemented at all. |
| E4 | 🟠 | **VoiceControl maps `palette.open` → nonexistent `palette.open` registry id** — voice “open command palette” always falls into “Command not recognized”. |
| E5 | 🟠 | **GeneratorDialog provider check `A && B \|\| C`** — with a local provider configured (`isConfigured=true, privacy='device'`), the first clause is false, then `kind !== 'local'` is false → OK; but with `openai` configured and check failing, error path shows raw error instead of falling back to template. Also `applyStyle(i === 0 ? 'Heading1' : 'Heading1')` — both branches identical. |
| E6 | 🟡 | **FindReplaceDialog “Match All Word Forms” checkbox does nothing** (it's not passed to the engine at all). |
| E7 | 🟡 | **StatusBar “Read Mode / Web Layout” buttons are dead** — no onClick. |
| E8 | 🟡 | **WordCountDialog pages uses /250, everything else /320** (see B2). |
| E9 | 🟡 | **FileMenu > Open > Recent says “No recent documents” always** — the actual recents are shown only on the Start page. |
| E10 | 🟡 | **ContextMenu uses emoji icons** (✂ 📋 📄 🔗 ✨) violating the project's own “zero emojis” rule; also `paragraphSettings` item does nothing. |
| E11 | 🟡 | **NavigationPane history/attachments tabs are static placeholders** (“No prior versions”) even when snapshots exist. |
| E12 | 🟡 | **export.test expects 320/page, WordCountDialog 250** — inconsistent tests vs UI. |

## F. Services / storage

| # | Sev | Bug | Detail |
|---|-----|-----|--------|
| F1 | 🟠 | **authStore stores passwords in plaintext localStorage** (self-acknowledged in comments, but it also *pre-fills* the demo reset code in a toast — trivially insecure). |
| F2 | 🟡 | **snapshots key grows unbounded across docs** — per-doc cap (12) but no global cap; localStorage quota errors are swallowed silently. |
| F3 | 🟡 | **Feature flags exist but nothing reads them** — `loadFlags` is only used inside the Settings dialog; every feature is permanently enabled regardless of the toggles. |

---

## What's working ✅

- Build, typecheck, all 54 tests, zero-warning lint.
- Core typing/caret/click/drag-select, undo/redo (single doc), tables insert + render, images, shapes/charts placeholders render, markdown round-trip, smart paste, TF-IDF brain + semantic search, Ask-document RAG flow, local on-device AI tools, timeline snapshots + visual diff, health/cleanup/docTest/design-inspector analytics, theme system, auth flows (client-side), recents + autosave.

---

## Fix plan (applied in this branch)

1. **P0 engine correctness** — Replace All loop, replace-current, selection-scoped formatting/case/clear/style, hyperlink & comment range targeting, AutoCorrect offset fix, orientation, selectLine, insertTableWithData targeting, TOC update, ` executive` typo.
2. **Real features** — Track Changes recording + accept/reject, Format Painter apply, numbered-list numbering context.
3. **UI** — Home/End to line, Ctrl+E/L/R/J & Ctrl+]/[ on selection, voice palette mapping, generator fallback + style bug, page-setup apply all fields, ruler math, PropertiesPanel live values, print CSS, dead buttons.
4. Tests updated where behavior intentionally changed (e.g. replace semantics).

---

# Fix status (applied & verified)

**Verification after fixes: `tsc --noEmit` ✅ · `eslint --max-warnings 0` ✅ · `vitest` 54/54 ✅ · `vite build` ✅**

## Fixed in `engine/DocumentEngine.ts`

| Ref | Fix |
|-----|-----|
| A1 | `replaceAllText` is now a single forward pass — replacing `x`→`xx` can no longer loop. |
| A2 | `replaceText` accepts an `onlyAt` cursor position; the Find & Replace dialog can target the exact current match. |
| A3 | `applyFormattingToSelection` splits partially-selected runs and only touches the selected slice; `changeCase`, `clearFormatting` likewise scoped to the selection. |
| A4 | AutoCorrect now verifies the full trigger against the text before the caret across run boundaries (word-boundary + cross-run replace) — no more mid-word or cross-run corruption. |
| A6 | `insertHyperlink` applies only to the selected range (with run splitting); collapsed caret still inserts a new link run. |
| A7 | `addComment` tags only runs intersecting the selection. |
| A8 | `selectWord` maps paragraph offsets back to run positions; `selectLine` uses the true paragraph end. |
| A9 | `tOGGLEcASE` flips each letter's real case. |
| A10 | `sentenceCase` rewritten with a correct pattern. |
| A11 | `setOrientation` swaps width/height per object based on its current orientation (no double swap). |
| A12 | `insertTableWithData` builds the table itself and fills it by captured id. |
| A13 | `updateTableOfContents` removes the previously generated TOC block run before regenerating. |
| A14 | **Track Changes implemented**: insertions and single-char deletions are recorded with range + content; `rejectChange` reverts text via absolute ranges; `rejectAllChanges` replays in reverse. Accept keeps text, per Word semantics. |
| A15 | **Format Painter implemented**: `startFormatPainter` captures the source run format, `applyFormatPainter` paints it onto the next selection (wired into canvas mouseup) and disarms. |
| A17 | ` executive` → `executive`. |
| — | New helpers: `selectedParagraphIds`, `deleteAbsoluteRange`, `insertAbsoluteText`, `extendSelectionToStartOfLine/ToEndOfLine`, `applyFormatPainter`, `isFormatPainterArmed`. |

## Fixed in the UI

| Ref | Fix |
|-----|-----|
| B1 | Numbered lists count consecutive same-level siblings — real `1. 2. 3.` numbering in the canvas. |
| C1 | Home/End (with/without Shift) now move/extend within the current line. |
| C2 | Ctrl+E/L/R/J paragraph alignment shortcuts added; Ctrl+]/[ now read the selection's run size instead of pending format. |
| C4 | `@media print` added so Ctrl+P prints only the document pages. |
| D1 | FileMenu “New”/template picks dispatch `word:new-document`; AppShell handles it tab-aware (stashes the current doc first). |
| E1 | PageSetupDialog applies gutter, header/footer distances; “wide” preset actually wide. |
| E2 | PropertiesPanel shows live indent/spacing/line-spacing values from the cursor paragraph. |
| E3 | Ruler margin drag converts px→twips correctly (removed the ×10 bug). |
| E4 | Voice “open command palette” opens the dialog directly (dead registry id removed). |
| E5 | Generator uses a correct provider check, falls back to the built-in template on AI failure (with a clear note), and styles every section Heading1 with the paragraph split fixed. |

## Remaining (documented, not yet fixed — lower risk, follow-ups)

- A5/A16: selection-delete paragraph merging + post-delete caret anchoring edge cases.
- A18: undo stack memory (per-keystroke deep clone) — needs burst coalescing.
- A20: Word-style wildcard translation for Find.
- B2/E8/E12: real pagination instead of words/320 estimates (two different constants in use).
- B3: chart types beyond column, real equation rendering, SmartArt layouts.
- D2/D3/D4: close-tab unsaved prompt, autosave semantics, beforeunload while saving.
- E6: “Match All Word Forms” checkbox is inert.
- E7: Read/Web layout buttons are placeholders.
- E9–E12: FileMenu recents list, context-menu emoji icons, NavigationPane history/attachments placeholders.
- F1: client-side auth stores plaintext passwords (needs a real backend).
- F3: feature flags are stored but not enforced at runtime.
