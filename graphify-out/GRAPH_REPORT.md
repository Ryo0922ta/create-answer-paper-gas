# Graph Report - create-answer-paper-gas  (2026-11-22)

## Corpus Check
- cluster-only mode — file stats not available

## Summary
- 349 nodes · 554 edges · 20 communities (16 shown, 4 thin omitted)
- Extraction: 98% EXTRACTED · 2% INFERRED · 0% AMBIGUOUS · INFERRED: 9 edges (avg confidence: 0.85)
- Token cost: 0 input · 0 output

## Graph Freshness
- Built from commit: `f9f186e1`
- Run `git rev-parse HEAD` and compare to check if the graph is stale.
- Run `graphify update .` after code changes (no API cost).

## Community Hubs (Navigation)
- Community 0
- Community 1
- Community 2
- Community 3
- Community 4
- Community 5
- Community 6
- Community 7
- Community 8
- Community 9
- Community 10
- Community 11
- Community 12
- Community 13
- Community 14
- Community 15
- Community 16
- Community 17
- Community 18

## God Nodes (most connected - your core abstractions)
1. `updateOutputParts_()` - 24 edges
2. `outputPlan_()` - 15 edges
3. `FakeSheet` - 13 edges
4. `FakeRange` - 12 edges
5. `readDraft_()` - 12 edges
6. `outputFingerprint_()` - 10 edges
7. `outputOverlap_()` - 10 edges
8. `outputA1_()` - 9 edges
9. `outputConfirmedAdditions_()` - 9 edges
10. `outputMergeRect_()` - 9 edges

## Surprising Connections (you probably didn't know these)
- `outputRect_()` --indirect_call--> `addBasicPartLogged_()`  [INFERRED]
  src/OutputUpdate.js → src/DraftState.js
- `削除ダイアログと削除状況の確認` --conceptually_related_to--> `配置済みパーツの削除確認ダイアログ`  [INFERRED]
  src/PartEditor.html → src/DeletePartDialog.html

## Import Cycles
- None detected.

## Communities (20 total, 4 thin omitted)

### Community 0 - "Community 0"
Cohesion: 0.03
Nodes (55): altered, apiGrid, appended, assert, baseName, baseScore, blank, book (+47 more)

### Community 1 - "Community 1"
Cohesion: 0.07
Nodes (46): answerPattern_(), answerRowLayout_(), calculatePlacements_(), cancelDeletePart(), collectAnswerLabels_(), configureBasicSheet_(), configureOutputSheet_(), confirmDeletePart() (+38 more)

### Community 2 - "Community 2"
Cohesion: 0.14
Nodes (44): outputA1_(), outputBasicBottomRightRect_(), outputCache_(), outputCachedGrid_(), outputCanonical_(), outputCellBlank_(), outputCellPosition_(), outputConfirmedAdditions_() (+36 more)

### Community 3 - "Community 3"
Cohesion: 0.06
Nodes (15): assert, basic, book, columnNumber(), context, FakeBook, FakeRange, FakeSheet (+7 more)

### Community 4 - "Community 4"
Cohesion: 0.26
Nodes (19): addBasicPart(), addBasicPartLogged_(), applyDraftPartEdit_(), basicPartFromTemplate_(), commitDraft_(), deleteDraftPart_(), DRAFT, draftSheet_() (+11 more)

### Community 5 - "Community 5"
Cohesion: 0.13
Nodes (13): additional, assert, compactGroup(), context, differences, expected, fs, normalizeLabel() (+5 more)

### Community 6 - "Community 6"
Cohesion: 0.13
Nodes (14): assert, candidates, changed, context, drawn, edited, essay, fs (+6 more)

### Community 7 - "Community 7"
Cohesion: 0.13
Nodes (13): assert, basic, choiceStart, ctx, elements, end, form, fs (+5 more)

### Community 8 - "Community 8"
Cohesion: 0.23
Nodes (11): analyzeDocumentText(), analyzeDocumentTextLogged_(), analyzeLeaf_(), analyzeMajorBlock_(), analyzeSubgroup_(), DOCUMENT_ANALYSIS, expandAnswerRanges_(), importAnalyzedDocument() (+3 more)

### Community 9 - "Community 9"
Cohesion: 0.15
Nodes (9): assert, ctx, fs, max, path, {performance}, report, samples (+1 more)

### Community 10 - "Community 10"
Cohesion: 0.18
Nodes (9): alerts, assert, context, draft, fs, path, sheets, templateRows (+1 more)

### Community 11 - "Community 11"
Cohesion: 0.18
Nodes (8): assert, candidate, context, fs, logs, path, root, vm

### Community 12 - "Community 12"
Cohesion: 0.29
Nodes (7): パーツエディタ, 大問・小問・解答欄の判定ルール, 問題文から大問と小問を解析, 基本パーツの出力追加, 変更済みパーツの登録解除, 解答欄プレビュー, 配置済みパーツの編集と複製

### Community 13 - "Community 13"
Cohesion: 0.33
Nodes (5): dependencies, enabledAdvancedServices, exceptionLogging, runtimeVersion, timeZone

### Community 14 - "Community 14"
Cohesion: 0.70
Nodes (4): installActionButtons(), onOpen(), onOpenLogged_(), removeLegacySheetButtons_()

### Community 16 - "Community 16"
Cohesion: 0.67
Nodes (3): 削除操作識別子, 配置済みパーツの削除確認ダイアログ, 削除ダイアログと削除状況の確認

## Knowledge Gaps
- **10 isolated node(s):** `大問・小問・解答欄の判定ルール`, `基本パーツの出力追加`, `変更済みパーツの登録解除`, `解答欄プレビュー`, `削除操作識別子` (+5 more)
  These have ≤1 connection - possible missing edges. (Counts symbols only; 188 node(s) total have ≤1 connection when file, concept and rationale nodes are included.)
- **4 thin communities (<3 nodes) omitted from report** — run `graphify query` to explore isolated nodes.

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **Why does `outputRect_()` connect `Community 2` to `Community 4`?**
  _High betweenness centrality (0.015) - this node is a cross-community bridge._
- **What connects `大問・小問・解答欄の判定ルール`, `基本パーツの出力追加`, `変更済みパーツの登録解除` to the rest of the system?**
  _10 weakly-connected nodes found - possible missing edges._
- **Should `Community 0` be split into smaller, more focused modules?**
  _Cohesion score 0.03225806451612903 - nodes in this community are weakly interconnected._
- **Should `Community 1` be split into smaller, more focused modules?**
  _Cohesion score 0.06676342525399129 - nodes in this community are weakly interconnected._
- **Should `Community 2` be split into smaller, more focused modules?**
  _Cohesion score 0.1396854764107308 - nodes in this community are weakly interconnected._
- **Should `Community 3` be split into smaller, more focused modules?**
  _Cohesion score 0.05731707317073171 - nodes in this community are weakly interconnected._
- **Should `Community 5` be split into smaller, more focused modules?**
  _Cohesion score 0.13333333333333333 - nodes in this community are weakly interconnected._