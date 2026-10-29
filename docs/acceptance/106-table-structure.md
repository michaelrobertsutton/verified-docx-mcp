# #106 — row operations and faithful table copies

Implementation branch: `codex/106-table-structure`.

Automated: `tests/unit/test_table_structure.py` covers property-preserving insert/delete,
unchanged neighbors, invalid positions and mismatched text leaving packages unchanged,
faithful table properties, and tracked row markers. The full unit gate passed
1,144 tests and 297 subtests on 2026-10-29 before the additional live formatting checks.

Real Word acceptance: PENDING. Use `/private/tmp/verified-docx-106/rows.docx`,
an independent scratch fixture with 8.5pt runs, Compact paragraph properties,
direct borders, fixed layout and non-splitting rows. The isolated pane listens on
53145/53146; the regular production bridge is unchanged.

Before landing, demonstrate insertion/deletion without closing Word; compare
canonical direct table/row/cell/paragraph/run formatting with untouched rows;
exercise faithful copies and explicit overrides; confirm guard refusals leave
content unchanged and every successful/uncertain write is recorded in the audit.

The native row-insertion path currently refuses a nonadjacent template when its
row/cell layout differs from the positional row. This limitation must be resolved
or explicitly reconciled with the issue before acceptance and merge.
