# Shared grid source provenance

Extracted from QuickExcel.Grid.Wpf in migration A1 (2026-11-03). `SeriesSizes`,
`FastGridCellAddress` and `ActiveSeries` originate from FastWpfGrid commit
`8bcfaf37df343e0b7ccffb3718d9a973a1f58b40`. Their original hashes remain in
`UPSTREAM-SOURCES.json`, and the unchanged MIT notice ships with this assembly.
`GridCellRange` and range/merged-selection algorithms are QuickExcel additions.

The `FastWpfGrid` namespace is retained for source compatibility; this assembly
targets net10.0, has no UI or native dependencies, and does not reference WPF.
The files are moved, not linked or duplicated. Existing hidden/frozen mapping,
sparse sizes, oversized-cell alignment and bounded range enumeration are unchanged.
See [historical modifications](../QuickExcel.Grid.Wpf/SOURCES.md).
