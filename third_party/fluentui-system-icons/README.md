# Selected Fluent UI System Icons

Copyright (c) 2020 Microsoft Corporation. Distributed under the [MIT license](LICENSE).

These are unmodified, monochrome **20px Regular** SVGs selected from
https://github.com/microsoft/fluentui-system-icons at revision
`64e5aa757d1b70c24a283256ddcb8269d0f58773` (local source: `D:\Projects\fluentui-system-icons`).
`manifest.json` records each original path and SHA-256. No external checkout is
needed to build or run QuickExcel.

Run `python scripts/generate-ribbon-icons.py` from the QuickExcel root to regenerate
`QuickExcel.Desktop.Avalonia/Views/RibbonIconData.g.cs`, or add `--check` to verify
the source hashes and generated result. The converter rejects unsupported SVG
features and preserves the SVG fill rule. Runtime geometry keeps the complete
20x20 viewBox and inherits the button foreground; no SVG renderer, font, or raster
asset is required. The license is also published under
`Licenses/FLUENT-SYSTEM-ICONS-LICENSE.txt`.

The selection contains 94 icons, including 54 additional command icons for the
compact Ribbon. Explicit icons are required at each command call site,
independently of translated labels; Ribbon commands no longer use the generic
`apps` fallback. Home uses icon-only commands. Other pages display icons and
labels when their full contents fit; they hide command labels before resorting
to horizontal scrolling. Labels return when the page has enough space again.
All commands retain tooltips and accessible names. Font, size and number-format selectors and the style gallery
keep their current values/previews. Color and border controls keep live previews.
