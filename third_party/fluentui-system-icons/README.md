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

`RibbonIcon.Default` is Fluent `apps`: commands without a sufficiently clear icon
keep their full text alongside it. Explicit icons are assigned at each command
call site, independently of translated labels. All commands retain tooltips and
accessible names. Color dropdowns additionally retain their live color swatches.
