# Third-party notice: Office Scripts API definitions

QuickExcel vendors the TypeScript declaration file that Microsoft uses as input
for the Office Scripts API reference.

- Repository: <https://github.com/OfficeDev/office-scripts-docs-reference>
- Path: `generate-docs/script-inputs/office-scripts-docs.d.ts`
- Commit: `97f8f415842e13e004d7c53c1d24fe45653e50c0`
- Size: 487,506 bytes
- Git blob SHA-1: `a2ea76f0e65faab8bd70e2fcd683c9e00502b6bd`
- SHA-256: recorded in `source.lock.json` after the first verified fetch
- License: CC BY 4.0, <https://creativecommons.org/licenses/by/4.0/legalcode>
- Copyright: Microsoft Corporation and contributors

## Modifications

The declarations generated for QuickExcel differ from the source file:

- the `ExcelScript` namespace is renamed to `QuickExcel`;
- documentation comments are removed;
- the `OfficeScript` namespace is removed;
- QuickExcel host extension declarations are appended.

## Scope

Only type declarations are used. QuickExcel implements the described behavior
independently; no Microsoft runtime code is included, linked, or executed.

## Attribution

Any distribution that includes the vendored source, the generated declaration,
or a work derived from them must retain this notice, the CC BY 4.0 reference,
and the statement that the material was modified.
