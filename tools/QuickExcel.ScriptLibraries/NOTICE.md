# Built-in JavaScript libraries — library set 1

QuickExcel 7.0 bundles the following fixed upstream releases. Download locations
and npm integrity hashes, including transitive dependencies, are recorded in
`package-lock.json`. Downloads, npm caches and generated bundles stay under
`artifacts/script-libraries`; no npm lifecycle scripts are run.

| Import ID | Upstream package/version | License |
| --- | --- | --- |
| lodash | lodash 4.18.1 | MIT |
| decimal.js | decimal.js 10.6.0 | MIT |
| dayjs | dayjs 1.11.23 | MIT |
| papaparse | papaparse 5.7.1 | MIT |
| simple-statistics | simple-statistics 8.0.0 | ISC |
| zod | zod 4.6.7 | MIT |
| fuse.js | fuse.js 7.5.0 | Apache-2.0 |
| arquero | arquero 8.0.3 | BSD-3-Clause |
| mathjs | mathjs 15.2.0 | Apache-2.0 |

Sources: https://github.com/lodash/lodash,
https://github.com/MikeMcl/decimal.js, https://github.com/iamkun/dayjs,
https://github.com/mholt/PapaParse, https://github.com/simple-statistics/simple-statistics,
https://github.com/colinhacks/zod, https://github.com/krisk/Fuse,
https://github.com/uwdata/arquero, https://github.com/josdejong/mathjs.

`core-js-pure` 3.51.0 (MIT; https://github.com/zloirock/core-js) supplies a
data-only URL implementation. File is a QuickExcel-owned in-memory Blob subclass;
it does not access the operating-system filesystem. Lodash-es 4.18.1 supplies the
upstream export-name list, while execution uses the full Lodash build, including
chaining. Type declarations use the packages' own declarations and
@types/lodash 4.17.27 and @types/papaparse 5.5.2 (MIT). @types/lodash-es 4.17.12
is a pinned build input. Full package and declaration licenses are distributed in
`THIRD-PARTY-SCRIPT-LIBRARIES.txt`; original copyright notices remain in bundles.

## QuickExcel adaptations

- esbuild 0.28.2 produces self-contained ES2022 ESM modules. Browser mappings
  supplied by Papa Parse, Arquero and math.js dependencies avoid Node-only code.
  Remaining external imports fail the build; code is not downloaded at runtime.
- Lodash receives named ES exports in addition to its callable default export.
- Day.js preinstalls UTC, Timezone, CustomParseFormat, Duration, IsBetween,
  IsSameOrBefore, IsSameOrAfter, IsoWeek, WeekOfYear, QuarterOfYear, AdvancedFormat,
  LocalizedFormat, LocaleData, MinMax and RelativeTime; English remains the
  default locale and Chinese (zh-cn) is bundled.
- Papa Parse accepts strings; direct file, download and browser-worker modes
  are explicitly rejected. Use `workbook.files` or permission-gated Action
  fetch, then parse the text. Upstream declaration overloads for unavailable
  file/stream/worker paths are removed; other declaration bodies retain attribution.
- Two upstream declaration defects are repaired only in generated copies:
  Arquero's optional rest parameter and Flechette's untyped `getFactory` member.
  No `skipLibCheck` or blanket `any` module declaration hides compiler errors.
  Internal declaration imports use relative bundle paths so same-named user
  modules cannot redirect library-to-library types in the editor.
- Arquero's four ordinary remote loaders use the existing permission-gated fetch.
  Its four stream readers collect bounded text/UTF-8/Arrow bytes (10 MiB maximum)
  without browser stream constructors. They buffer input rather than providing
  incremental processing. Compressed files still require an unavailable codec.
  CSV/JSON/fixed-width stream declarations also accept byte readers.
- Lodash timer helpers remain unavailable because execution has no background
  timer scheduler. The worker keeps its OS, memory, stack, native watchdog and
  host transaction limits.

Runtime bundle and declaration hashes form a catalog identity included in the
worker handshake, compiler identity/cache and execution-consent policy. Existing
workbook modules with the same ID take precedence, preserving their behavior.
The current product ships one fixed set; updating it requires a deliberate
product release and renewed execution-policy acceptance, not a background npm
update. A source signature alone does not pin a historical runtime. Source-only
XLSQ files do not carry npm binaries or installation paths.

Rebuild: `node tools/QuickExcel.ScriptLibraries/build.mjs`; missing pinned
TypeScript sources are prepared automatically. Managed builds prepare resources through
QuickExcel.ScriptContracts, and Web builds reuse the same declarations. Generated
resources are embedded in the worker/host assemblies and the offline editor;
Node/npm are build-time tools, not installed-application dependencies.
