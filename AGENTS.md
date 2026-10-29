# AGENTS.md

This file provides guidance to coding agents working with code in this repository.

## Project

`vaadin-gridexport` is a Vaadin 25 add-on that exports the contents of a Vaadin `Grid` to an Excel workbook (Apache POI). It is a small library published to Maven Central, not an application.

## Build & test commands

Run from the repository root (no Maven wrapper; use system `mvn`):

```bash
mvn clean install                 # full build (root POM + gridexport module)
mvn -pl gridexport compile        # compile only the add-on module
mvn test                          # tests — note: src/test/java currently contains no tests
mvn -pl gridexport test -Dtest=ClassName   # run a single test once tests exist
mvn -Prelease deploy              # release: GPG-sign, attach sources/javadoc, stage to OSSRH
```

- Java 25 LTS (`maven.compiler.release=25`); Vaadin 25 requires Java 21+, but this project targets 25. The compiler runs with `-Xlint:all` (all warnings surfaced, `path` excluded).
- Versions live in the root [pom.xml](pom.xml): Vaadin 25.3.4 (`vaadin-bom`), POI 6.0.0.
- All library dependencies (POI, `flow-server`, `flow-data`, `vaadin-grid-flow`) are `provided` — consumers supply them; do not change to compile scope.

## Architecture

Data flow: **Vaadin `Grid` → `GridHolder` → `ExcelExport` → `InputStream`** consumed by a `FileDownloader`/`StreamResource`.

- [GridHolder.java](gridexport/src/main/java/com/tfyre/vaadin/gridexport/GridHolder.java) — abstraction over grid data: column keys, header labels, `ValueType`, value extraction, alignment. Implement this to export from non-standard data sources.
- [DefaultGridHolder.java](gridexport/src/main/java/com/tfyre/vaadin/gridexport/DefaultGridHolder.java) — default implementation. **It does not discover columns from the Grid**; every exported column must be registered explicitly via `addColumn(key, gridColumn, valueProvider, headerName, valueType)`. `getItems()` materializes the full data set via `dataProvider.fetch(buildQuery(0, Integer.MAX_VALUE))`.
- [GridExport.java](gridexport/src/main/java/com/tfyre/vaadin/gridexport/GridExport.java) — converter contract (`convertGrid()`, `getInputStream()`).
- [ExcelExport.java](gridexport/src/main/java/com/tfyre/vaadin/gridexport/ExcelExport.java) — the only implementation. Template-method pipeline in `convertGrid()`: sheet setup → optional title row → header row → data rows → optional totals row → final formatting (formula evaluation + column autosize).
- [ValueType.java](gridexport/src/main/java/com/tfyre/vaadin/gridexport/ValueType.java) — `DATETIME`, `INTEGERBASE`, `FLOATBASE`, `STRING`; the `numeric` flag drives right-alignment, numeric cell conversion (parsed as `double`), and default Excel formats.

Typical consumer wiring (from [README.md](README.md)):

```java
new FileDownloader(new StreamResource(() -> {
    ExcelExport<BEAN> ee = new ExcelExport<>(new DefaultGridHolder<>(grid));
    ee.setReportTitle("Report Title");
    return ee.getInputStream();
}, "Report Title.xls")).extend(button);
```

### Key behaviors and gotchas

- `ExcelExport` builds an `HSSFWorkbook` by default → legacy `.xls` (65,536-row limit, default filename `Grid-Export.xls`). The longest constructor accepts any `Workbook` (e.g. `XSSFWorkbook`) for `.xlsx`.
- `getInputStream()` calls `convertGrid()` itself. Do not call `convertGrid()` first — the pipeline is not idempotent and would append a second copy of the rows.
- Totals row (enabled by default in the convenience constructors) writes `SUM(...)` formulas for numeric columns and the literal `"Total"` in the first column.
- Per-column formats: `setExcelFormatOfProperty(columnKey, "#,##0.00")`; global defaults via `setDoubleDataFormat` / `setIntegerDataFormat` / `setDateDataFormat`. Cell styles are cached by data-format index in `dataFormatCellStylesMap`.
- Customization is intended through subclassing: all pipeline steps (`initialSheetSetup`, `addTitleRow`, `addHeaderRow`, `addDataRows`/`addDataRow`/`setupCell`, `getCellStyle`, `addTotalsRow`/`setupTotalCell`, `finalSheetFormat`) and the `default*CellStyle` factories are `protected` hooks.
- `setExcelFormatOfProperty` stores `propertyId.toString()` as the map key, so column keys must be strings for lookups in `getCellStyle` to match.

## Git & GitHub conventions

- The default branch is `master`.
- Use the `git` CLI for all local git work (status, diff, log, fetch, commit, branch, merge).
- For GitHub operations — pull requests, issues, repository/release lookups — use the MCP GitHub tools, **not** the `gh` CLI.
- Commit messages must follow [Conventional Commits](https://www.conventionalcommits.org/en/v1.0.0/): `<type>(<scope>): <subject>` in the imperative mood. Examples: `feat(export): allow a custom Workbook`, `fix: handle null cell values`, `docs: add usage notes`, `chore(deps): bump Vaadin to 25.3.4`. Common types: `feat`, `fix`, `docs`, `style`, `refactor`, `perf`, `test`, `build`, `ci`, `chore`, `revert`.

## Code style

Existing code uses `final` on parameters and locals, modern Java features (records, switch expressions), and Javadoc with `@author` headers — match this in new code. `grid.getDataCommunicator()` is used for item fetching, so exports assume the grid is attached and backed by a backend `DataProvider`.
