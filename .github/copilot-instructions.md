# Project conventions

## Parsing and behavior

- Reuse the existing push-based worksheet readers and handlers. Avoid parallel parsing paths, per-record objects, and unnecessary allocations during ordinary cell reading.
- Consider XLS, XLSX, Strict XLSX, XLSB, SpreadsheetML 2003 XML, and CSV when changing shared behavior. Keep format-specific limitations explicit.
- Preserve both normal and `SinglePassMode` behavior, including row ordering, gaps, field counts, reset, and navigation.
- Keep optional metadata separate from ordinary cell values. Do not implicitly change `GetValue()` or `AsDataSet()` output when adding metadata accessors.

## Tests and fixtures

- Put shared spreadsheet behavior in `src\ExcelDataReader.Tests\ExcelTestBase.cs` and OpenXML-only shared behavior in `ExcelOpenXmlReaderBase.cs`. Keep format-specific and parser regressions in focused suites.
- Prefer real producer-saved workbooks for compatibility coverage. Use synthetic records for malformed input and precise edge cases.
- Test applicable reading modes and formats. Include Strict XLSX when changing shared OpenXML behavior.

## Build and validation

- The solution is `ExcelDataReader.slnx`. CI restores with `dotnet restore --locked-mode` and builds with `dotnet build ExcelDataReader.slnx --no-restore -c Release /p:TreatWarningsAsErrors=true`.
- Run tests with `dotnet test src\ExcelDataReader.Tests\ExcelDataReader.Tests.csproj -c Release`. Use a focused `--filter` for targeted changes; Windows runs both test target frameworks.
- Never overlap build or test commands that share project outputs.
- Follow `.editorconfig` and keep formatting changes scoped to the task.
- Preserve the library's target-framework compatibility. Avoid unnecessary dependencies and SDK-version-dependent package lock changes; see the explanation in `Directory.Build.props`.

## Public documentation

- Add public reader APIs to the existing README API overview table and relevant usage section.
- Keep README and XML documentation focused on caller-visible behavior, support boundaries, and meaningful costs or safety caveats. Avoid unnecessary implementation details such as cache lifetime, object sharing, and parser mechanics.
