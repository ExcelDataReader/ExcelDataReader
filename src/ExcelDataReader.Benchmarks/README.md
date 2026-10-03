# Allocation benchmarks

`ReadTextFragments` covers single-fragment and rich XLSX strings, SpreadsheetML
Data text, and OOXML escapes. `ParseCustomFormats` covers custom number-format
parsing. Both use `MemoryDiagnoser`; workbook generation happens in `GlobalSetup`,
outside the measured operation. Existing `ReadAllStrings`, `ReadRealWorldFiles`,
and `SinglePassRead` benchmarks provide whole-file comparisons.

Run the allocation benchmarks on net10.0:

```powershell
dotnet run --project src\ExcelDataReader.Benchmarks\ExcelDataReader.Benchmarks.csproj -c Release -f net10.0 -- --filter "*ReadTextFragments*" "*ParseCustomFormats*" --exporters json
```

The net10.0 host uses the library's net8.0 build through normal project-reference
selection. The library continues to support net462, netstandard2.0,
netstandard2.1, and net8.0, with optimized paths gated by API availability.

Run builds, tests, and benchmarks serially: they share output directories. Capture
unchanged-code and candidate results in separate artifact directories on the same
machine with the same runtime and library target. Use warmed, repeated
measurements, not a `Dry` job, to assess throughput. Keep an allocation optimization
only when it reduces allocations without a statistically meaningful slowdown,
including multipart text and ordinary strings.
