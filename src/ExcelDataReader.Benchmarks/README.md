# Allocation benchmarks

`ReadTextFragments` covers single-fragment and rich XLSX strings, SpreadsheetML
Data text, and OOXML escapes. `ParseCustomFormats` covers custom number-format
parsing, including simple decimals, conditions, dates, exponentials, and fractions.
Both use `MemoryDiagnoser`; workbook generation happens in `GlobalSetup`,
outside the measured operation. Existing `ReadAllStrings`, `ReadRealWorldFiles`,
and `SinglePassRead` benchmarks provide whole-file comparisons.

`SharedStringStorage` compares Default, below-threshold SpillToDisk, and disk-spilled SST
storage using generated million-entry workbooks. See [shared string storage
details and benchmarks](SharedStringStorage.md) for measured memory/throughput
tradeoffs, corpus generation, and reproduction commands.

`SharedStringConstruction` measures reader creation/disposal without traversal
for the workbook selected by `EDR_SST_INPUT`. Optional `EDR_SST_LIBRARY` selects
a preserved baseline library DLL. See the storage document for reproduction
commands and storage tradeoffs.

`SharedStringTraversal` compares Default, SpillToDisk1 (1 MiB), and SpillToDisk256 (256 MiB) over the existing workbook
selected by `EDR_SST_INPUT`, including reader creation, complete traversal,
checksum validation and disposal with SinglePassMode enabled. It is intended
to expose decoding/repeated-reference costs that construction-only timings miss.
For the supplied 100k/one-million-entry corpora, SpillToDisk1 forces spill and SpillToDisk256
stays below threshold. Confirm spill with `--sst-memory`; its forced-GC
checkpoints measure retention, not throughput.

`--sst-io` diagnoses injected MemoryStream/FileStream construction and lookup
calls over sequential, reverse-like, random, and hot Unicode references.
It runs three warmups and five measured repetitions; counts are managed-stream
calls, not OS I/O. Timing includes output validation and is diagnostic rather
than a replacement for whole-read BenchmarkDotNet results.

`ParseLongFormats` covers long digit/date patterns, repeated bracket directives,
and quoted Unicode text. The format parser stores temporary source ranges rather
than token strings and only produces validity/date/duration classification.
Modern targets validate numeric slices with spans; older targets retain numeric
string fallbacks where the corresponding span APIs are unavailable.

Run the allocation benchmarks on net10.0:

```powershell
dotnet run --project src\ExcelDataReader.Benchmarks\ExcelDataReader.Benchmarks.csproj -c Release -f net10.0 -- --filter "*ReadTextFragments*" "*ParseCustomFormats*" "*ParseLongFormats*" --exporters json
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
