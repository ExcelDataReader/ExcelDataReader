# Shared string storage: details and benchmarks

See the [reader configuration example](../../README.md#large-shared-string-tables).

## Storage behavior

XLSX, XLSB, and BIFF8 XLS use shared string tables (SSTs). `Default` eagerly
stores decoded strings in the normal in-memory list. Opt-in `SpillToDisk` uses
the same representation until the threshold is exceeded. Before spill,
SpillToDisk lookups reuse stored values without an additional decoded cache.
Declared XML SST capacity is reserved only when it fits within the remaining
budget.

Both binary formats synchronously send completed strings through the same
borrowed UTF-16 sink. The in-memory table decodes and owns each string during
the call; the spill store consumes the buffer without retaining it and writes
valid UTF-16 code units to disk after spill. XLS uses its own
continuation-aware parser and reuses one assembly buffer; XLSB reads standalone
string records directly from each record buffer. Their parsing state machines
remain separate. Eager Default XLS decodes the complete SST during workbook
opening, so short-circuit reads still pay that decoding cost. XLSB trailing
rich-text and phonetic data are ignored. A character count that does not fit
within an XLSB record raises the standard BIFF string-size error.

### Eager XLS policy comparison (2026-10-07)

The selected implementation was compared with the previous lazy XLS path using
100,000 unique 128-character strings. Times are full open/read/dispose totals;
allocations cover opening plus traversal.

| Runtime / workload | Lazy ms | Eager ms | Time change | Allocation lazy → eager |
| --- | ---: | ---: | ---: | ---: |
| .NET 10 sequential full | 106.0 | 90.7 | -14.5% | 81.94 → 50.57 MiB |
| .NET 10 hot full | 138.2 | 105.8 | -23.5% | 88.04 → 56.67 MiB |
| .NET 10 random full | 111.5 | 90.7 | -18.6% | 81.94 → 50.57 MiB |
| .NET 10 sequential 1% | 44.0 | 60.0 | +36.4% | 49.46 → 44.53 MiB |
| .NET Framework sequential full | 440.2 | 402.4 | -8.6% | 82.84 → 51.37 MiB |
| .NET Framework hot full | 614.2 | 586.0 | -4.6% | 88.96 → 57.49 MiB |
| .NET Framework random full | 389.0 | 337.0 | -13.4% | 82.84 → 51.37 MiB |
| .NET Framework sequential 1% | 55.9 | 103.3 | +84.8% | 49.50 → 45.31 MiB |

The full-read improvement and partial-read regression were explicit policy
tradeoffs, not a workload-independent performance claim. Measurements used
paired, reversed-order processes, checksum/count/identity validation, and
workstation GC on Windows; they are synthetic and do not establish a real-world
workload mix.

Capacity growth, strings and lookup cache are accounted. When an add exceeds
the budget, SpillToDisk migrates one entry at a time to disk, then releases the
normal table. It does not create a second full decoded table during migration.
The workbook owns and seals the shared store after parsing; reset and sheet
changes retain it. XML whitespace, rich text, escapes, and binary surrogate
replacement retain Default behavior.

### Disk representation and cache

After spill, text is stored uniformly as exact UTF-16 little-endian code units.
The disk index uses **16 bytes per entry**, containing a 64-bit byte offset and
32-bit character count with four padding bytes. No resident per-entry disk index
is retained. After flushing all writes, the store creates unnamed, read-only
file mappings. A lazily created view per file is bounded to 64 MiB on 64-bit
processes or 8 MiB on 32-bit processes; the entire file need not fit in a
contiguous address range. Safe bulk reads copy exact code units into the
existing buffers without acquiring or retaining unmanaged pointers.

Lookup reuses the 32,768-byte serialization buffer as a payload read window
and uses a 4,096-byte index window. Adjacent uncached references in either
direction use buffered aligned read-ahead. Unrelated misses use mapped reads
when covered by the current view, otherwise exact-sized buffered reads.
They do not continually remap large files. This is an explicit access policy,
not a fallback after a mapping error. Strings can cross buffer and view
boundaries, and offsets remain 64-bit.

Mappings and windows are initialized only after sealing. File lengths are
checked before mapping; index offsets, alignment
and lengths are checked before decoding. Short buffered reads are completed
explicitly; failed fills never expose an unread tail. The additional index
capacity is 4,080 bytes, and mapping/view bookkeeping reserves a conservative
2,048 accounted bytes per nonempty mapped file. The last-value/second-touch
decoded cache is independently bounded.

The cache budget is min(one eighth of the configured limit, 8 MiB), including
metadata, with at most 4,096 slots. One-off scans do not admit entries.
Oversized values are returned without retention. Reference identity is not
guaranteed after spill.

### Budget semantics

`SharedStringSpillThreshold` is measured in bytes, defaults to 64 MiB, and must be
at least 1 MiB in SpillToDisk mode. Default ignores the threshold and temporary directory.

The SpillToDisk budget accounts for allocated reference capacity (including
unused slots), conservative object/array/string overhead, and retained decoded
strings. Empty strings share the runtime singleton. Disk buffers and cache are
allocated only after spill.

This is **not a total managed-heap or process-memory limit**. Parser temporaries,
a single decoded string, current rows, input buffering, decryption, workbook
metadata, OS file caches, mapped file pages and values retained by callers are
outside it. Mapping reserves virtual address space, not a resident managed
byte array; touched pages can increase the process working set independently
of the configured budget. The two views together reserve at most 128 MiB on
64-bit or 16 MiB on 32-bit processes, excluding OS alignment/page rounding.
Migration can transiently exceed the threshold while serializing an entry.

`AsDataSet()` retains its own data. `SinglePassMode` skips worksheet preparation,
but not SST parsing. SST options do not affect CSV, SpreadsheetML, or older XLS
formats without SSTs.

### Temporary files and failures

Files are created only when spill occurs, in the configured existing directory.
`SharedStringTemporaryDirectory = null` uses the system temporary directory.
The directory needs read/write permissions and sufficient free space.

Payload and index files use exclusive delete-on-close handles and bounded
buffers. `Close()`/`Dispose()` releases views and mappings before their backing
streams, independently of `LeaveOpen`. File and mapping creation, view creation,
write, flush, seek and lookup failures propagate; there is no silent fallback
to unbounded RAM or suppression of mapping errors. Spill requires platform
support for persisted memory-mapped files. Windows .NET 10 and .NET Framework
were tested; API availability on the library targets is not a claim of
validation on every operating system.

Temporary strings are **plaintext even for password-protected workbooks**.
Secure deletion and cleanup after abrupt process termination are not guaranteed.

## Benchmark corpus

Large worksheets do not necessarily have large SSTs: `10x10000.xlsx` has only
20 unique strings and a 558-byte SST. The benchmark project includes a streaming
generator with seed 741, unique index prefixes and varied suffixes. It validates
cell counts, character counts and a stable content checksum.

| Profile | Unique entries | Characters per entry | Purpose |
|---|---:|---:|---|
| Control | 100,000 | 64 | Below the normal 64 MiB threshold |
| Large SST | 1,000,000 | 64 | Compare in-memory storage and actual spill |
| Long strings | 1,000,000 | 256 | Make retained text dominate |

The generator supports XLSX, XLSB and raw BIFF8 XLS. Reference patterns are
`Sequential`, `Permuted`, `Random`, `Hot` and `Sparse`; an optional boolean
selects non-Latin-1 text. Generation is outside measured operations, and all
modes read the same logical corpus. Large generated files are not committed.

## Measurements (2026-10-05)

Windows 11, Core i5-13400F, SDK 10.0.401 / .NET 10.0.12, x64, local C: storage,
Workstation GC. Paired Stopwatch hosts referenced preserved library binaries:
the buffered baseline at commit `325098c` and the current mapped/buffered policy.
Each operation opened the reader with SinglePassMode, completely traversed it,
validated cell/character counts and a stable checksum, and disposed it.
The generated files were identical for both versions.

There were three warmups and eight measured iterations per process, two
processes per version/workload, with version order reversed on the second run.
All 16 observations were retained. `DOTNET_TieredCompilation=0` avoided tiering
transitions observed in short processes. These are warmed-cache comparisons,
not cold-storage measurements or BenchmarkDotNet summaries. Means and sample
standard deviations are milliseconds; they should not be mixed with timings
from different JIT settings or harnesses.

| Format / pattern / entries | Current Default ms | Buffered baseline SpillToDisk1 ms | Current SpillToDisk1 ms |
|---|---:|---:|---:|
| XLSX / Sequential / 100,000 | 199.30 ± 7.66 | 289.69 ± 13.44 | 272.83 ± 22.84 |
| XLSB / Sequential / 100,000 | 74.83 ± 5.47 | 141.09 ± 10.24 | 145.83 ± 14.94 |
| XLS / Sequential / 100,000 | 102.18 ± 8.40 | 125.18 ± 12.94 | 126.29 ± 15.57 |
| XLSB / Permuted / 100,000 | 72.82 ± 4.30 | 139.75 ± 17.91 | 148.03 ± 16.59 |
| XLSB / Random / 100,000 | 80.82 ± 8.35 | 655.21 ± 38.39 | 165.12 ± 18.16 |
| XLSX / Hot Unicode / 100,000 | 256.98 ± 9.15 | 328.08 ± 25.36 | 325.31 ± 26.39 |
| XLSB / Sequential / 1,000,000 | 745.99 ± 17.08 | 1,486.83 ± 241.47 | 1,414.06 ± 287.46 |
| XLSB / Random / 1,000,000 | 890.28 ± 48.09 | 9,136.56 ± 392.74 | 3,725.89 ± 169.35 |

Random full traversal improved about 4.0x at 100,000 entries and 2.5x at one
million. The million-entry payload exceeds a view; out-of-view random misses
therefore still perform buffered I/O. Adjacent and hot workloads do not show
a consistent full-operation gain; opening/spill costs remain substantial.
Sequential controls retain checksum `9989736530551659307`, and the million
sequential corpus retains `7742303954537846853`.

For XLSB random, opening averaged 115.19 → 117.13 ms at 100,000 entries and
1,228.21 → 1,270.90 ms at one million. Traversal averaged 538.63 → 45.72 ms
and 7,896.32 → 2,438.21 ms, respectively. Disposal averaged 1.39 → 2.27 ms
and 12.02 → 16.78 ms. These phases demonstrate that the benefit is in lookup,
not removal of parsing, serialization or disposal costs. The XLSB rows above
predate synchronous XLSB string ingestion; see below for its effect.

### XLSB string ingestion (2026-10-06)

A paired Stopwatch host referenced preserved library binaries built from commit
`965037b` and from the current source. Fixtures had 100,000 unique 128-character
entries (ASCII or Cyrillic text) referenced `Sequential`, `Hot` or `Random`;
`Partial` read only the first 16 cells. Profiles were Default, SpillToDisk256
(no spill) and SpillToDisk1 (spill). Each process ran three warmups and ten
measured operations; two timing processes per version used reversed order, and
a third process per version recorded forced-GC retention. Tiered compilation was
disabled with Workstation GC. Every operation validated cell count, checksum,
SST count and, without spill, reference identity. Values are medians.

| Profile | Runtime | Opening allocation MiB | Total time change |
|---|---|---:|---:|
| Default, SpillToDisk256 | .NET 10 | 31.04 → 28.75 | -9% to +5% |
| Default, SpillToDisk256 | .NET Framework | 32.28–34.16 → 29.99–31.86 | -13% to +3% |
| SpillToDisk1 | .NET 10 | 29.14 → 1.10 | -11% to -33% |
| SpillToDisk1 | .NET Framework | 30.39–32.26 → 1.52–3.38 | -10% to -41% |

Without spill, opening allocation falls by 2.29 MiB, the per-string record
objects; retained memory and traversal allocation are unchanged. Most timing
differences without spill are within process-to-process variance. A 40-sample
recheck of the largest .NET 10 increase (Hot, SpillToDisk256) measured
136.4 → 129.3 ms. On .NET Framework, lower opening allocation moves a gen2
collection from opening into traversal, so traversal can be 2–4 ms slower while
total time is still lower.

With spill, opening no longer allocates a decoded string per entry. For
Cyrillic text, opening medians fell from 78–81 to 52–68 ms on .NET 10, and
from 98–115 to 53–66 ms on .NET Framework. Traversal time and allocation,
accounted resident bytes (83,136, or 156,864 with hot-value caching) and disk
bytes (27,200,000) are unchanged.

For 1,000,000 unique 64-character ASCII entries read sequentially on .NET 10,
Default opening allocation fell from 183.90 to 161.01 MiB and total time from
731 to 669 ms. SpillToDisk1 opening allocation fell from 168.06 to 1.11 MiB and
total time from 976 to 868 ms, with unchanged 83,008 accounted and 144,000,000
disk bytes.

### XLS spill ingestion (2026-10-06)

This comparison predates the common eager Default policy documented above.
Default-specific figures below describe the source used for this 2026-10-06
comparison, not the selected eager implementation.

The XLS parser sends each assembled UTF-16 string synchronously to the same
`ISharedStringSink` used by XLSB, for both Default and SpillToDisk. It reuses
one assembly buffer and does not create an `XlsUnicodeString` wrapper per
entry. The XLS continuation/parser state machine remains format-specific.

A Stopwatch host referenced preserved `net8.0` and `net462` library binaries
from commit `ca4e31c` and the candidate. Each corpus had 100,000 unique
128-character strings, either compressed single-byte text or Unicode, with
sequential references. Two processes per binary/workload ran three warmups and
ten measured full open/read/dispose operations; baseline/candidate process
order was reversed. Tiered compilation was disabled on .NET 10. Times below
are medians across 20 operations in ms; allocation is the average opening
allocation in MiB.

| Runtime | Mode / text | Opening allocation baseline → candidate | Open ms baseline → candidate | Full ms baseline → candidate |
|---|---|---:|---:|---:|
| .NET 10 | SpillToDisk256 / narrow | 75.06 → 44.47 | 59.57 → 69.34 | 80.68 → 90.40 |
| .NET 10 | SpillToDisk256 / Unicode | 87.82 → 57.13 | 67.69 → 65.55 | 90.98 → 91.06 |
| .NET 10 | SpillToDisk1 / narrow | 47.41 → 16.81 | 61.93 → 63.95 | 104.20 → 106.17 |
| .NET 10 | SpillToDisk1 / Unicode | 60.17 → 29.47 | 62.30 → 62.97 | 103.54 → 104.59 |
| .NET Framework | SpillToDisk256 / narrow | 75.95 → 45.26 | 110.12 → 101.66 | 337.41 → 326.56 |
| .NET Framework | SpillToDisk256 / Unicode | 88.40 → 57.59 | 127.24 → 107.48 | 350.02 → 330.12 |
| .NET Framework | SpillToDisk1 / narrow | 47.49 → 16.80 | 75.84 → 73.67 | 333.90 → 333.10 |
| .NET Framework | SpillToDisk1 / Unicode | 59.90 → 29.14 | 79.86 → 74.44 | 339.25 → 336.10 |

The parser handoff lowers opening allocation by about 30.6 MiB in all four
text/runtime combinations; for forced spill that is 65% for narrow and 51% for
Unicode input on .NET 10. Forced-spill full-operation times stayed within about 2% of baseline medians.
The no-spill SpillToDisk256 narrow-text profile was about 12% slower on .NET 10,
while the Unicode control was unchanged and both .NET Framework controls were
faster. This is an allocation benefit, not a consistent throughput improvement. Default
retained identical opening allocation (49.13/61.88 MiB on .NET 10 and
49.17/61.59 MiB on .NET Framework) and matching checksums. All modes returned
100,000 strings with the expected checksum; the forced-spill payload and index
totaled 27,200,000 bytes.

A real-file direct-store diagnostic, independently of workbook parsing, used
the same warmup/repetition scheme and JIT setting. At 100,000 entries:

| Lookup pattern | Buffered baseline ms | Current ms |
|---|---:|---:|
| Sequential | 17.37 ± 0.73 | 17.72 ± 1.97 |
| Permuted | 17.03 ± 1.23 | 16.77 ± 0.75 |
| Random | 514.77 ± 16.68 | 39.59 ± 1.29 |

Sealing/flushing plus mapping setup averaged 0.15 ms for the random mapped
case versus 0.02 ms buffered. Mapping setup allocated 496 managed bytes;
first lookup, including lazy view creation, averaged 0.038 versus 0.026 ms.
Lookup allocation itself was unchanged at 15,199,848 bytes in that diagnostic.
Full random traversals allocated 38.5967 → 38.5978 MiB at 100,000 entries
and 383.6361 → 383.6372 MiB at one million.

### Managed memory versus mapped pages

Fresh `--sst-memory` forced-GC checkpoints are **not timing evidence**.
For the million-entry sequential XLSB corpus, SpillToDisk1 had 144,000,000 disk
bytes, 83,008 accounted bytes after traversal, 64,192 retained managed bytes
after opening and 97,384 after traversal. Mapping bookkeeping
adds 4,096 accounted bytes relative to the buffered baseline.

Separate fresh-process random checkpoints recorded active view capacity,
managed retention and aggregate process working set. At 100,000 entries the
two active views covered 14,400,000 bytes, and working set changed from
69.14 MB after traversal to 54.78 MB after disposal. At one million entries,
view capacity was 83,108,864 bytes (64 MiB payload plus 16 MB index), and
working set changed from 137.86 to 54.76 MB. Views were absent after disposal.
Managed retention stayed roughly 0.2 MB after traversal in these cold-host
checkpoints, including runtime/parser metadata and diagnostic overhead.

View capacity is virtual coverage, not a measurement of resident pages.
Working set is process-wide and includes runtime memory, not just mappings;
the observed drops are not a portable formula for mapped RAM. These results
illustrate why the managed SST budget is not a process-memory limit.

The `--sst-io` diagnostic deliberately wraps injected streams. Such wrappers
exercise the buffered path, not the real-file mapped path. Its counts are
managed stream calls, **not OS I/O operations**, and MemoryStream timings
include validation, hashing, allocation and caching rather than isolating
decoding. Use complete real-file traversal to measure mapped access.

All four library targets and three benchmark targets built without warnings.
Full tests passed on .NET 10 and .NET Framework; additional store tests passed
with the Framework host explicitly forced to x86. A sparse-file diagnostic
verified a 32 KiB read crossing a 64 KiB view boundary beyond a 4 GiB offset
on x64 without allocating a multi-gigabyte managed array. Framework throughput,
non-Windows execution and controlled cold-cache performance were not measured.

## Reproduction

Run from the repository root with an existing writable corpus directory.
Explicit generation leaves files for reuse; remove them when finished.

```powershell
dotnet run --project src\ExcelDataReader.Benchmarks\ExcelDataReader.Benchmarks.csproj -c Release -f net10.0 -- --sst-generate C:\bench\control.xlsx 100000 64 Sequential
dotnet run --project src\ExcelDataReader.Benchmarks\ExcelDataReader.Benchmarks.csproj -c Release -f net10.0 --no-build -- --sst-generate C:\bench\million.xlsx 1000000 64 Sequential
dotnet run --project src\ExcelDataReader.Benchmarks\ExcelDataReader.Benchmarks.csproj -c Release -f net10.0 --no-build -- --sst-memory C:\bench\million.xlsx true SpillToDisk64
$env:EDR_SST_INPUT = 'C:\bench\control.xlsx'
dotnet run --project src\ExcelDataReader.Benchmarks\ExcelDataReader.Benchmarks.csproj -c Release -f net10.0 --no-build -- --filter "*SharedStringTraversal*" --job Short --launchCount 1 --warmupCount 3 --iterationCount 5 --artifacts artifacts\sst-auto
dotnet run --project src\ExcelDataReader.Benchmarks\ExcelDataReader.Benchmarks.csproj -c Release -f net10.0 --no-build -- --sst-io
```

The extension selects the format. For Unicode/hot data, append `Hot true` instead
of `Sequential` to the generation command. The memory harness boolean selects
SinglePassMode; profiles include Default, SpillToDisk1, SpillToDisk16, SpillToDisk64 and SpillToDisk256.
Thresholds are in MiB. Confirm actual spill: SpillToDisk256 is not universally below
budget. Use `--filter "*SharedStringConstruction*"` for opening-only timings,
or `--filter "*SharedStringStorage*"` for generated million-entry comparisons.

The net10.0 host selects the library's net8.0 build. Use `-f net462` on Windows
for the Framework host. Optional `EDR_SST_LIBRARY` selects a compatible preserved
library for construction-only baseline comparisons; unset it for the project
reference. It also selects the library for the direct-store `--sst-io` diagnostic,
which runs independently of `EDR_SST_INPUT` and creates delete-on-close files in
the current directory. Run builds, tests and benchmarks serially because output directories
are shared. Forced-GC harness timings are diagnostic, not throughput evidence.
