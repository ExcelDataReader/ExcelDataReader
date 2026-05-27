using System.CommandLine;
using System.Data;
using System.Diagnostics;
using System.Text;
using ExcelDataReader;
using Spectre.Console;

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

var rootCommand = new RootCommand("ExcelDataReader sample console application");
rootCommand.Subcommands.Add(BuildExcelCommand());
rootCommand.Subcommands.Add(BuildCsvCommand());
return rootCommand.Parse(args).Invoke();

// ---- excel subcommand -------------------------------------------------------
static Command BuildExcelCommand()
{
    var fileArg = new Argument<FileInfo>("file") { Description = "Path to the Excel file (XLS, XLSX, XLSB)" };
    var sheetNameOpt = new Option<string[]>("--sheet-name", ["-n"]) { Description = "Filter by sheet name (repeatable; names may contain commas)", AllowMultipleArgumentsPerToken = false };
    var sheetIndexOpt = new Option<string[]>("--sheet-index", ["-i"]) { Description = "Filter by 1-based sheet index, comma-separated or repeatable (e.g. 2,5,7)", AllowMultipleArgumentsPerToken = false };
    var noHeaderOpt = new Option<bool>("--no-header", ["-H"]) { Description = "Don't treat first row as column names" };
    var fillMergedOpt = new Option<bool>("--fill-merged") { Description = "Fill merged cell values across the merged range (implies --dataset)" };
    var singlePassOpt = new Option<bool>("--single-pass") { Description = "Enable single pass mode (skips pre-scan for row/column counts)" };
    var outputOpt = new Option<OutputFormat>("--output", ["-o"]) { Description = "Data output format: table, csv, tsv (default: no data output, stats only)", DefaultValueFactory = _ => OutputFormat.None };
    var passwordOpt = new Option<string?>("--password", ["-p"]) { Description = "Password for protected workbooks" };
    var encodingOpt = new Option<string>("--encoding", ["-e"]) { Description = "Fallback encoding for XLS BIFF2-5 (default: windows-1252)", DefaultValueFactory = _ => "windows-1252" };
    var dataSetOpt = new Option<bool>("--dataset") { Description = "Use AsDataSet extension (loads all data into a DataSet in memory)" };

    var cmd = new Command("excel", "Read XLS, XLSX, or XLSB files");
    cmd.Arguments.Add(fileArg);
    cmd.Options.Add(sheetNameOpt);
    cmd.Options.Add(sheetIndexOpt);
    cmd.Options.Add(noHeaderOpt);
    cmd.Options.Add(fillMergedOpt);
    cmd.Options.Add(singlePassOpt);
    cmd.Options.Add(outputOpt);
    cmd.Options.Add(passwordOpt);
    cmd.Options.Add(encodingOpt);
    cmd.Options.Add(dataSetOpt);

    cmd.SetAction(parseResult =>
    {
        var file = parseResult.GetValue(fileArg)!;
        var sheetNames = parseResult.GetValue(sheetNameOpt) ?? [];
        var sheetIndexTokens = parseResult.GetValue(sheetIndexOpt) ?? [];
        var noHeader = parseResult.GetValue(noHeaderOpt);
        var fillMerged = parseResult.GetValue(fillMergedOpt);
        var singlePass = parseResult.GetValue(singlePassOpt);
        var output = parseResult.GetValue(outputOpt);
        var password = parseResult.GetValue(passwordOpt);
        var encodingName = parseResult.GetValue(encodingOpt)!;
        var useDataSet = parseResult.GetValue(dataSetOpt) || fillMerged;

        // Expand comma-separated index tokens into a set of 1-based indices.
        var nameSet = new HashSet<string>(sheetNames, StringComparer.Ordinal);
        var indexSet = new HashSet<int>();
        foreach (var token in sheetIndexTokens)
        {
            foreach (var part in token.Split(','))
            {
                if (!int.TryParse(part.Trim(), out int idx) || idx < 1)
                {
                    Console.Error.WriteLine($"Invalid sheet index: '{part.Trim()}' (must be a positive integer)");
                    return;
                }

                indexSet.Add(idx);
            }
        }

        bool hasFilter = nameSet.Count > 0 || indexSet.Count > 0;

        var memBefore = Process.GetCurrentProcess().WorkingSet64;
        var sw = Stopwatch.StartNew();

        using var stream = file.OpenRead();
        using var reader = ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration
        {
            Password = password,
            FallbackEncoding = Encoding.GetEncoding(encodingName),
            SinglePassMode = singlePass,
        });

        var openMs = sw.ElapsedMilliseconds;

        if (useDataSet)
        {
            var ds = reader.AsDataSet(new ExcelDataSetConfiguration
            {
                UseColumnDataType = false,
                FilterSheet = (tableReader, sheetIndex) =>
                {
                    if (!hasFilter)
                        return true;

                    // sheetIndex is 0-based; expose as 1-based to the user.
                    return nameSet.Contains(tableReader.Name) || indexSet.Contains(sheetIndex + 1);
                },
                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = !noHeader,
                    FillMergedCellsValue = fillMerged,
                }
            });

            var readMs = sw.ElapsedMilliseconds - openMs;
            var memAfter = Process.GetCurrentProcess().WorkingSet64;

            PrintStats(openMs, readMs, ds, memBefore, memAfter);

            if (output != OutputFormat.None)
                RenderSheets(ds, output);
        }
        else
        {
            var sheets = ReadRaw(reader, noHeader, hasFilter, nameSet, indexSet, output);

            var readMs = sw.ElapsedMilliseconds - openMs;
            var memAfter = Process.GetCurrentProcess().WorkingSet64;

            PrintRawStats(openMs, readMs, sheets, memBefore, memAfter);
        }
    });

    return cmd;
}

// ---- csv subcommand ---------------------------------------------------------
static Command BuildCsvCommand()
{
    var fileArg = new Argument<FileInfo>("file") { Description = "Path to the CSV file" };
    var noHeaderOpt = new Option<bool>("--no-header", ["-H"]) { Description = "Don't treat first row as column names" };
    var outputOpt = new Option<OutputFormat>("--output", ["-o"]) { Description = "Data output format: table, csv, tsv (default: no data output, stats only)", DefaultValueFactory = _ => OutputFormat.None };
    var encodingOpt = new Option<string>("--encoding", ["-e"]) { Description = "Fallback encoding when no BOM / not UTF-8 (default: windows-1252)", DefaultValueFactory = _ => "windows-1252" };
    var noTrimOpt = new Option<bool>("--no-trim") { Description = "Don't trim whitespace in values" };
    var separatorsOpt = new Option<string?>("--separators") { Description = "Separator candidates, e.g. \",;\" -- use \\t for TAB (default: , ; TAB | #)" };
    var quoteCharOpt = new Option<string?>("--quote-char") { Description = "Quote character (default: \")" };
    var escapeCharOpt = new Option<string?>("--escape-char") { Description = "Escape character for quoted fields (default: disabled)" };
    var dataSetOpt = new Option<bool>("--dataset") { Description = "Use AsDataSet extension (loads all data into a DataSet in memory)" };

    var cmd = new Command("csv", "Read CSV files");
    cmd.Arguments.Add(fileArg);
    cmd.Options.Add(noHeaderOpt);
    cmd.Options.Add(outputOpt);
    cmd.Options.Add(encodingOpt);
    cmd.Options.Add(noTrimOpt);
    cmd.Options.Add(separatorsOpt);
    cmd.Options.Add(quoteCharOpt);
    cmd.Options.Add(escapeCharOpt);
    cmd.Options.Add(dataSetOpt);

    cmd.SetAction(parseResult =>
    {
        var file = parseResult.GetValue(fileArg)!;
        var noHeader = parseResult.GetValue(noHeaderOpt);
        var output = parseResult.GetValue(outputOpt);
        var encodingName = parseResult.GetValue(encodingOpt)!;
        var noTrim = parseResult.GetValue(noTrimOpt);
        var separatorsStr = parseResult.GetValue(separatorsOpt);
        var quoteCharStr = parseResult.GetValue(quoteCharOpt);
        var escapeCharStr = parseResult.GetValue(escapeCharOpt);
        var useDataSet = parseResult.GetValue(dataSetOpt);

        var config = new ExcelReaderConfiguration
        {
            FallbackEncoding = Encoding.GetEncoding(encodingName),
            TrimWhiteSpace = !noTrim,
            QuoteChar = quoteCharStr?.Length > 0 ? quoteCharStr[0] : '"',
            EscapeChar = escapeCharStr?.Length > 0 ? escapeCharStr[0] : null,
        };
        if (separatorsStr is not null)
            config.AutodetectSeparators = separatorsStr.Replace("\\t", "\t").ToCharArray();

        var memBefore = Process.GetCurrentProcess().WorkingSet64;
        var sw = Stopwatch.StartNew();

        using var stream = file.OpenRead();
        using var reader = ExcelReaderFactory.CreateCsvReader(stream, config);

        var openMs = sw.ElapsedMilliseconds;

        if (useDataSet)
        {
            var ds = reader.AsDataSet(new ExcelDataSetConfiguration
            {
                UseColumnDataType = false,
                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = !noHeader,
                }
            });

            var readMs = sw.ElapsedMilliseconds - openMs;
            var memAfter = Process.GetCurrentProcess().WorkingSet64;

            PrintStats(openMs, readMs, ds, memBefore, memAfter);

            if (output != OutputFormat.None)
                RenderSheets(ds, output);
        }
        else
        {
            var sheets = ReadRaw(reader, noHeader, hasFilter: false, nameSet: [], indexSet: [], output);

            var readMs = sw.ElapsedMilliseconds - openMs;
            var memAfter = Process.GetCurrentProcess().WorkingSet64;

            PrintRawStats(openMs, readMs, sheets, memBefore, memAfter);
        }
    });

    return cmd;
}

// ---- raw reader loop --------------------------------------------------------
static List<(string Name, long Rows, int Cols)> ReadRaw(
    IExcelDataReader reader,
    bool noHeader,
    bool hasFilter,
    HashSet<string> nameSet,
    HashSet<int> indexSet,
    OutputFormat output)
{
    var sheets = new List<(string Name, long Rows, int Cols)>();
    bool firstOutput = true;
    int sheetNumber = 0;
    const int tableRowCap = 100;

    do
    {
        sheetNumber++;
        if (hasFilter && !nameSet.Contains(reader.Name) && !indexSet.Contains(sheetNumber))
            continue;

        string sheetName = reader.Name;
        long sheetRows = 0;
        int sheetCols = 0;
        string[]? headers = null;
        List<string[]>? tableBuffer = output == OutputFormat.Table ? [] : null;

        // Read header row if applicable.
        if (!noHeader && reader.Read())
        {
            headers = new string[reader.FieldCount];
            for (int i = 0; i < reader.FieldCount; i++)
                headers[i] = reader.GetValue(i)?.ToString() ?? string.Empty;

            sheetCols = Math.Max(sheetCols, reader.FieldCount);
        }

        // Emit header to separated output immediately.
        if (output is OutputFormat.Csv or OutputFormat.Tsv)
        {
            char sep = output == OutputFormat.Csv ? ',' : '\t';
            if (!firstOutput)
                Console.WriteLine();

            if (sheets.Count > 0)
                Console.WriteLine($"# Sheet: {sheetName}");

            if (headers is not null)
                WriteSeparatedRow(headers, sep);
        }

        // Read data rows.
        while (reader.Read())
        {
            sheetRows++;
            sheetCols = Math.Max(sheetCols, reader.FieldCount);

            var cells = new string[reader.FieldCount];
            for (int i = 0; i < reader.FieldCount; i++)
                cells[i] = reader.GetValue(i)?.ToString() ?? string.Empty;

            if (output is OutputFormat.Csv or OutputFormat.Tsv)
                WriteSeparatedRow(cells, output == OutputFormat.Csv ? ',' : '\t');
            else if (output == OutputFormat.Table && sheetRows <= tableRowCap)
                tableBuffer!.Add(cells);
        }

        // Render buffered table output.
        if (output == OutputFormat.Table)
        {
            if (sheets.Count > 0)
                AnsiConsole.MarkupLine($"[bold]{Markup.Escape(sheetName)}[/]");

            if (sheetRows > tableRowCap)
                Console.Error.WriteLine($"(showing first {tableRowCap:N0} of {sheetRows:N0} rows — use --output csv for full output)");

            RenderRawTable(headers, tableBuffer!);
        }

        sheets.Add((sheetName, sheetRows, sheetCols));
        firstOutput = false;
    }
    while (reader.NextResult());

    return sheets;
}

// ---- helpers ----------------------------------------------------------------
static void PrintStats(long openMs, long readMs, DataSet ds, long memBefore, long memAfter)
{
    long totalRows = 0;
    int maxCols = 0;

    foreach (DataTable dt in ds.Tables)
    {
        totalRows += dt.Rows.Count;
        if (dt.Columns.Count > maxCols)
            maxCols = dt.Columns.Count;
    }

    long totalMs = openMs + readMs;
    string rateStr = BuildRateStr(totalRows * maxCols, totalMs);
    long memDeltaMb = (memAfter - memBefore) / (1024 * 1024);

    Console.Error.WriteLine($"Open:   {openMs,6:N0} ms");
    Console.Error.Write($"Read:   {readMs,6:N0} ms  ({totalRows:N0} rows x {maxCols} cols");
    if (ds.Tables.Count > 1)
    {
        Console.Error.WriteLine($" across {ds.Tables.Count} sheets)");
        foreach (DataTable dt in ds.Tables)
            Console.Error.WriteLine($"  {dt.TableName}: {dt.Rows.Count:N0} rows x {dt.Columns.Count} cols");
    }
    else
    {
        Console.Error.WriteLine(")");
    }

    Console.Error.WriteLine($"Total:  {totalMs,6:N0} ms{rateStr}");
    Console.Error.WriteLine($"Memory: {(memDeltaMb >= 0 ? "+" : string.Empty)}{memDeltaMb} MB");
}

static void PrintRawStats(long openMs, long readMs, List<(string Name, long Rows, int Cols)> sheets, long memBefore, long memAfter)
{
    long totalRows = sheets.Sum(s => s.Rows);
    int maxCols = sheets.Count > 0 ? sheets.Max(s => s.Cols) : 0;

    long totalMs = openMs + readMs;
    string rateStr = BuildRateStr(totalRows * maxCols, totalMs);
    long memDeltaMb = (memAfter - memBefore) / (1024 * 1024);

    Console.Error.WriteLine($"Open:   {openMs,6:N0} ms");
    Console.Error.Write($"Read:   {readMs,6:N0} ms  ({totalRows:N0} rows x {maxCols} cols");
    if (sheets.Count > 1)
    {
        Console.Error.WriteLine($" across {sheets.Count} sheets)");
        foreach (var (name, rows, cols) in sheets)
            Console.Error.WriteLine($"  {name}: {rows:N0} rows x {cols} cols");
    }
    else
    {
        Console.Error.WriteLine(")");
    }

    Console.Error.WriteLine($"Total:  {totalMs,6:N0} ms{rateStr}");
    Console.Error.WriteLine($"Memory: {(memDeltaMb >= 0 ? "+" : string.Empty)}{memDeltaMb} MB");
}

static string BuildRateStr(long totalCells, long totalMs)
{
    if (totalMs <= 0)
        return string.Empty;

    double cellsPerSec = totalCells * 1000.0 / totalMs;
    return cellsPerSec >= 1_000_000
        ? $"  (~{cellsPerSec / 1_000_000:F1}M cells/sec)"
        : cellsPerSec >= 1_000
        ? $"  (~{cellsPerSec / 1_000:F1}K cells/sec)"
        : $"  (~{cellsPerSec:F0} cells/sec)";
}

static void RenderSheets(DataSet ds, OutputFormat output)
{
    bool multiSheet = ds.Tables.Count > 1;
    bool first = true;

    foreach (DataTable dt in ds.Tables)
    {
        if (output == OutputFormat.Table)
        {
            if (multiSheet)
                AnsiConsole.MarkupLine($"[bold]{Markup.Escape(dt.TableName)}[/]");
            RenderSpectreTable(dt);
        }
        else
        {
            if (!first)
                Console.WriteLine();
            if (multiSheet)
                Console.WriteLine($"# Sheet: {dt.TableName}");
            WriteSeparated(dt, output == OutputFormat.Csv ? ',' : '\t');
        }

        first = false;
    }
}

static void RenderSpectreTable(DataTable dt)
{
    const int tableRowCap = 100;
    var table = new Table();
    foreach (DataColumn col in dt.Columns)
        table.AddColumn(Markup.Escape(col.ColumnName));

    int rendered = 0;
    foreach (DataRow row in dt.Rows)
    {
        if (rendered >= tableRowCap)
            break;

        table.AddRow(row.ItemArray.Select(v => Markup.Escape(v?.ToString() ?? string.Empty)).ToArray());
        rendered++;
    }

    if (dt.Rows.Count > tableRowCap)
        Console.Error.WriteLine($"(showing first {tableRowCap:N0} of {dt.Rows.Count:N0} rows — use --output csv for full output)");

    AnsiConsole.Write(table);
}

static void RenderRawTable(string[]? headers, List<string[]> rows)
{
    // Determine the actual maximum column count across headers and all rows.
    // In single-pass mode rows may have fewer fields than the header (trailing
    // empty cells are not stored), and later rows may have more columns than
    // earlier ones. Pad everything to the widest row.
    int colCount = headers?.Length ?? 0;
    foreach (var row in rows)
        colCount = Math.Max(colCount, row.Length);

    var table = new Table();
    if (headers is not null)
    {
        foreach (var h in headers)
            table.AddColumn(Markup.Escape(h));

        for (int i = headers.Length; i < colCount; i++)
            table.AddColumn($"Col{i + 1}");
    }
    else
    {
        for (int i = 0; i < colCount; i++)
            table.AddColumn($"Col{i + 1}");
    }

    foreach (var row in rows)
    {
        var cells = new string[colCount];
        for (int i = 0; i < colCount; i++)
            cells[i] = i < row.Length ? Markup.Escape(row[i]) : string.Empty;

        table.AddRow(cells);
    }

    AnsiConsole.Write(table);
}

static void WriteSeparated(DataTable dt, char sep)
{
    static string Escape(string s, char sep) =>
        s.Contains(sep) || s.Contains('"') || s.Contains('\n') || s.Contains('\r')
            ? $"\"{s.Replace("\"", "\"\"")}\""
            : s;

    foreach (DataRow row in dt.Rows)
        Console.WriteLine(string.Join(sep, row.ItemArray.Select(v => Escape(v?.ToString() ?? string.Empty, sep))));
}

static void WriteSeparatedRow(string[] cells, char sep)
{
    static string Escape(string s, char sep) =>
        s.Contains(sep) || s.Contains('"') || s.Contains('\n') || s.Contains('\r')
            ? $"\"{s.Replace("\"", "\"\"")}\""
            : s;

    Console.WriteLine(string.Join(sep, cells.Select(v => Escape(v, sep))));
}

#pragma warning disable SA1649 // File name should match first type name (top-level statements use implicit Program class)
internal enum OutputFormat
{
    None,
    Table,
    Csv,
    Tsv,
}
