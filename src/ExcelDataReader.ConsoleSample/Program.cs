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
    var fillMergedOpt = new Option<bool>("--fill-merged") { Description = "Fill merged cell values across the merged range" };
    var singlePassOpt = new Option<bool>("--single-pass") { Description = "Enable single pass mode (skips pre-scan for row/column counts)" };
    var outputOpt = new Option<OutputFormat>("--output", ["-o"]) { Description = "Data output format: table, csv, tsv (default: no data output, stats only)", DefaultValueFactory = _ => OutputFormat.None };
    var passwordOpt = new Option<string?>("--password", ["-p"]) { Description = "Password for protected workbooks" };
    var encodingOpt = new Option<string>("--encoding", ["-e"]) { Description = "Fallback encoding for XLS BIFF2-5 (default: windows-1252)", DefaultValueFactory = _ => "windows-1252" };

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

    cmd.SetAction(parseResult =>
    {
        var file = parseResult.GetValue(fileArg)!;
        var sheetNames = parseResult.GetValue(sheetNameOpt) ?? [];
        var sheetIndexTokens = parseResult.GetValue(sheetIndexOpt) ?? [];
        var noHeader = parseResult.GetValue(noHeaderOpt);
        var noFillMerged = !parseResult.GetValue(fillMergedOpt);
        var singlePass = parseResult.GetValue(singlePassOpt);
        var output = parseResult.GetValue(outputOpt);
        var password = parseResult.GetValue(passwordOpt);
        var encodingName = parseResult.GetValue(encodingOpt)!;

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
                FillMergedCellsValue = !noFillMerged,
            }
        });

        var readMs = sw.ElapsedMilliseconds - openMs;
        var memAfter = Process.GetCurrentProcess().WorkingSet64;

        PrintStats(openMs, readMs, ds, memBefore, memAfter);

        if (output != OutputFormat.None)
            RenderSheets(ds, output);
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

    var cmd = new Command("csv", "Read CSV files");
    cmd.Arguments.Add(fileArg);
    cmd.Options.Add(noHeaderOpt);
    cmd.Options.Add(outputOpt);
    cmd.Options.Add(encodingOpt);
    cmd.Options.Add(noTrimOpt);
    cmd.Options.Add(separatorsOpt);
    cmd.Options.Add(quoteCharOpt);
    cmd.Options.Add(escapeCharOpt);

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
    });

    return cmd;
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

    long totalCells = totalRows * maxCols;
    double cellsPerSec = readMs > 0 ? totalCells * 1000.0 / readMs : double.PositiveInfinity;
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

    Console.Error.WriteLine($"Total:  {openMs + readMs,6:N0} ms  (~{cellsPerSec / 1_000_000:F1}M cells/sec)");
    Console.Error.WriteLine($"Memory: {(memDeltaMb >= 0 ? "+" : string.Empty)}{memDeltaMb} MB");
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
    var table = new Table();
    foreach (DataColumn col in dt.Columns)
        table.AddColumn(Markup.Escape(col.ColumnName));

    foreach (DataRow row in dt.Rows)
        table.AddRow(row.ItemArray.Select(v => Markup.Escape(v?.ToString() ?? string.Empty)).ToArray());

    AnsiConsole.Write(table);
}

static void WriteSeparated(DataTable dt, char sep)
{
    static string Escape(string s, char sep) =>
        s.Contains(sep) || s.Contains('"') || s.Contains('\n')
            ? $"\"{s.Replace("\"", "\"\"")}\""
            : s;

    foreach (DataRow row in dt.Rows)
        Console.WriteLine(string.Join(sep, row.ItemArray.Select(v => Escape(v?.ToString() ?? string.Empty, sep))));
}

#pragma warning disable SA1649 // File name should match first type name (top-level statements use implicit Program class)
internal enum OutputFormat
{
    None,
    Table,
    Csv,
    Tsv,
}