using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Xml;
using ExcelDataReader.Core.OpenXmlFormat.BinaryFormat;
using ExcelDataReader.Core.OpenXmlFormat.XmlFormat;

#nullable enable

namespace ExcelDataReader.Core.OpenXmlFormat
{
    internal sealed partial class ZipWorker : IDisposable
    {
        private const string DefaultFileWorkbook = "xl/workbook.";

        private const string Format = "xml";
        private const string BinFormat = "bin";

        private static readonly XmlReaderSettings XmlSettings = new() 
        {
            IgnoreComments = true, 
            IgnoreWhitespace = true,
        };

        private readonly Dictionary<string, ZipArchiveEntry> _entries = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> _worksheetRels = new();

        private readonly string _fileWorkbook;
        private readonly string? _fileSharedStrings;
        private readonly string? _fileStyles;

        private ZipArchive? _zipFile;

        /// <summary>
        /// Initializes a new instance of the <see cref="ZipWorker"/> class. 
        /// </summary>
        /// <param name="fileStream">The zip file stream.</param>
        public ZipWorker(Stream fileStream)
        {
            _zipFile = new ZipArchive(fileStream ?? throw new ArgumentNullException(nameof(fileStream)));

            // Entries use '/' but not if Switch.System.IO.Compression.ZipFile.UseBackslash compat switch is enabled
            foreach (var entry in _zipFile.Entries)
            {
                _entries.Add(entry.FullName.Replace('\\', '/'), entry);
            }

            var fileWorkbook = ReadRootRels();
            if (fileWorkbook == null || !_entries.ContainsKey(fileWorkbook))
            {
                fileWorkbook = CheckPath(DefaultFileWorkbook + Format) ?? CheckPath(DefaultFileWorkbook + BinFormat);
            }

            _fileWorkbook = fileWorkbook ?? throw new Exceptions.HeaderException(Errors.ErrorZipNoOpenXml);

            string[] parts = _fileWorkbook.Split('/');
            string? basePath = parts.Length <= 1 ? null : string.Join("/", parts, 0, parts.Length - 1) + "/";
            string path = basePath + "_rels/" + parts[parts.Length - 1] + ".rels";
            var workbookRelsEntry = FindEntry(path);
            if (workbookRelsEntry == null)
                return;

            using var reader = XmlReader.Create(workbookRelsEntry.Open(), XmlSettings);
            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.Element || reader.Name != "Relationship")
                    continue;

                var id = reader.GetAttribute("Id");
                var type = reader.GetAttribute("Type");
                var target = reader.GetAttribute("Target");

                switch (type)
                {
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet":
                    case "http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet":
                        _worksheetRels[id] = ResolvePath(basePath, target);
                        break;
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles":
                    case "http://purl.oclc.org/ooxml/officeDocument/relationships/styles":
                        _fileStyles = ResolvePath(basePath, target);
                        break;
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings":
                    case "http://purl.oclc.org/ooxml/officeDocument/relationships/sharedStrings":
                        _fileSharedStrings = ResolvePath(basePath, target);
                        break;
                }
            }

            static string ResolvePath(string? basePath, string path)
            {
                // Can there be relative paths?
                if (path.StartsWith("/", StringComparison.Ordinal))
                    return path.Substring(1);
                return basePath + path;
            }

            string? CheckPath(string path)
            {
                if (_entries.ContainsKey(path))
                    return path;
                return null;
            }

            string? ReadRootRels()
            {
                var entry = FindEntry("_rels/.rels");
                if (entry == null)
                    return null;

                using var reader = XmlReader.Create(entry.Open(), XmlSettings);
                while (reader.Read())
                {
                    if (reader.NodeType != XmlNodeType.Element || reader.Name != "Relationship")
                        continue;

                    var type = reader.GetAttribute("Type");
                    var target = reader.GetAttribute("Target");

                    switch (type)
                    {
                        case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument":
                        case "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument":
                            return target;
                    }
                }

                return null;
            }
        }

        /// <summary>
        /// Gets the shared strings reader.
        /// </summary>
        public RecordReader? GetSharedStringsReader()
        {
            if (FindEntry(_fileSharedStrings) is { } entry)
            {
                if (entry.FullName.EndsWith(".xml", StringComparison.Ordinal))
                    return new XmlSharedStringsReader(XmlReader.Create(entry.Open(), XmlSettings));

                if (entry.FullName.EndsWith(".bin", StringComparison.Ordinal))
                    return new BiffSharedStringsReader(entry.Open());
            }

            return null;
        }

        /// <summary>
        /// Gets the styles reader.
        /// </summary>
        public RecordReader? GetStylesReader()
        {
            if (FindEntry(_fileStyles) is { } entry)
            {
                if (entry.FullName.EndsWith(".xml", StringComparison.Ordinal))
                    return new XmlStylesReader(XmlReader.Create(entry.Open(), XmlSettings));

                if (entry.FullName.EndsWith(".bin", StringComparison.Ordinal))
                    return new BiffStylesReader(entry.Open());
            }

            return null;
        }

        /// <summary>
        /// Gets the workbook reader.
        /// </summary>
        public RecordReader? GetWorkbookReader()
        {
            if (FindEntry(_fileWorkbook) is { } entry)
            { 
                if (entry.FullName.EndsWith(".xml", StringComparison.Ordinal))
                    return new XmlWorkbookReader(XmlReader.Create(entry.Open(), XmlSettings), _worksheetRels);
                else if (entry.FullName.EndsWith(".bin", StringComparison.Ordinal))
                    return new BiffWorkbookReader(entry.Open(), _worksheetRels);
            }

            throw new Exceptions.HeaderException(Errors.ErrorZipNoOpenXml);
        }

        public RecordReader? GetWorksheetReader(string sheetPath)
        {
            // its possible sheetPath starts with /xl. in this case trim the /
            // see the test "Issue_11522_OpenXml"
            if (sheetPath.StartsWith("/xl/", StringComparison.OrdinalIgnoreCase))
                sheetPath = sheetPath.Substring(1);

            var zipEntry = FindEntry(sheetPath);
            if (zipEntry != null)
            {
                return Path.GetExtension(sheetPath) switch
                {
                    ".xml" => new XmlWorksheetReader(XmlReader.Create(zipEntry.Open(), XmlSettings)),
                    ".bin" => new BiffWorksheetReader(zipEntry.Open()),
                    _ => null,
                };
            }

            return null;
        }

        private ZipArchiveEntry? FindEntry(string? name)
        {
            if (name != null && _entries.TryGetValue(name, out var entry))
                return entry;
            return null;
        }
    }

    internal partial class ZipWorker
    {
        ~ZipWorker()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);

            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            if (disposing)
            {
                _zipFile?.Dispose();
                _zipFile = null;
            }
        }
    }
}