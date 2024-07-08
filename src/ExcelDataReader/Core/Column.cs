#nullable enable

namespace ExcelDataReader.Core;

internal sealed record Column(int Minimum, int Maximum, bool Hidden, double? Width);