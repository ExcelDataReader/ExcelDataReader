#nullable enable

namespace ExcelDataReader.Core
{
    internal class Column
    {
        public Column(int minimum, int maximum, bool hidden, double? width)
        {
            Minimum = minimum;
            Maximum = maximum;
            Hidden = hidden;
            Width = width;
        }

        public int Minimum { get; }

        public int Maximum { get; }

        public bool Hidden { get; }

        public double? Width { get; }
    }
}
