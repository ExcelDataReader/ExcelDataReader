#nullable enable

namespace ExcelDataReader.Core
{
    internal sealed class Column
    {
        public Column(int minimum, int maximum, bool isHidden, double? width)
        {
            Minimum = minimum;
            Maximum = maximum;
            IsHidden = isHidden;
            Width = width;
        }

        public int Minimum { get; }

        public int Maximum { get; }

        public bool IsHidden { get; }

        public double? Width { get; }
    }
}
