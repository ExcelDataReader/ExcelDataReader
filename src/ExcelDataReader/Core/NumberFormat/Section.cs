using System.Collections.Generic;

namespace ExcelDataReader.Core.NumberFormat
{
    internal class Section
    {
        public SectionType Type { get; set; }

        public Color Color { get; set; }

        public Condition Condition { get; set; }

        public ExponentialSection Exponential { get; set; }

        public FractionSection Fraction { get; set; }

        public DecimalSection Number { get; set; }

        public List<string> GeneralTextDateDurationParts { get; set; }
    }
}