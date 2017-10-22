using System.Collections.Generic;
using System.Globalization;
using ExcelDataReader.Core.NumberFormat;

namespace ExcelDataReader
{
    /// <summary>
    /// Parse ECMA-376 number format strings and format values like Excel and other spreadsheet softwares.
    /// </summary>
    public class NumberFormatString
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="NumberFormatString"/> class.
        /// </summary>
        /// <param name="formatString">The number format string.</param>
        public NumberFormatString(string formatString)
        {
            var tokenizer = new Tokenizer(formatString);
            var sections = new List<Section>();
            var isValid = true;
            while (true)
            {
                var section = Parser.ParseSection(tokenizer, out var syntaxError);

                if (syntaxError)
                    isValid = false;

                if (section == null)
                    break;

                sections.Add(section);
            }

            IsValid = isValid;
            FormatString = formatString;

            if (isValid)
            {
                Sections = sections;
            }
            else
            {
                Sections = new List<Section>();
            }
        }

        /// <summary>
        /// Gets a value indicating whether the number format string is valid.
        /// </summary>
        public bool IsValid { get; }

        /// <summary>
        /// Gets the number format string.
        /// </summary>
        public string FormatString { get; }

        /// <summary>
        /// Gets a value indicating whether the format represents a DateTime
        /// </summary>
        public bool IsDateTimeFormat
        {
            get
            {
                return GetFirstSection(SectionType.Date) != null;
            }
        }

#if NET20
        internal IList<Section> Sections { get; }
#else
        internal IReadOnlyList<Section> Sections { get; }
#endif

        /// <summary>
        /// Formats a value with this number format in a specified culture.
        /// </summary>
        /// <param name="value">The value to format.</param>
        /// <param name="culture">The culture to use for formatting.</param>
        /// <returns>The formatted string.</returns>
        public string Format(object value, CultureInfo culture)
        {
            return Formatter.Format(value, this, culture);
        }

        internal Section GetSection(object value)
        {
            // TODO:
            // if datetime, return first datetime section, else null
            // if string, return first text section, else null
            // if double, int, check condition, order, negative etc, exponential, fraction
            if (Sections.Count > 0)
                return Sections[0];
            return null;
        }

        private Section GetFirstSection(SectionType type)
        {
            foreach (var section in Sections)
            {
                if (section.Type == type)
                {
                    return section;
                }
            }

            return null;
        }
    }
}
