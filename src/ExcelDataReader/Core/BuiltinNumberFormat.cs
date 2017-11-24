using System.Collections.Generic;
using ExcelDataReader.Core.NumberFormat;

namespace ExcelDataReader.Core
{
    internal static class BuiltinNumberFormat
    {
        private static Dictionary<int, NumberFormatString> Formats { get; } = new Dictionary<int, NumberFormatString>()
        {
            { 0, new NumberFormatString("General") },
            { 1, new NumberFormatString("0") },
            { 2, new NumberFormatString("0.00") },
            { 3, new NumberFormatString("#,##0") },
            { 4, new NumberFormatString("#,##0.00") },
            { 5, new NumberFormatString("\"$\"#,##0_);(\"$\"#,##0)") },
            { 6, new NumberFormatString("\"$\"#,##0_);[Red](\"$\"#,##0)") },
            { 7, new NumberFormatString("\"$\"#,##0.00_);(\"$\"#,##0.00)") },
            { 8, new NumberFormatString("\"$\"#,##0.00_);[Red](\"$\"#,##0.00)") },
            { 9, new NumberFormatString("0%") },
            { 10, new NumberFormatString("0.00%") },
            { 11, new NumberFormatString("0.00E+00") },
            { 12, new NumberFormatString("# ?/?") },
            { 13, new NumberFormatString("# ??/??") },
            { 14, new NumberFormatString("d/m/yyyy") },
            { 15, new NumberFormatString("d-mmm-yy") },
            { 16, new NumberFormatString("d-mmm") },
            { 17, new NumberFormatString("mmm-yy") },
            { 18, new NumberFormatString("h:mm AM/PM") },
            { 19, new NumberFormatString("h:mm:ss AM/PM") },
            { 20, new NumberFormatString("h:mm") },
            { 21, new NumberFormatString("h:mm:ss") },
            { 22, new NumberFormatString("m/d/yy h:mm") },

            // 23..36 international/unused
            { 37, new NumberFormatString("#,##0_);(#,##0)") },
            { 38, new NumberFormatString("#,##0_);[Red](#,##0)") },
            { 39, new NumberFormatString("#,##0.00_);(#,##0.00)") },
            { 40, new NumberFormatString("#,##0.00_);[Red](#,##0.00)") },
            { 41, new NumberFormatString("_(\"$\"* #,##0_);_(\"$\"* (#,##0);_(\"$\"* \"-\"_);_(@_)") },
            { 42, new NumberFormatString("_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)") },
            { 43, new NumberFormatString("_(\"$\"* #,##0.00_);_(\"$\"* (#,##0.00);_(\"$\"* \"-\"??_);_(@_)") },
            { 44, new NumberFormatString("_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)") },
            { 45, new NumberFormatString("mm:ss") },
            { 46, new NumberFormatString("[h]:mm:ss") },
            { 47, new NumberFormatString("mm:ss.0") },
            { 48, new NumberFormatString("##0.0E+0") },
            { 49, new NumberFormatString("@") },
        };

        public static NumberFormatString GetBuiltinNumberFormat(int numFmtId)
        {
            if (Formats.TryGetValue(numFmtId, out var result))
                return result;

            return null;
        }
    }
}
