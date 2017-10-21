namespace ExcelDataReader.Core
{
    internal static class BuiltinNumberFormat
    {
        public static string GetBuiltinNumberFormat(int numFmtId)
        {
            switch (numFmtId)
            {
                case 0: return "General";
                case 1: return "0";
                case 2: return "0.00";
                case 3: return "#,##0";
                case 4: return "#,##0.00";
                case 5: return "\"$\"#,##0_);(\"$\"#,##0)";
                case 6: return "\"$\"#,##0_);[Red](\"$\"#,##0)";
                case 7: return "\"$\"#,##0.00_);(\"$\"#,##0.00)";
                case 8: return "\"$\"#,##0.00_);[Red](\"$\"#,##0.00)";
                case 9: return "0%";
                case 10: return "0.00%";
                case 11: return "0.00E+00";
                case 12: return "# ?/?";
                case 13: return "# ??/??";
                case 14: return "mm-dd-yy";
                case 15: return "d-mmm-yy";
                case 16: return "d-mmm";
                case 17: return "mmm-yy";
                case 18: return "h:mm AM/PM";
                case 19: return "h:mm:ss AM/PM";
                case 20: return "h:mm";
                case 21: return "h:mm:ss";
                case 22: return "m/d/yy h:mm";

                // 23..36 international/unused
                case 37: return "#,##0_);(#,##0)";
                case 38: return "#,##0_);[Red](#,##0)";
                case 39: return "#,##0.00_);(#,##0.00)";
                case 40: return "#,##0.00_);[Red](#,##0.00)";
                case 41: return "_(\"$\"* #,##0_);_(\"$\"* (#,##0);_(\"$\"* \"-\"_);_(@_)";
                case 42: return "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)";
                case 43: return "_(\"$\"* #,##0.00_);_(\"$\"* (#,##0.00);_(\"$\"* \"-\"??_);_(@_)";
                case 44: return "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)";
                case 45: return "mm:ss";
                case 46: return "[h]:mm:ss";
                case 47: return "mm:ss.0";
                case 48: return "##0.0E+0";
                case 49: return "@";
            }

            return null;
        }
    }
}
