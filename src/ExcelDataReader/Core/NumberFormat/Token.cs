using System;

namespace ExcelDataReader.Core.NumberFormat
{
    internal static class Token
    {
        public static bool IsExponent(string token)
        {
            return
                (string.Compare(token, "e+", StringComparison.OrdinalIgnoreCase) == 0) ||
                (string.Compare(token, "e-", StringComparison.OrdinalIgnoreCase) == 0);
        }

        public static bool IsLiteral(string token)
        {
            return
                token.StartsWith("_") ||
                token.StartsWith("\\") ||
                token.StartsWith("\"") ||
                token.StartsWith("*") ||
                token == "," ||
                token == "!" ||
                token == "&" ||
                token == "%" ||
                token == "+" ||
                token == "-" ||
                token == "$" ||
                token == "€" ||
                token == "£" ||
                token == "1" ||
                token == "2" ||
                token == "3" ||
                token == "4" ||
                token == "5" ||
                token == "6" ||
                token == "7" ||
                token == "8" ||
                token == "9" ||
                token == "{" ||
                token == "}" ||
                token == "(" ||
                token == ")" ||
                token == " ";
        }

        public static bool IsNumberLiteral(string token)
        {
            return
                IsPlaceholder(token) ||
                IsLiteral(token) ||
                token == ".";
        }

        public static bool IsPlaceholder(string token)
        {
            return token == "0" || token == "#" || token == "?";
        }

        public static bool IsGeneral(string token)
        {
            return string.Compare(token, "general", StringComparison.OrdinalIgnoreCase) == 0;
        }

        public static bool IsDatePart(string token)
        {
            return
                token.StartsWith("y", StringComparison.OrdinalIgnoreCase) ||
                token.StartsWith("m", StringComparison.OrdinalIgnoreCase) ||
                token.StartsWith("d", StringComparison.OrdinalIgnoreCase) ||
                token.StartsWith("s", StringComparison.OrdinalIgnoreCase) ||
                token.StartsWith("h", StringComparison.OrdinalIgnoreCase) ||
                (token.StartsWith("g", StringComparison.OrdinalIgnoreCase) && !IsGeneral(token)) ||
                string.Compare(token, "am/pm", StringComparison.OrdinalIgnoreCase) == 0 ||
                string.Compare(token, "a/p", StringComparison.OrdinalIgnoreCase) == 0 ||
                IsDurationPart(token);
        }

        public static bool IsDurationPart(string token)
        {
            return
                token.StartsWith("[h", StringComparison.OrdinalIgnoreCase) ||
                token.StartsWith("[m", StringComparison.OrdinalIgnoreCase) ||
                token.StartsWith("[s", StringComparison.OrdinalIgnoreCase);
        }

        public static bool IsDigit09(string token)
        {
            return token == "0" || IsDigit19(token);
        }

        public static bool IsDigit19(string token)
        {
            switch (token)
            {
                case "1":
                case "2":
                case "3":
                case "4":
                case "5":
                case "6":
                case "7":
                case "8":
                case "9":
                    return true;
                default:
                    return false;
            }
        }
    }
}
