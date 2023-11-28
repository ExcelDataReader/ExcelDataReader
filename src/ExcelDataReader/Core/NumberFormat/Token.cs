using System;

namespace ExcelDataReader.Core.NumberFormat
{
    internal static class Token
    {
        public static bool IsExponent(string token)
        {
            return
                string.Equals(token, "e+", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(token, "e-", StringComparison.OrdinalIgnoreCase);
        }

        public static bool IsLiteral(string token)
        {
            return
#if NETSTANDARD2_1_OR_GREATER
                token.StartsWith('_') ||
                token.StartsWith('\\') ||
                token.StartsWith('\"') ||
                token.StartsWith('*') ||
#else
                token.StartsWith("_", StringComparison.Ordinal) ||
                token.StartsWith("\\", StringComparison.Ordinal) ||
                token.StartsWith("\"", StringComparison.Ordinal) ||
                token.StartsWith("*", StringComparison.Ordinal) ||
#endif
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
            return string.Equals(token, "general", StringComparison.OrdinalIgnoreCase);
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
                string.Equals(token, "am/pm", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(token, "a/p", StringComparison.OrdinalIgnoreCase) ||
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

        public static bool IsDigit19(string token) => token switch
        {
            "1" or "2" or "3" or "4" or "5" or "6" or "7" or "8" or "9" => true,
            _ => false,
        };
    }
}
