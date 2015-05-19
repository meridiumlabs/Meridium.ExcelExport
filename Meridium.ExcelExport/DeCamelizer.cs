using System.Text.RegularExpressions;

namespace Meridium.ExcelExport {
    /// <summary>
    /// Helper class that can split a string in camelcase into its separate words.
    /// </summary>
    internal static class DeCamelizer {
        /// <summary>
        /// Splits the string argument on capital letters and rejoins it with spaces.
        /// </summary>
        public static string Split(string camelCasedString) {
            if (string.IsNullOrWhiteSpace(camelCasedString)) return string.Empty;

            var parts = CamelSplitter.Split(camelCasedString);
            return string.Join(" ", parts);
        }

        private static readonly Regex CamelSplitter =
            new Regex(
                @" # Split the string if any of the following three cases are true:

                (?<=\P{Lu})(?=\p{Lu})         # Non uppcase letter followed by an uppercase letter
		        |                             # -or-
		        (?<=\p{Lu})(?=\p{Lu}\p{Ll})   # Uppercase letter followed by uppercase then any letter
                |                             # -or-
                (?<=\p{L})(?=\P{L})           # Any letter followed by a non-letter, e.g. digits",
                RegexOptions.Compiled | RegexOptions.IgnorePatternWhitespace );
    }
}