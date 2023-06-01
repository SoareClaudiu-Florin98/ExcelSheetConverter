namespace ExcelSheetConverter
{
    public class ExcelSheetConverter
    {
        private const int numberOfAlphabetLetters = 26;
        private const int offsetAdjustment = 1;
        private const char startChar = 'A'; //41 is ASCII table

        public static int TitleToNumber(string columnTitle)
        {
            int result = 0;

            for (int i = 0; i < columnTitle.Length; i++)
            {
                char currentChar = columnTitle[i];
                int columnValue = currentChar - startChar + offsetAdjustment;

                // ex: currentChar = B = 42 in ASCII
                // startChar = A = 41 in ASCII  => columnValue = 2 for B

                result = result * numberOfAlphabetLetters + columnValue;
            }
            return result;
        }

        public static string NumberToTitle(int columnNumber)
        {
            string result = string.Empty;

            while (columnNumber > 0)
            {
                int remainder = (columnNumber - offsetAdjustment) % numberOfAlphabetLetters;
                char letter = (char)(startChar + remainder);
                result = letter + result;
                columnNumber = (columnNumber - offsetAdjustment) / numberOfAlphabetLetters;
            }
            return result;
        }
    }
}
