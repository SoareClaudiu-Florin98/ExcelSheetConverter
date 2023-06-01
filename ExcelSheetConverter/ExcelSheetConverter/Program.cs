namespace ExcelSheetConverter
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Test Column Letter to Column Number
            string columnTitle = "A";
            int columnNumber = ExcelSheetConverter.TitleToNumber(columnTitle);
            Console.WriteLine($"Column Title: {columnTitle}, Column Number: {columnNumber}");

            // Test Column Number to Column Letter
            int number = 702;
            string columnName = ExcelSheetConverter.NumberToTitle(number);
            Console.WriteLine($"Number: {number}, Column Name: {columnName}");
        }
    }
}