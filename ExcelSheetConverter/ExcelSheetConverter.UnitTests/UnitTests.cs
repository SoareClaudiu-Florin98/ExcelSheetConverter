namespace ExcelSheetConverter.UnitTests
{
    public class Tests
    {

        [Test]
        [TestCase("A", 1)]
        [TestCase("B", 2)]
        [TestCase("C", 3)]
        [TestCase("Z", 26)]
        [TestCase("AA", 27)]
        [TestCase("AB", 28)]
        [TestCase("BB", 54)]
        [TestCase("BC", 55)]
        [TestCase("AAA", 703)]
        public void TitleToNumberShouldReturn(string columnTitle, int number)
        {
            int returnedNumber = ExcelSheetConverter.TitleToNumber(columnTitle);

            Assert.That(returnedNumber, Is.EqualTo(number));
        }

        [Test]
        [TestCase(26,"Z")]
        [TestCase(51, "AY")]
        [TestCase(52, "AZ")]
        [TestCase(80, "CB")]
        [TestCase(676, "YZ")]
        [TestCase(702, "ZZ")]
        [TestCase(705, "AAC")]
        public void NumberToTitleShouldReturn(int columnNumber, string title)
        {
            string returnedTitle = ExcelSheetConverter.NumberToTitle(columnNumber);

            Assert.That(returnedTitle, Is.EqualTo(title));
        }
    }
}