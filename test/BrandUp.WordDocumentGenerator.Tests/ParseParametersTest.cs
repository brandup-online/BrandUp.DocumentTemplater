namespace BrandUp.DocumentTemplater
{
    public class ParseParametersTest
    {
        /// <summary>
        /// Запятая внутри кавычек не должна разбивать параметр —
        /// иначе ломаются форматы вида "#,##0.00".
        /// </summary>
        [Fact]
        public void CommaInsideQuotesIsNotSeparator()
        {
            var result = WordDocumentTemplater.ParseParameters("Amount, \"#,##0.00\"");

            Assert.Equal(2, result.Count);
            Assert.Equal("Amount", result[0]);
            Assert.Equal("#,##0.00", result[1]);
        }

        [Fact]
        public void SingleQuotesAreSupported()
        {
            var result = WordDocumentTemplater.ParseParameters("Date, 'dd, MM, yyyy'");

            Assert.Equal(2, result.Count);
            Assert.Equal("Date", result[0]);
            Assert.Equal("dd, MM, yyyy", result[1]);
        }

        [Fact]
        public void TrimsWhitespaceAroundUnquotedValues()
        {
            var result = WordDocumentTemplater.ParseParameters(" Tax ,  ##;##;без НДС ");

            Assert.Equal(2, result.Count);
            Assert.Equal("Tax", result[0]);
            Assert.Equal("##;##;без НДС", result[1]);
        }

        [Fact]
        public void EmptyInputReturnsEmptyList()
        {
            Assert.Empty(WordDocumentTemplater.ParseParameters(""));
            Assert.Empty(WordDocumentTemplater.ParseParameters(null));
        }

        [Fact]
        public void EmptyEntriesAreOmitted()
        {
            var result = WordDocumentTemplater.ParseParameters("a,,b");

            Assert.Equal(2, result.Count);
            Assert.Equal("a", result[0]);
            Assert.Equal("b", result[1]);
        }
    }
}
