namespace BrandUp.DocxGenerator
{
    public class WordDocumentGeneratorTest
    {
        [Fact]
        public async void Success_TestObject()
        {
            var data = new TestObject
            {
                Number = 102321,
                Values = new()
                {
                    new() { Name = "Tree", Price = 10000, Note = "Big" },
                    new() { Name = "Garland", Price = 5000, Note = "Long" },
                    new() { Name = "Tariff", Price = 700, Note = "30 days" },
                    new() { Name = "Toys", Price = 10, Note = "Round" }
                }
            };

            using var template = new MemoryStream(Properties.Resources.test);
            using var resultData = await WordDocumentGenerator.GenerateDocument(data, template, CancellationToken.None);

            using var output = File.Create("Success_TestObject.docx");
            resultData.CopyTo(output);
        }

        [Fact]
        public async void Success_Dictionary()
        {
            var data = new Dictionary<string, object>
            {
                { "Number" , 102321 },
                { "Values" , new List<Row>()
                        {
                            new() { Name = "Tree", Price = 10000, Note = "Big" },
                            new() { Name = "Garland", Price = 5000, Note = "Long" }
                        }
                }
            };

            using var template = new MemoryStream(Properties.Resources.test);
            using var resultData = await WordDocumentGenerator.GenerateDocument(data, template, CancellationToken.None);

            using var output = File.Create("Success_Dictionary.docx");
            resultData.CopyTo(output);
        }

        [Fact]
        public async void Success_Formats()
        {
            var data = new Dictionary<string, object>
            {
                { "FullDate" , new DateTime(2023,12, 22, 13, 55, 16) },
                { "Span" , new TimeSpan(13, 55, 16) },
                { "IntValue" , 12 },
                { "DecValue" ,123123123123m },
                { "FloatValue" , 0.9200f },
                { "DoubleValue" , double.MaxValue },
                { "Yes" , true },
                { "No" , false },
            };
            using var template = new MemoryStream(Properties.Resources.formats);
            using var resultData = await WordDocumentGenerator.GenerateDocument(data, template, CancellationToken.None);

            using var output = File.Create("Success_Formats.docx");
            resultData.CopyTo(output);
        }

        #region Helpers

        class TestObject
        {
            public int Number { get; set; }
            public List<Row> Values { get; set; }
        }
        class Row
        {
            public string Name { get; set; }
            public int Price { get; set; }
            public string Note { get; set; }
        }

        #endregion
    }
}