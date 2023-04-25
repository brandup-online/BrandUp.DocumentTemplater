namespace BrandUp.DocxGenerator
{
    public class WordDocumentGeneratorTest
    {
        [Fact]
        public void Success_TestObject()
        {
            var data = new TestObject
            {
                Number = 1,
                Values = new()
                {
                    new() { Name = "Tree", Price = 10000, Note = "Big" },
                    new() { Name = "Garland", Price = 5000, Note = "Long" }
                }
            };
            using var resultData = WordDocumentGenerator.GenerateDocument(data, Properties.Resources.test);

            using var output = File.Create("Success_TestObject.docx");
            resultData.CopyTo(output);
        }

        class TestObject
        {
            public int Number { get; set; }
            public List<TestInner> Values { get; set; }
        }
        class TestInner
        {
            public string Name { get; set; }
            public int Price { get; set; }
            public string Note { get; set; }
        }
    }
}