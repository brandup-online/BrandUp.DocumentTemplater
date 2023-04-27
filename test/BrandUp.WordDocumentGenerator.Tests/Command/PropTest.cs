using BrandUp.DocumentTemplater;

namespace BrandUp.DocxGenerator.Command
{
    public class PropTest : TestBase
    {
        [Fact]
        public async void Success()
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
            using var template = new MemoryStream(Properties.Resources.prop);
            using var resultData = await WordDocumentTemplater.GenerateDocument(data, template, CancellationToken.None);

            Save(resultData, "Success_Prop.docx");
        }
    }
}
