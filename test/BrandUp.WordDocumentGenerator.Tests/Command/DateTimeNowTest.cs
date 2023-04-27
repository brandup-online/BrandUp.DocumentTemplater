using BrandUp.DocumentTemplater;

namespace BrandUp.DocxGenerator.Command
{
    public class DateTimeNowTest : TestBase
    {
        [Fact]
        public async void Success()
        {
            using var template = new MemoryStream(Properties.Resources.datetimenow);
            using var resultData = await WordDocumentTemplater.GenerateDocument(new { }, template, CancellationToken.None);

            Save(resultData, "Success_DateTimeNow.docx");
        }
    }
}
