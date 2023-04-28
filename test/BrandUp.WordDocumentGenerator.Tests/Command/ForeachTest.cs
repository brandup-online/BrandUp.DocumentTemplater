using BrandUp.DocumentTemplater;

namespace BrandUp.DocxGenerator.Command
{
    public class ForeachTest : TestBase
    {
        [Fact]
        public async void Success()
        {
            var data = new
            {
                Obj = new List<object>
                {
                    new
                    {

                        o1 = 15,
                        o2 = "info1",
                        o3 = new DateTime(2023, 04, 06, 12, 11, 09)
                    },
                    new
                    {

                        o1 = 16,
                        o2 = "info2",
                        o3 = new DateTime(2022, 11, 06, 20, 10, 19)
                    },

                    new
                    {

                        o1 = 17,
                        o2 = "info3",
                        o3 = new DateTime(2023, 04, 07, 13, 10, 08)
                    }
                },
                Numbers = new int[]
                {
                    1,2,3
                }
            };

            using var template = new MemoryStream(Properties.Resources._foreach);
            using var resultData = await WordDocumentTemplater.GenerateDocument(data, template, CancellationToken.None);

            Save(resultData, "Success_Foreach.docx");
        }
    }
}
