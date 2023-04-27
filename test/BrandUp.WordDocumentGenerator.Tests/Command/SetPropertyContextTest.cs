using BrandUp.DocumentTemplater;

namespace BrandUp.DocxGenerator.Command
{
    public class SetPropertyContextTest : TestBase
    {
        [Fact]
        public async void Success()
        {
            var data = new
            {
                Top = new
                {
                    Prop = "Раз",
                    Level1 = new
                    {
                        Prop = "Два"
                    },
                    Level2 = new
                    {
                        Level3 = new
                        {
                            Level4 = new
                            {
                                Prop = "Три"
                            }
                        }
                    }
                },

                Bottom = new
                {
                    Level1 = new
                    {
                        Prop = "Четыре"
                    }
                },
            };

            using var template = new MemoryStream(Properties.Resources.setpropertycontext);
            using var resultData = await WordDocumentTemplater.GenerateDocument(data, template, CancellationToken.None);

            Save(resultData, "Success_SetPropertyContext.docx");
        }
    }
}
