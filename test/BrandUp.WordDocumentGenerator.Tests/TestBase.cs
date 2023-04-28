using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BrandUp.DocumentTemplater
{
    public abstract class TestBase
    {
        const string TestDirectory = "../Test";
        public TestBase()
        {
            var info = Directory.CreateDirectory(TestDirectory);
        }

        protected virtual void Save(Stream data, string name)
        {
            using var output = File.Create(Path.Combine(TestDirectory, name));
            data.CopyTo(output);
        }

        protected static IEnumerable<string> GetTagsFromTemplate(Stream stream)
        {
            var list = new List<string>();

            using (var wordDocument = WordprocessingDocument.Open(stream, false))
            {
                list = wordDocument.MainDocumentPart.Document.Body.Descendants<SdtElement>()
                    .Where(it => it.SdtProperties.GetFirstChild<Tag>() != null)
                    .Select(it => it.SdtProperties.GetFirstChild<Tag>().Val.Value).ToList();
            }

            stream.Seek(0, SeekOrigin.Begin);
            return list;
        }

        protected static List<ResultData> GetElementsFromResult(Stream stream)
        {
            var dict = new List<ResultData>();
            using (var wordDocument = WordprocessingDocument.Open(stream, false))
            {
                dict = wordDocument.MainDocumentPart.Document.Body.Descendants<SdtElement>()
                    .Where(it => it.SdtProperties.GetFirstChild<Tag>() != null).Select(it => new ResultData(it.SdtProperties.GetFirstChild<Tag>().Val.Value, it.InnerText)).ToList();
            }

            return dict;
        }

        /// <summary>
        /// Проверяем что все данные в результирующем документе установленны правильно. Корректность обработки комманд проверяется в тестах команд.
        /// </summary>
        /// <param name="templateCommands">Комманды которые были установленны в шаблоне</param>
        /// <param name="resultData">Пара тег и значение которое было записанно в документ</param>
        /// <param name="processedData">Данные которые были получены после обработки комманды</param>
        protected static void Test(IEnumerable<string> templateCommands, IEnumerable<ResultData> resultData, IDictionary<string, List<string>> processedData)
        {
            foreach (var command in templateCommands.Where(c => c.StartsWith("{prop")) /*все места в которые должны быть записаны данные */)
            {
                var resultValues = resultData.Where(t => t.TagCommand == command).ToList(); // Данные которые записанны в результирующем документе. список потому что foreach
                Assert.NotEmpty(resultValues);
                if (processedData.TryGetValue(command, out var proccessValues))
                {
                    // Все обработанные данные записанны в результирующий документ
                    Assert.True(resultValues.Select(it => it.Value).SequenceEqual(proccessValues));
                }
                Assert.NotNull(proccessValues);
            }
        }
    }

    public class ResultData
    {
        public string TagCommand { get; set; }
        public string Value { get; set; }

        public ResultData(string tagCommand, string value)
        {
            TagCommand = tagCommand ?? throw new ArgumentNullException(nameof(tagCommand));
            Value = value ?? throw new ArgumentNullException(nameof(value));
        }
    }
}
