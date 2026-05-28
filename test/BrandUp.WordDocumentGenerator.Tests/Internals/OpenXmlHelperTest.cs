using BrandUp.DocumentTemplater.Internals;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BrandUp.DocumentTemplater.Internals
{
    public class OpenXmlHelperTest
    {
        /// <summary>
        /// Дублирующиеся id элементов управления (возникают при клонировании
        /// в foreach) должны стать уникальными.
        /// </summary>
        [Fact]
        public void SetUniqueContentControlIds_MakesDuplicatesUnique()
        {
            var body = new Body(
                NewSdtBlock(5),
                NewSdtBlock(5),
                NewSdtBlock(5),
                NewSdtBlock(7));

            OpenXmlHelper.SetUniquecontentControlIds(body, []);

            var ids = body.Descendants<SdtId>().Select(x => x.Val.Value).ToList();
            Assert.Equal(4, ids.Count);
            Assert.Equal(4, ids.Distinct().Count());
        }

        /// <summary>
        /// Уже уникальные id не должны меняться.
        /// </summary>
        [Fact]
        public void SetUniqueContentControlIds_KeepsAlreadyUniqueIds()
        {
            var body = new Body(
                NewSdtBlock(1),
                NewSdtBlock(2),
                NewSdtBlock(3));

            OpenXmlHelper.SetUniquecontentControlIds(body, []);

            var ids = body.Descendants<SdtId>().Select(x => x.Val.Value).ToList();
            Assert.Equal(new[] { 1, 2, 3 }, ids);
        }

        /// <summary>
        /// Элементы управления без заданного id (Val == null) пропускаются
        /// без исключений и не влияют на уже занятые id.
        /// </summary>
        [Fact]
        public void SetUniqueContentControlIds_SkipsNullVal()
        {
            var withNull = new SdtBlock(new SdtProperties(new SdtId()));
            var withValue = NewSdtBlock(5);
            var body = new Body(withNull, withValue);

            var exception = Record.Exception(() => OpenXmlHelper.SetUniquecontentControlIds(body, []));

            Assert.Null(exception);
            var ids = body.Descendants<SdtId>().ToList();
            Assert.Null(ids[0].Val);
            Assert.Equal(5, ids[1].Val.Value);
        }

        static SdtBlock NewSdtBlock(int id)
            => new(new SdtProperties(new SdtId { Val = id }));
    }
}
