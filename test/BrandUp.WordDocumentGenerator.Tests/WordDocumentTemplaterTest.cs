using BrandUp.DocumentTemplater.Models;

namespace BrandUp.DocumentTemplater
{
    public class WordDocumentTemplaterTest : TestBase
    {
        [Fact]
        public async void WrongFiles()
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

            using var template = new MemoryStream(Properties.Resources.doc);
            await Assert.ThrowsAsync<FileFormatException>(async () =>
            {
                using var resultData = await WordDocumentTemplater.GenerateDocument(data, template, CancellationToken.None);
            });
        }

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
            var controls = GetTagsFromTemplate(template);
            using var resultData = await WordDocumentTemplater.GenerateDocument(data, template, CancellationToken.None);

            var list = GetElementsFromResult(resultData);
            Test(controls, list, WordDocumentTemplater.testPropertyValues);

            resultData.Seek(0, SeekOrigin.Begin);
            Save(resultData, "Success_TestObject.docx");
        }

        [Fact]
        public async void Success_Dictionary()
        {
            var data = new Dictionary<string, object>
            {
                { "Number" , 102321 },
                { "Values" , new List<Models.Row>()
                        {
                            new() { Name = "Tree", Price = 10000, Note = "Big" },
                            new() { Name = "Garland", Price = 5000, Note = "Long" },
                            new() { Name = "Tariff", Price = 700, Note = "30 days" },
                            new() { Name = "Toys", Price = 10, Note = "Round" }
                        }
                }
            };

            using var template = new MemoryStream(Properties.Resources.test);
            var controls = GetTagsFromTemplate(template);
            using var resultData = await WordDocumentTemplater.GenerateDocument(data, template, CancellationToken.None);

            var list = GetElementsFromResult(resultData);
            Test(controls, list, WordDocumentTemplater.testPropertyValues);

            resultData.Seek(0, SeekOrigin.Begin);
            Save(resultData, "Success_Dictionary.docx");
        }

        [Fact]
        public async void Success_Invoice()
        {
            var data = new Invoice
            {
                Id = 1232,
                Date = DateTime.Now,
                OurCompany = new()
                {
                    INN = "212341311",
                    KPP = "540601431",
                    Name = "ОБЩЕСТВО С ОГРАНИЧЕННОЙ\r\nОТВЕТСТВЕННОСТЬЮ \"СЕНДЕР\"",
                    BankName = "АО \"Тинькофф Банк\" МОСКВА",
                    BankAccount = "40702810016541028499",
                    BIC = "033549918",
                    CorrespondentAccount = "1020171514525555974",
                    Address = "Новосибирская обл., г.о. город Новосибирск, г. Новосибирск, ул. Автодорожная, Зд. 15/3, ОФИС 102, 630132",
                    Email = "it@email.ru",
                    Link = "https://sender.ru",
                    Phone = "8 (908) 223-55-12"
                },
                Counterparty = new()
                {
                    Name = "ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ \"ЁЛКА ТРЕНД\"",
                    INN = "5401895257",
                    KPP = "540601121",
                    Address = "Новосибирская обл., г.о. город Новосибирск, г. Новосибирск, ул. Железнодорожная, Зд. 15/2, ПОМЕЩ. 102, 630132"
                },
                Content = new InvoiceContent
                {
                    Products = new List<Product>
                    {
                        new Product
                        {
                            Id = 1,
                            Name = "Тариф \"Лёгкий старт\"",
                            Quantity = 0,
                            Amount = 700,
                            Price = 700,
                        },
                        new Product
                        {
                            Id = 2,
                            Name = "Тариф \"Професиональный\"",
                            Quantity = 0,
                            Amount = 1500,
                            Price = 1500,
                        },
                    },
                    Tax = 0,
                    TotalSum = 2100
                }
            };
            using var template = new MemoryStream(Properties.Resources.Invoice);
            var controls = GetTagsFromTemplate(template);
            using var resultData = await WordDocumentTemplater.GenerateDocument(data, template, CancellationToken.None);

            var list = GetElementsFromResult(resultData);
            Test(controls, list, WordDocumentTemplater.testPropertyValues);

            resultData.Seek(0, SeekOrigin.Begin);
            Save(resultData, "Success_Invoice.docx");
        }

        #region Helpers



        #endregion
    }
}