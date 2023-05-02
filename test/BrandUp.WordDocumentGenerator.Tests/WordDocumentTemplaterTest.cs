using BrandUp.DocumentTemplater.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections;
using System.Reflection;

namespace BrandUp.DocumentTemplater
{
    public class WordDocumentTemplaterTest
    {
        const string TestDirectory = "../Test";

        /// <summary>
        /// Попытка передать на вход устаревший формать файла
        /// </summary>
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

        /// <summary>
        /// Проверка логики выполняемой при команде <see cref="Handling.CommandOutputType.List"/>
        /// </summary>
        [Fact]
        public async void Success_Foreach()
        {
            var testObj = new { o1 = 1, o2 = "abc", o3 = new DateTime(2022, 04, 20, 13, 55, 15) };
            var list = new List<object>
            {
                testObj,
                testObj,
                testObj,
                testObj
            };

            var obj = new
            {
                ObjList = new List<object>
                    {
                        testObj,
                        testObj,
                        testObj,
                        testObj
                    },
                IntArray = new int[] { 1, 2, 3 }

            };
            using var template = new MemoryStream(Properties.Resources._foreach);
            using var resultDocument = await WordDocumentTemplater.GenerateDocument(obj, template, CancellationToken.None);

            #region Asserts

            var elements = GetElementsFromResult(resultDocument);
            try
            {
                var foreachElements = elements.Where(e => e.TagCommand.ToLower().Contains("foreach"));
                Assert.Equal(7, foreachElements.Count()); // 4 для ObjList, 3 для IntArray

                var o1Elements = elements.Where(e => e.TagCommand.ToLower() == "{prop(o1)}");
                Assert.Equal(4, o1Elements.Count());
                Assert.All(o1Elements, e => Assert.Equal(testObj.o1.ToString(), e.Value));

                var o2Elements = elements.Where(e => e.TagCommand.ToLower() == "{prop(o2)}");
                Assert.Equal(4, o2Elements.Count());
                Assert.All(o2Elements, e => Assert.Equal(testObj.o2.ToString(), e.Value));

                var o3Elements = elements.Where(e => e.TagCommand.ToLower() == "{prop(o3)}");
                Assert.Equal(4, o3Elements.Count());
                Assert.All(o3Elements, e => Assert.Equal(testObj.o3.ToString(), e.Value));

                var intElements = elements.Where(e => e.TagCommand.ToLower() == "{prop()}");
                Assert.Equal(3, intElements.Count());
                Assert.True(intElements.Select(e => e.Value).SequenceEqual(obj.IntArray.Select(e => e.ToString())));
            }
            finally
            {
                Save(resultDocument, "Success_Foreach.docx");
            }

            #endregion
        }

        /// <summary>
        /// Проверка логики выполняемой при команде <see cref="Handling.CommandOutputType.Content"/>
        /// </summary>
        [Fact]
        public async void Success_Prop()
        {
            var data = new
            {
                FullDate = new DateTime(2023, 12, 22, 13, 55, 16),
                Span = new TimeSpan(13, 55, 16),
                IntValue = 12,
                DecValue = 123123123123m,
                FloatValue = 0.9210f,
                DoubleValue = double.MaxValue,
                Yes = true,
                No = false
            };

            using var template = new MemoryStream(Properties.Resources.prop);
            using var resultDocument = await WordDocumentTemplater.GenerateDocument(data, template, CancellationToken.None);

            #region Asserts

            var elements = GetElementsFromResult(resultDocument);
            try
            {
                Assert.Equal(8, elements.Count());
                Assert.Collection(elements,
                    e => Assert.Equal(data.FullDate.ToString("dd-MM-yyyy hh-mm-ss"), e.Value), //Потому что даты зависят от культуры
                    e => Assert.Equal("13:55.16", e.Value),
                    e => Assert.Equal("12 градусов", e.Value),
                    e => Assert.Equal("#123123123123#", e.Value),
                    e => Assert.Equal(data.FloatValue.ToString("#.##"), e.Value),
                    e => Assert.Equal(data.DoubleValue.ToString("0.0e+00"), e.Value),
                    e => Assert.Equal("да", e.Value),
                    e => Assert.Equal("Нет", e.Value));
            }
            finally
            {
                Save(resultDocument, "Success_Prop.docx");
            }

            #endregion
        }

        /// <summary>
        /// Проверка сложного шаблона
        /// </summary>
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
                    Name = "ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ \"СЕНДЕР\"",
                    Bank = "АО \"Тинькофф Банк\" МОСКВА",
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
            using var resultDocument = await WordDocumentTemplater.GenerateDocument(data, template, CancellationToken.None);

            #region Assert

            var elements = GetElementsFromResult(resultDocument);
            try
            {
                var tag = "{{prop({0})}}"; //Общий вид команды

                var testDictionary = new Dictionary<string, List<string>>();
                FillTestData(data, testDictionary);

                var ids = elements.Where(e => e.TagCommand.Equals(string.Format(tag, "Id"), StringComparison.OrdinalIgnoreCase)).ToList();
                Assert.Equal(4, ids.Count);
                Assert.True(ids.Select(e => e.Value).Distinct().SequenceEqual(testDictionary["Id"]));

                var dates = elements.Where(e => e.TagCommand.Equals(string.Format(tag, "Date, \"dd.MM.yyyy\""), StringComparison.OrdinalIgnoreCase)).ToList();
                Assert.Equal(2, dates.Count);
                Assert.True(dates.Select(e => e.Value).Distinct().SequenceEqual(testDictionary["Date"]));

                var INNs = elements.Where(e => e.TagCommand.Equals(string.Format(tag, "INN"), StringComparison.OrdinalIgnoreCase)).ToList();
                Assert.True(INNs.Select(e => e.Value).SequenceEqual(testDictionary["INN"]));

                var KPPs = elements.Where(e => e.TagCommand.Equals(string.Format(tag, "KPP"), StringComparison.OrdinalIgnoreCase)).ToList();
                Assert.True(KPPs.Select(e => e.Value).SequenceEqual(testDictionary["KPP"]));

                var Names = elements.Where(e => e.TagCommand.Equals(string.Format(tag, "Name"), StringComparison.OrdinalIgnoreCase)).ToList();
                Assert.Equal(4, Names.Count);
                Assert.True(Names.Select(e => e.Value).SequenceEqual(testDictionary["Name"]));

                var Banks = elements.Where(e => e.TagCommand.Equals(string.Format(tag, "Bank"), StringComparison.OrdinalIgnoreCase)).ToList();
                Assert.True(Banks.Select(e => e.Value).SequenceEqual(testDictionary["Bank"]));

                var BankAccounts = elements.Where(e => e.TagCommand.Equals(string.Format(tag, "BankAccount"), StringComparison.OrdinalIgnoreCase)).ToList();
                Assert.True(BankAccounts.Select(e => e.Value).SequenceEqual(testDictionary["BankAccount"]));

                var BICs = elements.Where(e => e.TagCommand.Equals(string.Format(tag, "BIC"), StringComparison.OrdinalIgnoreCase)).ToList();
                Assert.True(BICs.Select(e => e.Value).SequenceEqual(testDictionary["BIC"]));

                var CorrespondentAccounts = elements.Where(e => e.TagCommand.Equals(string.Format(tag, "CorrespondentAccount"), StringComparison.OrdinalIgnoreCase)).ToList();
                Assert.True(CorrespondentAccounts.Select(e => e.Value).SequenceEqual(testDictionary["CorrespondentAccount"]));

                var Addresses = elements.Where(e => e.TagCommand.Equals(string.Format(tag, "Address"), StringComparison.OrdinalIgnoreCase)).ToList();
                Assert.True(Addresses.Select(e => e.Value).SequenceEqual(testDictionary["Address"]));

                var Emails = elements.Where(e => e.TagCommand.Equals(string.Format(tag, "Email"), StringComparison.OrdinalIgnoreCase)).ToList();
                Assert.True(Emails.Select(e => e.Value).SequenceEqual(testDictionary["Email"]));

                var Links = elements.Where(e => e.TagCommand.Equals(string.Format(tag, "Link"), StringComparison.OrdinalIgnoreCase)).ToList();
                Assert.True(Links.Select(e => e.Value).SequenceEqual(testDictionary["Link"]));

                var Phones = elements.Where(e => e.TagCommand.Equals(string.Format(tag, "Phone"), StringComparison.OrdinalIgnoreCase)).ToList();
                Assert.True(Phones.Select(e => e.Value).SequenceEqual(testDictionary["Phone"]));

                var Quantitys = elements.Where(e => e.TagCommand.Equals(string.Format(tag, "Quantity"), StringComparison.OrdinalIgnoreCase)).ToList();
                Assert.True(Quantitys.Select(e => e.Value).SequenceEqual(testDictionary["Quantity"]));

                var Amounts = elements.Where(e => e.TagCommand.Equals(string.Format(tag, "Amount"), StringComparison.OrdinalIgnoreCase)).ToList();
                Assert.True(Amounts.Select(e => e.Value).SequenceEqual(testDictionary["Amount"]));

                var Prices = elements.Where(e => e.TagCommand.Equals(string.Format(tag, "Price"), StringComparison.OrdinalIgnoreCase)).ToList();
                Assert.True(Prices.Select(e => e.Value).SequenceEqual(testDictionary["Price"]));

                var Taxs = elements.Where(e => e.TagCommand.Equals(string.Format(tag, "Tax, \"##;##;без НДС\""), StringComparison.OrdinalIgnoreCase)).ToList();
                Assert.True(Taxs.Select(e => e.Value).SequenceEqual(new string[] { "без НДС" }));

                var TotalSums = elements.Where(e => e.TagCommand.Equals(string.Format(tag, "TotalSum"), StringComparison.OrdinalIgnoreCase)).ToList();
                Assert.True(TotalSums.Select(e => e.Value).SequenceEqual(testDictionary["TotalSum"]));
            }
            finally
            {
                Save(resultDocument, "Success_Invoice.docx");
            }

            #endregion
        }

        #region Helpers

        static void FillTestData(object data, IDictionary<string, List<string>> dictionary)
        {
            var type = data.GetType();
            var properties = type.GetProperties().ToList();
            foreach (var property in properties)
            {
                var propertyType = property.PropertyType;
                if (propertyType.IsClass && !propertyType.IsAssignableTo(typeof(IEnumerable)))
                {
                    FillTestData(property.GetValue(data), dictionary);
                }
                else if (propertyType == typeof(DateTime))
                {
                    AddToDictionary(property, data, dictionary, "dd.MM.yyyy");
                }
                else if (propertyType != typeof(string) && propertyType.IsAssignableTo(typeof(IEnumerable)))
                {
                    var list = (IEnumerable)property.GetValue(data);
                    foreach (var item in list)
                        FillTestData(item, dictionary);
                }
                else
                {
                    AddToDictionary(property, data, dictionary);
                }
            }
        }

        static void AddToDictionary(PropertyInfo property, object obj, IDictionary<string, List<string>> dictionary, string format = null)
        {
            if (obj == null)
                return;

            var value = property.GetValue(obj);
            if (value == null) return;

            if (dictionary.TryGetValue(property.Name, out List<string> list))
                list.Add(value.ToString(format));
            else dictionary.Add(property.Name, new() { property.GetValue(obj).ToString(format) });
        }

        static void Save(Stream data, string name)
        {
            data.Seek(0, SeekOrigin.Begin);
            using var output = File.Create(Path.Combine(TestDirectory, name));
            data.CopyTo(output);
        }

        static List<ResultData> GetElementsFromResult(Stream stream)
        {
            var dict = new List<ResultData>();
            using (var wordDocument = WordprocessingDocument.Open(stream, false))
            {
                dict = wordDocument.MainDocumentPart.Document.Body.Descendants<SdtElement>()
                    .Where(it => it.SdtProperties.GetFirstChild<Tag>() != null).Select(it => new ResultData(it.SdtProperties.GetFirstChild<Tag>().Val.Value, it.InnerText)).ToList();
            }

            return dict;
        }

        #endregion
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