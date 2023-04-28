using BrandUp.DocumentTemplater.Handling;

namespace BrandUp.DocumentTemplater.Commands
{
    public class ForeachTest : TestBase
    {
        /// <summary>
        /// Тест команды Foreach без параметров
        /// </summary>
        [Fact]
        public void SuccessParameterless()
        {
            var data = new int[]
               {
                   1,2,3
               };

            var command = new Foreach();
            var result = command.Execute(new() { }, data);

            Assert.Same(data, result.DataContext);
            Assert.Null(result.OutputContent);
            Assert.NotNull(result.OutputList);
            Assert.Collection(result.OutputList, x => Assert.Equal(1, x), x => Assert.Equal(2, x), x => Assert.Equal(3, x));
            Assert.Equal(CommandOutputType.List, result.OutputType);
        }

        /// <summary>
        /// Тест команды Foreach с параметром "Obj"
        /// </summary>
        [Fact]
        public void SuccessObj()
        {
            var data = new
            {
                Obj = new List<TestObj>
                {
                    new()
                    {

                        Id = 15,
                        Name = "info1",
                        Date = new DateTime(2023, 04, 06, 12, 11, 09)
                    },
                    new()
                    {

                        Id = 16,
                        Name = "info2",
                        Date = new DateTime(2022, 11, 06, 20, 10, 19)
                    },
                    new()
                    {

                        Id = 17,
                        Name = "info3",
                        Date = new DateTime(2023, 04, 07, 13, 10, 08)
                    }
                },
                Numbers = new int[]
                {
                    1,2,3
                }
            };


            var command = new Foreach();
            var result = command.Execute(new() { "Obj" }, data);

            Assert.Same(data, result.DataContext);
            Assert.Null(result.OutputContent);
            Assert.NotNull(result.OutputList);
            Assert.Collection(data.Obj,
                o =>
            {
                Assert.Equal(15, o.Id);
                Assert.Equal("info1", o.Name);
                Assert.Equal(new DateTime(2023, 04, 06, 12, 11, 09), o.Date);
            }, o =>
            {
                Assert.Equal(16, o.Id);
                Assert.Equal("info2", o.Name);
                Assert.Equal(new DateTime(2022, 11, 06, 20, 10, 19), o.Date);
            }, o =>
            {
                Assert.Equal(17, o.Id);
                Assert.Equal("info3", o.Name);
                Assert.Equal(new DateTime(2023, 04, 07, 13, 10, 08), o.Date);
            });
            Assert.Equal(CommandOutputType.List, result.OutputType);
        }

        class TestObj
        {
            public int Id { get; set; }
            public string Name { get; set; }
            public DateTime Date { get; set; }
        }
    }
}
