namespace BrandUp.DocumentTemplater.Models
{
    class TestObject
    {
        public int Number { get; set; }
        public List<Row> Values { get; set; }
    }
    class Row
    {
        public string Name { get; set; }
        public int Price { get; set; }
        public string Note { get; set; }
    }
}
