namespace BrandUp.DocxGenerator
{
    public abstract class TestBase
    {
        const string TestDirectory = "../Test";
        public TestBase()
        {
            var info = Directory.CreateDirectory(TestDirectory);
        }

        public virtual void Save(Stream data, string name)
        {
            using var output = File.Create(Path.Combine(TestDirectory, name));
            data.CopyTo(output);
        }
    }
}
