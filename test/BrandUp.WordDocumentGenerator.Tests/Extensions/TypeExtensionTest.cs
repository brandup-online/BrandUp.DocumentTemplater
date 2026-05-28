using BrandUp.DocumentTemplater.Exceptions;

namespace BrandUp.DocumentTemplater
{
    public class TypeExtensionTest
    {
        class Node
        {
            public string Name { get; set; }
            public Node Inner { get; set; }
            public object Empty { get; set; }
        }

        [Fact]
        public void Object_SimpleProperty()
        {
            var obj = new Node { Name = "root" };

            Assert.Equal("root", obj.GetPropertyValue("Name"));
        }

        [Fact]
        public void Object_NestedPath()
        {
            var obj = new Node { Inner = new Node { Name = "deep" } };

            Assert.Equal("deep", obj.GetPropertyValue("Inner.Name"));
        }

        [Fact]
        public void Dictionary_KeyLookup()
        {
            var dict = new Dictionary<string, object> { ["k"] = 123 };

            Assert.Equal(123, dict.GetPropertyValue("k"));
        }

        [Fact]
        public void Object_MissingProperty_Throws()
        {
            Assert.Throws<InvalidPropertyNameException>(() => new Node().GetPropertyValue("Nope"));
        }

        [Fact]
        public void Dictionary_MissingKey_Throws()
        {
            var dict = new Dictionary<string, object> { ["k"] = 1 };

            Assert.Throws<InvalidPropertyNameException>(() => dict.GetPropertyValue("missing"));
        }

        [Fact]
        public void Object_NullProperty_Throws()
        {
            var obj = new Node { Empty = null };

            Assert.Throws<ContextValueNullException>(() => obj.GetPropertyValue("Empty"));
        }

        [Fact]
        public void Object_NullValueInPath_Throws()
        {
            var obj = new Node { Inner = null };

            Assert.Throws<ContextValueNullException>(() => obj.GetPropertyValue("Inner.Name"));
        }

        [Fact]
        public void NullObject_Throws()
        {
            object obj = null;

            Assert.Throws<ArgumentNullException>(() => obj.GetPropertyValue("x"));
        }

        [Fact]
        public void NullPropertyName_Throws()
        {
            Assert.Throws<ArgumentNullException>(() => new Node().GetPropertyValue(null));
        }
    }
}
