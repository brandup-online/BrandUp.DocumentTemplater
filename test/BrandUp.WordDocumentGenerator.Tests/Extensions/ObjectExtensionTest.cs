namespace BrandUp.DocumentTemplater
{
    public class ObjectExtensionTest
    {
        /// <summary>
        /// Числовые типы помимо int (long, short, byte, …) должны форматироваться
        /// как IFormattable, а не уходить в string.Format и терять значение.
        /// </summary>
        [Fact]
        public void FormatsAllNumericTypesViaIFormattable()
        {
            long l = 123456789L;
            short s = 1234;
            byte b = 200;
            uint u = 4000000000;

            Assert.Equal(l.ToString("N0"), ObjectExtension.ToString(l, "N0"));
            Assert.Equal(s.ToString("D6"), ObjectExtension.ToString(s, "D6"));
            Assert.Equal(b.ToString("X2"), ObjectExtension.ToString(b, "X2"));
            Assert.Equal(u.ToString("N0"), ObjectExtension.ToString(u, "N0"));
        }

        [Fact]
        public void FormatsGuid()
        {
            var guid = Guid.NewGuid();
            Assert.Equal(guid.ToString("N"), ObjectExtension.ToString(guid, "N"));
        }

        [Theory]
        [InlineData(true, "b", "да")]
        [InlineData(false, "b", "нет")]
        [InlineData(true, "B", "Да")]
        [InlineData(false, "B", "Нет")]
        public void FormatsBoolean(bool value, string format, string expected)
        {
            Assert.Equal(expected, ObjectExtension.ToString(value, format));
        }

        [Fact]
        public void EmptyFormatFallsBackToToString()
        {
            Assert.Equal("42", ObjectExtension.ToString(42L, ""));
            Assert.Equal("42", ObjectExtension.ToString(42L, null));
        }
    }
}
