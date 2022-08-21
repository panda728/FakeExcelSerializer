using FluentAssertions;
using System.Diagnostics;
using System.Text;

namespace FakeExcelSerializer.Tests
{
    public class StandardTest
    {
        [Fact]
        public void Serializer_string()
        {
            var options = ExcelSerializerOptions.Default;
            var serializer = options.GetSerializer<string>();
            Assert.NotNull(serializer);
            if (serializer == null) return;

            var writer = new ExcelSerializerWriter(options);
            serializer.Serialize(ref writer, "column1", options);
            serializer.Serialize(ref writer, "column2", options);
            serializer.Serialize(ref writer, "column1", options);

            Assert.Equal(2, writer.SharedStrings.Count);

            var columnXml = writer.ToString();
            var sharedString1 = writer.SharedStrings.First().Key;
            var sharedString2 = writer.SharedStrings.Skip(1).First().Key;
            writer.Dispose();

            columnXml.Should().Be("<c t=\"s\"><v>0</v></c><c t=\"s\"><v>1</v></c><c t=\"s\"><v>0</v></c>");
            sharedString1.Should().Be("column1");
            sharedString2.Should().Be("column2");
        }
    }
}