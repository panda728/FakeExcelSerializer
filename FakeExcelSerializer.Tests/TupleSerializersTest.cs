using FluentAssertions;
using Xunit.Sdk;

namespace FakeExcelSerializer.Tests
{
    public class TupleSerializersTest
    {
        void RunStringColumnTest<T>(
            T value1, T value2,
            string value1ShouldBe, string value2ShouldBe,
            ExcelSerializerOptions option)
        {
            var serializer = option.GetSerializer<T>();
            Assert.NotNull(serializer);
            if (serializer == null) return;
            var writer = new ExcelSerializerWriter(option);
            try
            {
                serializer.Serialize(ref writer, value1, option);
                serializer.Serialize(ref writer, value2, option);
                serializer.Serialize(ref writer, value1, option);

                Assert.Equal(2, writer.SharedStrings.Count);

                var columnXml = writer.ToString();
                var sharedString1 = writer.SharedStrings.First().Key;
                var sharedString2 = writer.SharedStrings.Skip(1).First().Key;

                columnXml.Should().Be("<c t=\"s\"><v>0</v></c><c t=\"s\"><v>1</v></c><c t=\"s\"><v>0</v></c>");
                sharedString1.Should().Be(value1ShouldBe);
                sharedString2.Should().Be(value2ShouldBe);
            }
            catch
            {
                throw;
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void Serializer_tuple1()
        {
            var t = Tuple.Create(1);
            RunTest(t, "<c t=\"n\"><v>1</v></c>", ExcelSerializerOptions.Default);
        }

        [Fact]
        public void Serializer_tuple2()
        {
            var t = Tuple.Create(1,2);
            RunTest(t, "<c t=\"n\"><v>1</v></c><c t=\"n\"><v>2</v></c>", ExcelSerializerOptions.Default);
        }

        void RunTest<T>(
            T value1, string value1ShouldBe,
            ExcelSerializerOptions option)
        {
            var serializer = option.GetSerializer<T>();
            Assert.NotNull(serializer);
            if (serializer == null) return;
            var writer = new ExcelSerializerWriter(option);
            try
            {
                serializer.Serialize(ref writer, value1, option);
                Assert.Empty(writer.SharedStrings);
                writer.ToString().Should().Be(value1ShouldBe);
            }
            catch
            {
                throw;
            }
            finally
            {
                writer.Dispose();
            }
        }
    }
}
