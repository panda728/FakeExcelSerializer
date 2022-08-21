using System.Buffers;
using System.Text;

namespace FakeExcelSerializer.Serializers;

internal class BuiltinSerializers
{
    public sealed class StringExcelSerializer : IExcelSerializer<string?>
    {
        public void Serialize(ref ExcelSerializerWriter writer, string? value, ExcelSerializerOptions options)
            => writer.Write(value);
    }

    public sealed class CharExcelSerializer : IExcelSerializer<char>
    {
        public void Serialize(ref ExcelSerializerWriter writer, char value, ExcelSerializerOptions options)
            => writer.Write(value);
    }

    public sealed class GuidExcelSerializer : IExcelSerializer<Guid>
    {
        public void Serialize(ref ExcelSerializerWriter writer, Guid value, ExcelSerializerOptions options)
            => writer.Write($"{value}");
    }

    public sealed class EnumExcelSerializer : IExcelSerializer<Enum>
    {
        public void Serialize(ref ExcelSerializerWriter writer, Enum value, ExcelSerializerOptions options)
            => writer.Write($"{value}");
    }

    public sealed class DateTimeExcelSerializer : IExcelSerializer<DateTime>
    {
        public void Serialize(ref ExcelSerializerWriter writer, DateTime value, ExcelSerializerOptions options)
            => writer.WriteDateTime(value);
    }

    public sealed class DateTimeOffsetExcelSerializer : IExcelSerializer<DateTimeOffset>
    {
        public void Serialize(ref ExcelSerializerWriter writer, DateTimeOffset value, ExcelSerializerOptions options)
            => writer.Write(value.ToString(options.CultureInfo));
    }

    public sealed class TimeSpanExcelSerializer : IExcelSerializer<TimeSpan>
    {
        public void Serialize(ref ExcelSerializerWriter writer, TimeSpan value, ExcelSerializerOptions options)
            => writer.Write(value.ToString());
    }

    public sealed class UriExcelSerializer : IExcelSerializer<Uri>
    {
        public void Serialize(ref ExcelSerializerWriter writer, Uri value, ExcelSerializerOptions options)
        {
            if (value == null)
            {
                writer.WriteEmpty();
                return;
            }

            writer.Write($"{value}");
        }
    }

    public sealed class DateOnlyExcelSerializer : IExcelSerializer<DateOnly>
    {
        public void Serialize(ref ExcelSerializerWriter writer, DateOnly value, ExcelSerializerOptions options)
            => writer.WriteDateTime(value);
    }

    public sealed class TimeOnlyExcelSerializer : IExcelSerializer<TimeOnly>
    {
        public void Serialize(ref ExcelSerializerWriter writer, TimeOnly value, ExcelSerializerOptions options)
            => writer.WriteDateTime(value);
    }
}
