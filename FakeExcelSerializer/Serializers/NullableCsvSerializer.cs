using System.Buffers;

namespace FakeExcelSerializer.Serializers;

public sealed class NullableExcelSerializer<T> : IExcelSerializer<T?>
    where T : struct
{
    public void Serialize(ref ExcelSerializerWriter writer, T? value, ExcelSerializerOptions options)
    {
        if (value == null)
        {
            writer.WriteEmpty();
            return;
        }
        options.GetRequiredSerializer<T>().Serialize(ref writer, value.Value, options);
    }
}