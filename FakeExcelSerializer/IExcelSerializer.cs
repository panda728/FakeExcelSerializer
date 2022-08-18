using System.Buffers;

namespace FakeExcelSerializer;

public interface IExcelSerializer { }

public interface IExcelSerializer<T> : IExcelSerializer
{
    void Serialize(ref ExcelSerializerWriter writer, T value, ExcelSerializerOptions options);
}
