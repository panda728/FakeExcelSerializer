using FakeExcelSerializer.Serializers;

namespace FakeExcelSerializer.Providers;

public sealed class BuiltinExcelSerializerProvider : IExcelSerializerProvider
{
    public static IExcelSerializerProvider Instance { get; } = new BuiltinExcelSerializerProvider();
    readonly Dictionary<Type, IExcelSerializer> serializers = new()
        {
            { typeof(string), new BuiltinSerializers.StringExcelSerializer() },
            { typeof(Guid), new  BuiltinSerializers.GuidExcelSerializer() },
            { typeof(Enum), new  BuiltinSerializers.EnumExcelSerializer() },
            { typeof(DateTime), new  BuiltinSerializers.DateTimeExcelSerializer() },
            { typeof(DateTimeOffset), new  BuiltinSerializers.DateTimeOffsetExcelSerializer() },
            { typeof(TimeSpan), new  BuiltinSerializers.TimeSpanExcelSerializer() },
            { typeof(Uri), new  BuiltinSerializers.UriExcelSerializer() },
            { typeof(DateOnly), new  BuiltinSerializers.DateOnlyExcelSerializer() },
            { typeof(TimeOnly), new  BuiltinSerializers.TimeOnlyExcelSerializer() },
        };

    public IExcelSerializer<T>? GetSerializer<T>()
    {
        if (serializers.TryGetValue(typeof(T), out var value))
        {
            return (IExcelSerializer<T>)value;
        }
        return null;
    }
}