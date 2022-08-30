using System.Buffers;
using System.Collections.Concurrent;
using System.Linq.Expressions;

namespace FakeExcelSerializer.Serializers;

internal class ObjectFallbackExcelSerializer : IExcelSerializer<object>
{
    delegate void SerializeDelegate(ref ExcelSerializerWriter writer, object value, ExcelSerializerOptions options);
    static readonly ConcurrentDictionary<Type, SerializeDelegate> nongenericSerializers = new();
    static readonly Func<Type, SerializeDelegate> factory = CompileSerializeDelegate;

    public void WriteTitle(ref ExcelSerializerWriter writer, object value, ExcelSerializerOptions options, string name = "")
        => writer.Write(name);

    public void Serialize(ref ExcelSerializerWriter writer, object value, ExcelSerializerOptions options)
    {
        if (value == null)
        {
            writer.WriteEmpty();
            return;
        }

        var type = value.GetType();
        if (type == typeof(object))
        {
            writer.WriteEmpty();
            return;
        }

        var serializer = nongenericSerializers.GetOrAdd(type, factory);

        serializer.Invoke(ref writer, value, options);
    }

    static SerializeDelegate CompileSerializeDelegate(Type type)
    {
        // Serialize(ref ExcelSerializerWriter writer, object value, ExcelSerializerOptions options)
        //   options.GetRequiredSerializer<T>().Serialize(ref writer, (T)value, options)

        var writer = Expression.Parameter(typeof(ExcelSerializerWriter).MakeByRefType());
        var value = Expression.Parameter(typeof(object));
        var options = Expression.Parameter(typeof(ExcelSerializerOptions));

        var getRequiredSerializer = typeof(ExcelSerializerOptions).GetMethod("GetRequiredSerializer", 1, Type.EmptyTypes)!.MakeGenericMethod(type);
        var serialize = typeof(IExcelSerializer<>).MakeGenericType(type).GetMethod("Serialize")!;

        var body = Expression.Call(
            Expression.Call(options, getRequiredSerializer),
            serialize,
            writer,
            Expression.Convert(value, type),
            options);

        var lambda = Expression.Lambda<SerializeDelegate>(body, writer, value, options);
        return lambda.Compile();
    }
}