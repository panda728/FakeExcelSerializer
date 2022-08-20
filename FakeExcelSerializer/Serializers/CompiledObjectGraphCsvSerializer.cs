using System.Buffers;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.Serialization;

namespace FakeExcelSerializer.Serializers;

internal sealed class CompiledObjectGraphExcelSerializer<T> : IExcelSerializer<T>
{
    delegate void SerializeMethod(ref ExcelSerializerWriter writer, IExcelSerializer?[]? alternateSerializers, T value, ExcelSerializerOptions options);

    static readonly string[] names;
    static readonly IExcelSerializer?[]? alternateSerializers;
    static readonly SerializeMethod serialize;
    static readonly bool isReferenceType;

    static CompiledObjectGraphExcelSerializer()
    {
        isReferenceType = !typeof(T).IsValueType;

        var props = typeof(T).GetProperties();
        var fields = typeof(T).GetFields();
        var members = props.Cast<MemberInfo>().Concat(fields)
            .Where(x => x.GetCustomAttribute<IgnoreExcelSerializeAttribute>() == null)
            .Select((x, i) => new SerializableMemberInfo(x, i))
            .OrderBy(x => x.Order)
            .ThenBy(x => x.Name)
            .ToArray();

        names = members.Select(x => x.Name).ToArray();
        if (members.Any(x => x.ExcelSerializer != null))
        {
            alternateSerializers = members.Select(x => x.ExcelSerializer).ToArray();
        }
        serialize = CompileSerializer(typeof(T), members);
    }

    public static string[] Names => names;

    public void Serialize(ref ExcelSerializerWriter writer, T value, ExcelSerializerOptions options)
    {
        if (isReferenceType)
        {
            if (value == null)
            {
                writer.WriteEmpty();
                return;
            }
        }

        writer.EnterAndValidate();
        serialize(ref writer, alternateSerializers, value, options);
        writer.Exit();
    }

    static SerializeMethod CompileSerializer(Type valueType, SerializableMemberInfo[] memberInfos)
    {
        // SerializeMethod(ref ExcelSerializerWriter writer, IExcelSerializer[]? alternateSerializers, T value, ExcelSerializerOptions options)
        // foreach(members)
        //   if (value.Foo != null) // reference type || nullable type
        //     options.GetRequiredSerializer<T>() || ((ICsvSerialzier<T>)alternateSerializers[0] .Serialize(writer, value.Foo, options)
        //   else
        //     writer.WriteEmpty();

        var argWriterRef = Expression.Parameter(typeof(ExcelSerializerWriter).MakeByRefType());
        var argAlternateSerializers = Expression.Parameter(typeof(IExcelSerializer[]));
        var argValue = Expression.Parameter(valueType);
        var argOptions = Expression.Parameter(typeof(ExcelSerializerOptions));
        var foreachBodies = new List<Expression>();

        var i = 0;
        foreach (var memberInfo in memberInfos)
        {
            Expression serializer;
            if (memberInfo.ExcelSerializer == null)
            {
                serializer = Expression.Call(argOptions, ReflectionInfos.ExcelSerializerOptions_GetRequiredSerializer(memberInfo.MemberType));
            }
            else
            {
                serializer = Expression.Convert(
                    Expression.ArrayIndex(argAlternateSerializers, Expression.Constant(i, typeof(int))),
                    typeof(IExcelSerializer<>).MakeGenericType(memberInfo.MemberType));
            }

            var callSerializer = Expression.Call(serializer, ReflectionInfos.IExcelSerializer_Serialize(memberInfo.MemberType), argWriterRef, memberInfo.GetMemberExpression(argValue), argOptions);
            if (!memberInfo.MemberType.IsValueType || memberInfo.MemberType.IsNullable())
            {
                var nullExpr = Expression.Constant(null, memberInfo.MemberType);
                var ifBody = Expression.IfThenElse(
                    Expression.NotEqual(memberInfo.GetMemberExpression(argValue), nullExpr),
                    callSerializer,
                    Expression.Call(argWriterRef, ReflectionInfos.CsvWriter_Empty)
                );
                foreachBodies.Add(ifBody);
            }
            else
            {
                foreachBodies.Add(callSerializer);
            }
            i++;
        }

        var body = Expression.Block(foreachBodies);
        var lambda = Expression.Lambda<SerializeMethod>(body, argWriterRef, argAlternateSerializers, argValue, argOptions);
        return lambda.Compile();
    }

    internal static class ReflectionInfos
    {
        internal static MethodInfo CsvWriter_Empty { get; } = typeof(ExcelSerializerWriter).GetMethod("WriteEmpty")!;
        internal static MethodInfo ExcelSerializerOptions_GetRequiredSerializer(Type type) => typeof(ExcelSerializerOptions).GetMethod("GetRequiredSerializer", 1, Type.EmptyTypes)!.MakeGenericMethod(type);
        internal static MethodInfo IExcelSerializer_Serialize(Type type) => typeof(IExcelSerializer<>).MakeGenericType(type).GetMethod("Serialize")!;
    }
}

internal sealed class SerializableMemberInfo
{
    public string Name { get; }
    public int Order { get; }
    public IExcelSerializer? ExcelSerializer { get; }
    public Type MemberType { get; }
    public MemberInfo MemberInfo { get; }

    public SerializableMemberInfo(MemberInfo member, int i)
    {
        var dataMember = member.GetCustomAttribute<DataMemberAttribute>();

        MemberInfo = member;
        Name = dataMember?.Name ?? member.Name;
        Order = dataMember?.Order ?? i;

        MemberType = member switch
        {
            PropertyInfo pi => pi.PropertyType,
            FieldInfo fi => fi.FieldType,
            _ => throw new InvalidOperationException()
        };

        var serializerAttr = member.GetCustomAttribute<ExcelSerializerAttribute>();
        if (serializerAttr != null)
        {
            serializerAttr.Validate(MemberType);
            ExcelSerializer = (IExcelSerializer?)Activator.CreateInstance(serializerAttr.Type);
        }
    }

    public MemberExpression GetMemberExpression(Expression expression)
    {
        if (MemberInfo is FieldInfo fi)
        {
            return Expression.Field(expression, fi);
        }
        else if (MemberInfo is PropertyInfo pi)
        {
            return Expression.Property(expression, pi);
        }
        throw new InvalidOperationException();
    }
}