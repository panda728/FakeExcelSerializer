﻿// <auto-generated>
// THIS (.cs) FILE IS GENERATED BY `Serializers/PrimitiveSerializer.tt`. DO NOT CHANGE IT.
// </auto-generated>
#nullable enable
namespace FakeExcelSerializer.Serializers
{
    public sealed class BooleanExcelSerializer : IExcelSerializer<Boolean>
    {
        public void Serialize(ref ExcelSerializerWriter writer, Boolean value, ExcelSerializerOptions options)
        {
            writer.WritePrimitive(value);
        }
    }

    public sealed class ByteExcelSerializer : IExcelSerializer<Byte>
    {
        public void Serialize(ref ExcelSerializerWriter writer, Byte value, ExcelSerializerOptions options)
        {
            writer.WritePrimitive(value);
        }
    }

    public sealed class SByteExcelSerializer : IExcelSerializer<SByte>
    {
        public void Serialize(ref ExcelSerializerWriter writer, SByte value, ExcelSerializerOptions options)
        {
            writer.WritePrimitive(value);
        }
    }

    public sealed class DecimalExcelSerializer : IExcelSerializer<Decimal>
    {
        public void Serialize(ref ExcelSerializerWriter writer, Decimal value, ExcelSerializerOptions options)
        {
            writer.WritePrimitive(value);
        }
    }

    public sealed class DoubleExcelSerializer : IExcelSerializer<Double>
    {
        public void Serialize(ref ExcelSerializerWriter writer, Double value, ExcelSerializerOptions options)
        {
            writer.WritePrimitive(value);
        }
    }

    public sealed class SingleExcelSerializer : IExcelSerializer<Single>
    {
        public void Serialize(ref ExcelSerializerWriter writer, Single value, ExcelSerializerOptions options)
        {
            writer.WritePrimitive(value);
        }
    }

    public sealed class Int32ExcelSerializer : IExcelSerializer<Int32>
    {
        public void Serialize(ref ExcelSerializerWriter writer, Int32 value, ExcelSerializerOptions options)
        {
            writer.WritePrimitive(value);
        }
    }

    public sealed class UInt32ExcelSerializer : IExcelSerializer<UInt32>
    {
        public void Serialize(ref ExcelSerializerWriter writer, UInt32 value, ExcelSerializerOptions options)
        {
            writer.WritePrimitive(value);
        }
    }

    public sealed class Int64ExcelSerializer : IExcelSerializer<Int64>
    {
        public void Serialize(ref ExcelSerializerWriter writer, Int64 value, ExcelSerializerOptions options)
        {
            writer.WritePrimitive(value);
        }
    }

    public sealed class UInt64ExcelSerializer : IExcelSerializer<UInt64>
    {
        public void Serialize(ref ExcelSerializerWriter writer, UInt64 value, ExcelSerializerOptions options)
        {
            writer.WritePrimitive(value);
        }
    }

    public sealed class Int16ExcelSerializer : IExcelSerializer<Int16>
    {
        public void Serialize(ref ExcelSerializerWriter writer, Int16 value, ExcelSerializerOptions options)
        {
            writer.WritePrimitive(value);
        }
    }

    public sealed class UInt16ExcelSerializer : IExcelSerializer<UInt16>
    {
        public void Serialize(ref ExcelSerializerWriter writer, UInt16 value, ExcelSerializerOptions options)
        {
            writer.WritePrimitive(value);
        }
    }

}

namespace FakeExcelSerializer.Providers
{
    public sealed partial class PrimitiveExcelSerializerProvider
    {
        internal partial void InitPrimitives()
        {
            serializers[typeof(System.Boolean)] = new FakeExcelSerializer.Serializers.BooleanExcelSerializer();
            serializers[typeof(System.Byte)] = new FakeExcelSerializer.Serializers.ByteExcelSerializer();
            serializers[typeof(System.SByte)] = new FakeExcelSerializer.Serializers.SByteExcelSerializer();
            serializers[typeof(System.Decimal)] = new FakeExcelSerializer.Serializers.DecimalExcelSerializer();
            serializers[typeof(System.Double)] = new FakeExcelSerializer.Serializers.DoubleExcelSerializer();
            serializers[typeof(System.Single)] = new FakeExcelSerializer.Serializers.SingleExcelSerializer();
            serializers[typeof(System.Int32)] = new FakeExcelSerializer.Serializers.Int32ExcelSerializer();
            serializers[typeof(System.UInt32)] = new FakeExcelSerializer.Serializers.UInt32ExcelSerializer();
            serializers[typeof(System.Int64)] = new FakeExcelSerializer.Serializers.Int64ExcelSerializer();
            serializers[typeof(System.UInt64)] = new FakeExcelSerializer.Serializers.UInt64ExcelSerializer();
            serializers[typeof(System.Int16)] = new FakeExcelSerializer.Serializers.Int16ExcelSerializer();
            serializers[typeof(System.UInt16)] = new FakeExcelSerializer.Serializers.UInt16ExcelSerializer();
        }
    }
}