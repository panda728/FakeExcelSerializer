using System;
using System.Buffers;
using System.Runtime.CompilerServices;
using System.Text;

namespace FakeExcelSerializer;

public struct ExcelSerializerWriter : IDisposable
{
    readonly ArrayPoolBufferWriter _writer = new();
    readonly ExcelSerializerOptions _options;
    //private const int XF_NORMAL = 0;
    const int XF_WRAP_TEXT = 1;
    const int XF_DATETIME = 2;
    const int XF_DATE = 3;
    const int XF_TIME = 4;

    static readonly byte[] _emptyColumn = Encoding.UTF8.GetBytes("<c></c>");
    static readonly byte[] _colStartBoolean = Encoding.UTF8.GetBytes(@"<c t=""b""><v>");
    static readonly byte[] _colStartNumber = Encoding.UTF8.GetBytes(@"<c t=""n""><v>");
    static readonly byte[] _colStartStringWrap = Encoding.UTF8.GetBytes(@$"<c t=""s"" s=""{XF_WRAP_TEXT}""><v>");
    static readonly byte[] _colStartString = Encoding.UTF8.GetBytes(@$"<c t=""s""><v>");
    static readonly byte[] _colEnd = Encoding.UTF8.GetBytes(@"</v></c>");
    //readonly byte[] _quote;
    //readonly byte[] _quote2;
    //readonly byte[] _delimiter;
    //readonly byte[] _newLine;

    int _index = 0;
    /// <summary>
    /// 文字列の辞書を管理。同じ値は同じIDで出力するため
    /// </summary>
    public Dictionary<string, int> SharedStrings { get; } = new();

    public void Dispose() => _writer.Dispose();

    int _currentDepth;
    //bool _first;

    public ExcelSerializerWriter(ExcelSerializerOptions options)
    {
        _options = options;
        _currentDepth = 0;
        //_first = true;
        //_quote = options.Encoding.GetBytes($"{options.Quote}");
        //_quote2 = options.Encoding.GetBytes($"{options.Quote}{options.Quote}");
        //_delimiter = options.Encoding.GetBytes($"{options.Delimiter}");
        //_newLine = options.Encoding.GetBytes(options.NewLine);
    }

    private void Clear()
    {
        _currentDepth = 0;
        //_first = true;
        _writer.Clear();
    }

    public ReadOnlySpan<byte> AsSpan() => _writer.GetSpan();
    public ReadOnlyMemory<byte> AsMemory() => _writer.GetMemory();
    public long BytesCommitted() => _writer.BytesCommitted;
    public override string ToString() => Encoding.UTF8.GetString(_writer.GetSpan());

    public async Task CopyToAsync(Stream stream)
    {
        if (stream == null)
            throw new ArgumentNullException(nameof(stream));

        await _writer.CopyToAsync(stream);
        Clear();
    }

    public void CopyTo(Stream stream)
    {
        if (stream == null)
            throw new ArgumentNullException(nameof(stream));

        _writer.CopyTo(stream);
        Clear();
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void EnterAndValidate()
    {
        //_first = true;
        _currentDepth++;
        if (_currentDepth >= _options.MaxDepth)
        {
            ThrowReachedMaxDepth(_currentDepth);
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Exit()
    {
        _currentDepth--;
    }

    /// <summary>Write raw bytes.</summary>
    public void WriteRaw(ReadOnlySpan<byte> value) => _writer.Write(value);

    ///// <summary>Write '"'.</summary>
    //public void WriteQuote()
    //{
    //    if (_options.ShouldQuote)
    //        _writer.Write(_quote);
    //}

    /// <summary>Write "\"\"".</summary>
    public int WriteEmpty()
    {
        _writer.Write(_emptyColumn);
        return 0;
    }

    ///// <summary>Write ",".</summary>
    //public void WriteDelimiter()
    //{
    //    if (_first)
    //    {
    //        _first = false;
    //        return;
    //    }
    //    _writer.Write(_delimiter);
    //}

    ///// <summary>Write CRLF.</summary>
    //public void WriteLine() => _writer.Write(_newLine);

    /// <summary>Write string.</summary>
    public int Write(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return WriteEmpty();

        _writer.Write(
            value?.Contains(Environment.NewLine) ?? false
                ? _colStartStringWrap
                : _colStartString
        );

        var index = SharedStrings.TryAdd(value, _index)
            ? _index++
            : SharedStrings[value];

        Encoding.UTF8.GetBytes($"{index}", _writer);

        _writer.Write(_colEnd);
        return value.Length;
    }

    public int WritePrimitive(bool value)
    {
        var s = $"{value}";
        _writer.Write(_colStartBoolean);
        _ = Encoding.UTF8.GetBytes(s, _writer);
        _writer.Write(_colEnd);
        return s.Length;
    }
    public int WritePrimitive(byte value) => WriterNumber($"{value}");
    public int WritePrimitive(sbyte value) => WriterNumber($"{value}");
    public int WritePrimitive(char value) => WriterNumber($"{value}");
    public int WritePrimitive(decimal value) => WriterNumber($"{value}");
    public int WritePrimitive(double value) => WriterNumber($"{value}");
    public int WritePrimitive(float value) => WriterNumber($"{value}");
    public int WritePrimitive(int value) => WriterNumber($"{value}");
    public int WritePrimitive(uint value) => WriterNumber($"{value}");
    public int WritePrimitive(long value) => WriterNumber($"{value}");
    public int WritePrimitive(ulong value) => WriterNumber($"{value}");
    public int WritePrimitive(short value) => WriterNumber($"{value}");
    public int WritePrimitive(ushort value) => WriterNumber($"{value}");
    int WriterNumber(in ReadOnlySpan<char> chars)
    {
        _writer.Write(_colStartNumber);
        _ = Encoding.UTF8.GetBytes(chars, _writer);
        _writer.Write(_colEnd);
        return chars.Length;
    }
    const int LEN_DATE = 10;
    const int LEN_DATETIME = 18;
    const int LEN_TIME = 8;
    public int WriteDateTime(DateTime value)
    {
        var d = value;
        if (d == DateTime.MinValue) WriteEmpty();
        if (d.Hour == 0 && d.Minute == 0 && d.Second == 0)
        {
            Encoding.UTF8.GetBytes(@$"<c t=""d"" s=""{XF_DATE}""><v>{d:yyyy-MM-ddTHH:mm:ss}</v></c>", _writer);
            return LEN_DATE;
        }
        Encoding.UTF8.GetBytes(@$"<c t=""d"" s=""{XF_DATETIME}""><v>{d:yyyy-MM-ddTHH:mm:ss}</v></c>", _writer);
        return LEN_DATETIME;
    }

    public int WriteDateTime(DateOnly value)
    {
        Encoding.UTF8.GetBytes(@$"<c t=""d"" s=""{XF_DATE}""><v>{value:yyyy-MM-ddTHH:mm:ss}</v></c>", _writer);
        return LEN_DATE;
    }

    public int WriteDateTime(TimeOnly value)
    {
        Encoding.UTF8.GetBytes(@$"<c t=""d"" s=""{XF_TIME}""><v>{value:yyyy-MM-ddTHH:mm:ss}</v></c>", _writer);
        return LEN_TIME;
    }

    static void ThrowReachedMaxDepth(int depth)
    {
        throw new InvalidOperationException($"Serializer detects reached max depth:{depth}. Please check the circular reference.");
    }
}