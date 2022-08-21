using System.Buffers;
using System.Runtime.CompilerServices;
using System.Text;

namespace FakeExcelSerializer;

public class ExcelSerializerWriter : IDisposable
{
    //const int XF_NORMAL = 0;
    const int XF_WRAP_TEXT = 1;
    const int XF_DATETIME = 2;
    const int XF_DATE = 3;
    const int XF_TIME = 4;

    const int LEN_DATE = 10;
    const int LEN_DATETIME = 18;
    const int LEN_TIME = 8;

    static readonly byte[] _emptyColumn = Encoding.UTF8.GetBytes("<c></c>");
    static readonly byte[] _colStartBoolean = Encoding.UTF8.GetBytes(@"<c t=""b""><v>");
    static readonly byte[] _colStartNumber = Encoding.UTF8.GetBytes(@"<c t=""n""><v>");
    static readonly byte[] _colStartStringWrap = Encoding.UTF8.GetBytes(@$"<c t=""s"" s=""{XF_WRAP_TEXT}""><v>");
    static readonly byte[] _colStartString = Encoding.UTF8.GetBytes(@$"<c t=""s""><v>");
    static readonly byte[] _colEnd = Encoding.UTF8.GetBytes(@"</v></c>");

    readonly ArrayPoolBufferWriter _writer = new();
    readonly ExcelSerializerOptions _options;

    bool _countingCharLength;

    int _columnIndex = 0;
    int _currentDepth = 0;
    int _stringIndex = 0;

    public ExcelSerializerWriter(ExcelSerializerOptions options)
    {
        _options = options;
        _currentDepth = 0;
        _countingCharLength = options.AutoFitColumns;
    }

    public void Dispose()
    {
        _writer.Dispose();
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// Maintain a dictionary of strings. Output the same value with the same ID.
    /// </summary>
    public Dictionary<string, int> SharedStrings { get; } = new();
    public ReadOnlySpan<byte> AsSpan() => _writer.OutputAsSpan;
    public ReadOnlyMemory<byte> AsMemory() => _writer.OutputAsMemory;
    public long BytesCommitted() => _writer.BytesCommitted;
    public override string ToString() => Encoding.UTF8.GetString(_writer.OutputAsSpan);

    /// <summary>
    /// Tally the maximum number of characters per column. For automatic column width adjustment
    /// </summary>
    public Dictionary<int, int> ColumnMaxLength { get; } = new();
    public void StopCountingCharLength() => _countingCharLength = false;

    public void Clear()
    {
        _columnIndex = 0;
        _currentDepth = 0;
        _writer.Clear();
    }

    /// <summary>Writes a value to the Stream</summary>
    /// <remarks>Perform one line at a time.</remarks>
    public async Task CopyToAsync(Stream stream)
    {
        if (stream == null)
            throw new ArgumentNullException(nameof(stream));
        await _writer.CopyToAsync(stream);
        Clear();
    }

    /// <summary>Writes a value to the Stream</summary>
    /// <remarks>Perform one line at a time.</remarks>
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
            ThrowReachedMaxDepth(_currentDepth);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Exit()
    {
        _currentDepth--;
    }

    public void WriteRaw(ReadOnlySpan<byte> value) => _writer.Write(value);

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public int WriteEmpty()
    {
        _columnIndex++;
        _writer.Write(_emptyColumn);
        return 0;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    void SetMaxLength(int length)
    {
        if (!_countingCharLength)
            return;

        if (!ColumnMaxLength.TryAdd(_columnIndex, length))
        {
            if (ColumnMaxLength[_columnIndex] < length)
                ColumnMaxLength[_columnIndex] = length;
        }
        _columnIndex++;
    }

    /// <summary>Write string.</summary>
    public void Write(string? value)
    {
        if (string.IsNullOrEmpty(value))
        {
            WriteEmpty();
            return;
        }

        _writer.Write(
            value.Contains(Environment.NewLine)
                ? _colStartStringWrap
                : _colStartString
        );

        var index = SharedStrings.TryAdd(value, _stringIndex)
            ? _stringIndex++
            : SharedStrings[value];

        Encoding.UTF8.GetBytes($"{index}", _writer);

        _writer.Write(_colEnd);

        SetMaxLength(value.Length);
    }
    
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void WritePrimitive(bool value)
    {
        var s = $"{value}";

        _writer.Write(_colStartBoolean);
        _ = Encoding.UTF8.GetBytes(s, _writer);
        _writer.Write(_colEnd);

        SetMaxLength(s.Length);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    void WriterNumber(in ReadOnlySpan<char> chars)
    {
        _writer.Write(_colStartNumber);
        _ = Encoding.UTF8.GetBytes(chars, _writer);
        _writer.Write(_colEnd);

        SetMaxLength(chars.Length);
    }

    public void WritePrimitive(byte value) => WriterNumber($"{value}");
    public void WritePrimitive(sbyte value) => WriterNumber($"{value}");
    public void WritePrimitive(char value) => WriterNumber($"{value}");
    public void WritePrimitive(decimal value) => WriterNumber($"{value}");
    public void WritePrimitive(double value) => WriterNumber($"{value}");
    public void WritePrimitive(float value) => WriterNumber($"{value}");
    public void WritePrimitive(int value) => WriterNumber($"{value}");
    public void WritePrimitive(uint value) => WriterNumber($"{value}");
    public void WritePrimitive(long value) => WriterNumber($"{value}");
    public void WritePrimitive(ulong value) => WriterNumber($"{value}");
    public void WritePrimitive(short value) => WriterNumber($"{value}");
    public void WritePrimitive(ushort value) => WriterNumber($"{value}");

    public void WriteDateTime(DateTime value)
    {
        var d = value;
        if (d == DateTime.MinValue) WriteEmpty();
        if (d.Hour == 0 && d.Minute == 0 && d.Second == 0)
        {
            Encoding.UTF8.GetBytes(@$"<c t=""d"" s=""{XF_DATE}""><v>{d:yyyy-MM-ddTHH:mm:ss}</v></c>", _writer);
            SetMaxLength(LEN_DATE);
            return;
        }

        Encoding.UTF8.GetBytes(@$"<c t=""d"" s=""{XF_DATETIME}""><v>{d:yyyy-MM-ddTHH:mm:ss}</v></c>", _writer);
        SetMaxLength(LEN_DATETIME);
    }

    public void WriteDateTime(DateOnly value)
    {
        Encoding.UTF8.GetBytes(@$"<c t=""d"" s=""{XF_DATE}""><v>{value:yyyy-MM-dd}T00:00:00</v></c>", _writer);
        SetMaxLength(LEN_DATE);
    }

    public void WriteDateTime(TimeOnly value)
    {
        Encoding.UTF8.GetBytes(@$"<c t=""d"" s=""{XF_TIME}""><v>1900-01-01T{value:HH:mm:ss}</v></c>", _writer);
        SetMaxLength(LEN_TIME);
    }

    static void ThrowReachedMaxDepth(int depth)
    {
        throw new InvalidOperationException($"Serializer detects reached max depth:{depth}. Please check the circular reference.");
    }
}