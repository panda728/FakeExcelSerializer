//using System.Buffers;
//using System.Text;

//namespace FakeExcelSerializer
//{
//    /// <summary>
//    /// 文字列の辞書を管理。同じ値は同じIDで出力するため
//    /// </summary>
//    public class CellWriter
//    {
//        //private const int XF_NORMAL = 0;
//        const int XF_WRAP_TEXT = 1;
//        const int XF_DATE = 2;
//        const int XF_DATETIME = 3;

//        static readonly byte[] _emptyColumn = Encoding.UTF8.GetBytes("<c></c>");
//        static readonly byte[] _colStartBoolean = Encoding.UTF8.GetBytes(@"<c t=""b""><v>");
//        static readonly byte[] _colStartNumber = Encoding.UTF8.GetBytes(@"<c t=""n""><v>");
//        static readonly byte[] _colStartStringWrap = Encoding.UTF8.GetBytes(@$"<c t=""s"" s=""{XF_WRAP_TEXT}""><v>");
//        static readonly byte[] _colStartString = Encoding.UTF8.GetBytes(@$"<c t=""s""><v>");
//        static readonly byte[] _colEnd = Encoding.UTF8.GetBytes(@"</v></c>");

//        public int WriteEmpty(ref IBufferWriter<byte> _writer)
//        {
//            _writer.Write(_emptyColumn);
//            return 0;
//        }

//        public int Write(in string value, ref IBufferWriter<byte> _writer)
//        {
//            if (string.IsNullOrEmpty(value)) WriteEmpty(ref _writer);

//            _writer.Write(
//                value.Contains(Environment.NewLine)
//                    ? _colStartStringWrap
//                    : _colStartString
//            );

//            var index = SharedStrings.TryAdd(value, _index)
//                ? _index++
//                : SharedStrings[value];
            
//            Encoding.UTF8.GetBytes($"{index}", _writer);

//            _writer.Write(_colEnd);
//            return value.Length;
//        }

//        int _index = 0;
//        /// <summary>
//        /// 文字列の辞書を管理。同じ値は同じIDで出力するため
//        /// </summary>
//        public Dictionary<string, int> SharedStrings { get; } = new();
//        public void SharedStringsClear()
//        {
//            _index = 0;
//            SharedStrings.Clear();
//        }

//        public int Write(object? value, ref IBufferWriter<byte> _writer)
//            => Write(value?.ToString() ?? "", ref _writer);
//        public int Write(Guid value, ref IBufferWriter<byte> _writer)
//            => Write(value.ToString(), ref _writer);
//        public int Write(Enum value, ref IBufferWriter<byte> _writer)
//            => Write(value.ToString(), ref _writer);

//        public int Write(bool? value, ref IBufferWriter<byte> _writer)
//        {
//            if (value == null) return WriteEmpty(ref _writer);
//            var s = $"{value}";
//            if (s == null) return WriteEmpty(ref _writer);

//            _writer.Write(_colStartBoolean);
//            _ = Encoding.UTF8.GetBytes(s, _writer);
//            _writer.Write(_colEnd);
//            return s.Length;
//        }

//        int WriterNumber(in ReadOnlySpan<char> chars, ref IBufferWriter<byte> _writer)
//        {
//            _writer.Write(_colStartNumber);
//            _ = Encoding.UTF8.GetBytes(chars, _writer);
//            _writer.Write(_colEnd);
//            return chars.Length;
//        }

//        public int Write(int value, ref IBufferWriter<byte> _writer)
//            => WriterNumber($"{value}", ref _writer);
//        public int Write(long value, ref IBufferWriter<byte> _writer)
//            => WriterNumber($"{value}", ref _writer);
//        public int Write(float value, ref IBufferWriter<byte> _writer)
//            => WriterNumber($"{value}", ref _writer);
//        public int Write(double value, ref IBufferWriter<byte> _writer)
//            => WriterNumber($"{value}", ref _writer);
//        public int Write(decimal value, ref IBufferWriter<byte> _writer)
//            => WriterNumber($"{value}", ref _writer);
//        public int Write(short value, ref IBufferWriter<byte> _writer)
//            => WriterNumber($"{value}", ref _writer);
//        public int Write(ushort value, ref IBufferWriter<byte> _writer)
//            => WriterNumber($"{value}", ref _writer);

//        const int LEN_DATE = 10;
//        const int LEN_DATETIME = 18;
//        public int Write(DateTime value, ref IBufferWriter<byte> _writer)
//        {
//            var d = value;
//            if (d == DateTime.MinValue) WriteEmpty(ref _writer);
//            if (d.Hour == 0 && d.Minute == 0 && d.Second == 0)
//            {
//                Encoding.UTF8.GetBytes(@$"<c t=""d"" s=""{XF_DATE}""><v>{d:yyyy-MM-ddTHH:mm:ss}</v></c>", _writer);
//                return LEN_DATE;
//            }
//            Encoding.UTF8.GetBytes(@$"<c t=""d"" s=""{XF_DATETIME}""><v>{d:yyyy-MM-ddTHH:mm:ss}</v></c>", _writer);
//            return LEN_DATETIME;
//        }

//        public int Write(DateOnly? value, ref IBufferWriter<byte> _writer)
//        {
//            if (value == null) WriteEmpty(ref _writer);
//            Encoding.UTF8.GetBytes(@$"<c t=""d"" s=""{XF_DATE}""><v>{value:yyyy-MM-ddTHH:mm:ss}</v></c>", _writer);
//            return LEN_DATE;
//        }

//        public int Write(TimeOnly? value, ref IBufferWriter<byte> _writer)
//        {
//            if (value == null) WriteEmpty(ref _writer);
//            Encoding.UTF8.GetBytes(@$"<c t=""d"" s=""{XF_DATETIME}""><v>{value:yyyy-MM-ddTHH:mm:ss}</v></c>", _writer);
//            return LEN_DATETIME;
//        }
//    }
//}