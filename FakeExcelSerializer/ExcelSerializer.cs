using System;
using System.Buffers;
using System.Diagnostics;
using System.IO.Compression;
using System.Text;

namespace FakeExcelSerializer;

public class ExcelSerializer
{
    readonly byte[] _contentTypes = Encoding.UTF8.GetBytes(@"<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
<Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>
<Override PartName=""/book.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""/>
<Override PartName=""/sheet.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>
<Override PartName=""/strings.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml""/>
<Override PartName=""/styles.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml""/>
</Types>");
    readonly byte[] _rels = Encoding.UTF8.GetBytes(@"<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
<Relationship Id=""rId1"" Target=""book.xml"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument""/>
</Relationships>");

    readonly byte[] _book = Encoding.UTF8.GetBytes(@"<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
<sheets>
<sheet name=""Sheet"" sheetId=""1"" r:id=""rId1""/>
</sheets>
</workbook>");
    readonly byte[] _bookRels = Encoding.UTF8.GetBytes(@"<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
<Relationship Id=""rId1"" Target=""sheet.xml"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet""/>
<Relationship Id=""rId2"" Target=""strings.xml"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings""/>
<Relationship Id=""rId3"" Target=""styles.xml"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles""/>
</Relationships>");

    readonly string _styles = @"<styleSheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
<numFmts count=""2"">
<numFmt numFmtId=""1"" formatCode =""{0}"" />
<numFmt numFmtId=""2"" formatCode =""{1}"" />
<numFmt numFmtId=""3"" formatCode =""{2}"" />
</numFmts>
<fonts count=""1"">
<font/>
</fonts>
<fills count=""1"">
<fill/>
</fills>
<borders count=""1"">
<border/>
</borders>
<cellStyleXfs count=""1"">
<xf/>
</cellStyleXfs>
<cellXfs count=""4"">
<xf/>
<xf><alignment wrapText=""true""/></xf>
<xf numFmtId=""1""  applyNumberFormat=""1""></xf>
<xf numFmtId=""2""  applyNumberFormat=""1""></xf>
<xf numFmtId=""3""  applyNumberFormat=""1""></xf>
</cellXfs>
</styleSheet>";

    readonly byte[] _rowStart = Encoding.UTF8.GetBytes("<row>");
    readonly byte[] _rowEnd = Encoding.UTF8.GetBytes("</row>");
    readonly byte[] _colStart = Encoding.UTF8.GetBytes("<cols>");
    readonly byte[] _colEnd = Encoding.UTF8.GetBytes("</cols>");
    readonly byte[] _frozenTitleRow = Encoding.UTF8.GetBytes(@"<sheetViews>
<sheetView tabSelected=""1"" workbookViewId=""0"">
<pane ySplit=""1"" topLeftCell=""A2"" activePane=""bottomLeft"" state=""frozen""/>
</sheetView>
</sheetViews>");

    readonly byte[] _sheetStart = Encoding.UTF8.GetBytes(@"<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">");
    readonly byte[] _sheetEnd = Encoding.UTF8.GetBytes(@"</worksheet>");
    readonly byte[] _dataStart = Encoding.UTF8.GetBytes(@"<sheetData>");
    readonly byte[] _dataEnd = Encoding.UTF8.GetBytes(@"</sheetData>");

    readonly byte[] _sstStart = Encoding.UTF8.GetBytes(@"<sst xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">");
    //readonly byte[] _sstStart = Encoding.UTF8.GetBytes(@"<sst xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" uniqueCount=""1"">");
    readonly byte[] _sstEnd = Encoding.UTF8.GetBytes(@"</sst>");
    readonly byte[] _siStart = Encoding.UTF8.GetBytes("<si><t>");
    readonly byte[] _siEnd = Encoding.UTF8.GetBytes("</t></si>");

    const int COLUMN_WIDTH_MARGIN = 2;

    //public void ToFile<T>(IEnumerable<T> rows, string fileName, ExcelSerializerOptions _options)
    //{
    //    using (var stream = new FileStream(fileName, FileMode.Create, FileAccess.Write, FileShare.None))
    //        ToStream(rows, stream, _options);
    //}

    //public void ToStream<T>(IEnumerable<T> rows, Stream stream, ExcelSerializerOptions _options)
    //{
    //    using (var writer = new ExcelSerializerWriter(_options))
    //        WriteCsv(rows, stream, _options, writer);
    //}

    //static void WriteCsv<T>(IEnumerable<T> rows, Stream stream, ExcelSerializerOptions _options, ExcelSerializerWriter writer)
    //{
    //    if (_options.HasHeaderRecord && _options.HeaderTitles != null)
    //    {
    //        foreach (var t in _options.HeaderTitles)
    //        {
    //            writer.Write(t);
    //        }
    //        writer.WriteLine();
    //        writer.CopyTo(stream);
    //    }

    //    var serializer = _options.GetSerializer<T>();
    //    foreach (var row in rows)
    //    {
    //        if (row == null) continue;
    //        serializer?.Serialize(ref writer, row, _options);
    //        writer.WriteLine();
    //        writer.CopyTo(stream);
    //    }
    //}

    readonly ExcelSerializerOptions _options;
    public ExcelSerializer(ExcelSerializerOptions options)
    {
        _options = options;
    }

    public void ToFile<T>(IEnumerable<T> rows, string fileName)
    {
        var workPathRoot = Path.Combine(_options.WorkPath, Guid.NewGuid().ToString());
        if (!Directory.Exists(workPathRoot))
            Directory.CreateDirectory(workPathRoot);
        try
        {
            using (var sheetStream = CreateStream(Path.Combine(workPathRoot, "sheet.xml")))
            using (var stringsStream = CreateStream(Path.Combine(workPathRoot, "strings.xml")))
            using (var writer = new ExcelSerializerWriter(_options))
            {
                CreateSheet(rows, sheetStream, writer);
                WriteSharedStrings(stringsStream, writer);
            }

            var workRelPath = Path.Combine(workPathRoot, "_rels");
            if (!Directory.Exists(workRelPath))
                Directory.CreateDirectory(workRelPath);

            var _stylesBytes = Encoding.UTF8.GetBytes(string.Format(
                _styles,
                _options.DateTimeFormat,
                _options.DateFormat,
                _options.TimeFormat
            ));

            Task.WaitAll(
                new Task[] {
                    WriteStreamAsync(_contentTypes, Path.Combine(workPathRoot, "[Content_Types].xml")),
                    WriteStreamAsync(_book, Path.Combine(workPathRoot, "book.xml")),
                    WriteStreamAsync(_stylesBytes, Path.Combine(workPathRoot, "styles.xml")),
                    WriteStreamAsync(_rels, Path.Combine(workRelPath, ".rels")),
                    WriteStreamAsync(_bookRels, Path.Combine(workRelPath, "book.xml.rels"))
                });

            if (File.Exists(fileName))
                File.Delete(fileName);

            ZipFile.CreateFromDirectory(workPathRoot, fileName);
        }
        catch
        {
            throw;
        }
        finally
        {
            try
            {
                if (Directory.Exists(workPathRoot))
                    Directory.Delete(workPathRoot, true);
            }
            catch { }
        }
    }

    async Task WriteStreamAsync(ReadOnlyMemory<byte> bytes, string fileName)
    {
        using (var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write, FileShare.None))
            await fs.WriteAsync(bytes);
    }

    Stream CreateStream(string fileName)
        => new FileStream(fileName, FileMode.Create, FileAccess.Write, FileShare.None);

    void CreateSheet<T>(
        IEnumerable<T> rows,
        Stream stream,
        ExcelSerializerWriter writer
    )
    {
        stream.Write(_sheetStart);

        if (_options.HasHeaderRecord && _options.HeaderTitles != null && _options.HeaderTitles.Any())
            stream.Write(_frozenTitleRow);

        if (_options.AutoFitColumns)
            WriteCellWidth(rows, stream, writer);

        stream.Write(_dataStart);

        if (_options.HasHeaderRecord && _options.HeaderTitles != null && _options.HeaderTitles.Any())
        {
            writer.WriteRaw(_rowStart);
            foreach (var t in _options.HeaderTitles)
            {
                writer.Write(t);
            }
            writer.WriteRaw(_rowEnd);

            writer.CopyTo(stream);
        }

        var serializer = _options.GetSerializer<T>();
        foreach (var row in rows)
        {
            if (row == null) continue;
            writer.WriteRaw(_rowStart);
            serializer?.Serialize(ref writer, row, _options);
            writer.WriteRaw(_rowEnd);
            writer.CopyTo(stream);
        }

        writer.WriteRaw(_dataEnd);
        writer.WriteRaw(_sheetEnd);
        writer.CopyTo(stream);
    }

    void WriteCellWidth<T>(IEnumerable<T> rows, Stream stream, ExcelSerializerWriter writer)
    {
        var dummyStream = new MemoryStream();
        var serializer = _options.GetSerializer<T>();
        foreach (var row in rows.Take(_options.AutoFitDepth))
        {
            serializer?.Serialize(ref writer, row, _options);
            writer.CopyTo(dummyStream);
            dummyStream.Position = 0;
        }
        writer.Reset();

        using var buffer = new ArrayPoolBufferWriter();
        stream.Write(_colStart);
        foreach (var pair in writer.ColumnMaxLength)
        {
            var id = pair.Key + 1;
            var charLength = _options.HeaderTitles == null || _options.HeaderTitles.Length <= pair.Key
                ? pair.Value : Math.Max(pair.Value, _options.HeaderTitles[pair.Key].Length);
            var width = Math.Min(_options.AutoFitWidhtMax, charLength + COLUMN_WIDTH_MARGIN);

            Encoding.UTF8.GetBytes(
                @$"<col min=""{id}"" max =""{id}"" width =""{width:0.0}"" bestFit =""1"" customWidth =""1"" />",
                buffer);
            buffer.CopyTo(stream);
            buffer.Clear();
        }
        stream.Write(_colEnd);

        //var i = 0;
        //stream.Write(_colStart);

        //foreach (var f in writers)
        //{
        //    ++i;
        //    var max = rows
        //        .Take(100)
        //        .Select(r =>
        //        {
        //            var len = f?.Write(ref writer, r, cellWriter) ?? 0;
        //            buffer.Clear();
        //            return len;
        //        })
        //        .Max(x => x);

        //    var lenMax = Math.Max(max, f.Name.Length) + COLUMN_WIDTH_MARGIN;
        //    var maxLength = Math.Min(lenMax, COLUMN_WIDTH_MAX);
        //    Encoding.UTF8.GetBytes(
        //        @$"<col min=""{i}"" max =""{i}"" width =""{maxLength:0.0}"" bestFit =""1"" customWidth =""1"" />",
        //        buffer);
        //    stream.Write(buffer.WrittenSpan);
        //    buffer.Clear();
        //}
        //stream.Write(_colEnd);
    }

    void WriteSharedStrings(Stream stream, ExcelSerializerWriter writer)
    {
        stream.Write(_sstStart);
        using var buffer = new ArrayPoolBufferWriter();
        foreach (var s in writer.SharedStrings.Keys)
        {
            buffer.Write(_siStart);
            Encoding.UTF8.GetBytes(s, buffer);
            buffer.Write(_siEnd);
            buffer.CopyTo(stream);
        }
        stream.Write(_sstEnd);
    }
}
