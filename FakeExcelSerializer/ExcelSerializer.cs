using System.Buffers;
using System.IO.Compression;
using System.Text;

namespace FakeExcelSerializer;

public static class ExcelSerializer
{
    readonly static byte[] _contentTypes = Encoding.UTF8.GetBytes(@"<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
<Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>
<Override PartName=""/book.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""/>
<Override PartName=""/sheet.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>
<Override PartName=""/strings.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml""/>
<Override PartName=""/styles.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml""/>
</Types>");
    readonly static byte[] _rels = Encoding.UTF8.GetBytes(@"<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
<Relationship Id=""rId1"" Target=""book.xml"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument""/>
</Relationships>");

    readonly static byte[] _book = Encoding.UTF8.GetBytes(@"<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
<sheets>
<sheet name=""Sheet"" sheetId=""1"" r:id=""rId1""/>
</sheets>
</workbook>");
    readonly static byte[] _bookRels = Encoding.UTF8.GetBytes(@"<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
<Relationship Id=""rId1"" Target=""sheet.xml"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet""/>
<Relationship Id=""rId2"" Target=""strings.xml"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings""/>
<Relationship Id=""rId3"" Target=""styles.xml"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles""/>
</Relationships>");

    readonly static string _styles = @"<styleSheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
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

    readonly static byte[] _rowStart = Encoding.UTF8.GetBytes("<row>");
    readonly static byte[] _rowEnd = Encoding.UTF8.GetBytes("</row>");
    readonly static byte[] _colStart = Encoding.UTF8.GetBytes("<cols>");
    readonly static byte[] _colEnd = Encoding.UTF8.GetBytes("</cols>");
    readonly static byte[] _frozenTitleRow = Encoding.UTF8.GetBytes(@"<sheetViews>
<sheetView tabSelected=""1"" workbookViewId=""0"">
<pane ySplit=""1"" topLeftCell=""A2"" activePane=""bottomLeft"" state=""frozen""/>
</sheetView>
</sheetViews>");

    readonly static byte[] _sheetStart = Encoding.UTF8.GetBytes(@"<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">");
    readonly static byte[] _sheetEnd = Encoding.UTF8.GetBytes(@"</worksheet>");
    readonly static byte[] _dataStart = Encoding.UTF8.GetBytes(@"<sheetData>");
    readonly static byte[] _dataEnd = Encoding.UTF8.GetBytes(@"</sheetData>");

    readonly static byte[] _sstStart = Encoding.UTF8.GetBytes(@"<sst xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">");
    //readonly byte[] _sstStart = Encoding.UTF8.GetBytes(@"<sst xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" uniqueCount=""1"">");
    readonly static byte[] _sstEnd = Encoding.UTF8.GetBytes(@"</sst>");
    readonly static byte[] _siStart = Encoding.UTF8.GetBytes("<si><t>");
    readonly static byte[] _siEnd = Encoding.UTF8.GetBytes("</t></si>");

    const int COLUMN_WIDTH_MARGIN = 2;

    public static void ToFile<T>(IEnumerable<T> rows, string fileName, ExcelSerializerOptions options)
    {
        var workPathRoot = Path.Combine(options.WorkPath, Guid.NewGuid().ToString());
        if (!Directory.Exists(workPathRoot))
            Directory.CreateDirectory(workPathRoot);
        try
        {
            using (var sheetStream = CreateStream(Path.Combine(workPathRoot, "sheet.xml")))
            using (var stringsStream = CreateStream(Path.Combine(workPathRoot, "strings.xml")))
            using (var writer = new ExcelSerializerWriter(options))
            {
                CreateSheet(rows, sheetStream, writer, options);
                WriteSharedStrings(stringsStream, writer);
            }

            var workRelPath = Path.Combine(workPathRoot, "_rels");
            if (!Directory.Exists(workRelPath))
                Directory.CreateDirectory(workRelPath);

            var _stylesBytes = Encoding.UTF8.GetBytes(string.Format(
                _styles,
                options.DateTimeFormat,
                options.DateFormat,
                options.TimeFormat
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

    static async Task WriteStreamAsync(ReadOnlyMemory<byte> bytes, string fileName)
    {
        using (var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write, FileShare.None))
            await fs.WriteAsync(bytes);
    }

    static Stream CreateStream(string fileName)
        => new FileStream(fileName, FileMode.Create, FileAccess.Write, FileShare.None);

    static void CreateSheet<T>(
        IEnumerable<T> rows,
        Stream stream,
        ExcelSerializerWriter writer,
        ExcelSerializerOptions options
    )
    {
        stream.Write(_sheetStart);

        if (options.HasHeaderRecord && options.HeaderTitles != null && options.HeaderTitles.Any())
            stream.Write(_frozenTitleRow);

        if (options.AutoFitColumns)
            WriteCellWidth(rows, stream, writer, options);

        stream.Write(_dataStart);

        if (options.HasHeaderRecord && options.HeaderTitles != null && options.HeaderTitles.Any())
        {
            writer.WriteRaw(_rowStart);
            foreach (var t in options.HeaderTitles)
            {
                writer.Write(t);
            }
            writer.WriteRaw(_rowEnd);
            writer.CopyTo(stream);
        }

        var serializer = options.GetSerializer<T>();
        foreach (var row in rows)
        {
            if (row == null) continue;
            writer.WriteRaw(_rowStart);
            serializer?.Serialize(ref writer, row, options);
            writer.WriteRaw(_rowEnd);
            writer.CopyTo(stream);
        }

        writer.WriteRaw(_dataEnd);
        writer.WriteRaw(_sheetEnd);
        writer.CopyTo(stream);
    }

    static void WriteCellWidth<T>(
        IEnumerable<T> rows,
        Stream stream,
        ExcelSerializerWriter writer,
        ExcelSerializerOptions options
    )
    {
        // Counting the number of characters in Writer's internal process
        // The result is stored in writer.ColumnMaxLength 
        var serializer = options.GetSerializer<T>();

        if (options.HasHeaderRecord && options.HeaderTitles != null)
        {
            foreach (var t in options.HeaderTitles)
                writer.Write(t);
            writer.Clear();
        }
        foreach (var row in rows.Take(options.AutoFitDepth))
        {
            serializer?.Serialize(ref writer, row, options);
            writer.Clear();
        }
        writer.StopCountingCharLength();

        using var buffer = new ArrayPoolBufferWriter();
        stream.Write(_colStart);
        foreach (var pair in writer.ColumnMaxLength)
        {
            var id = pair.Key + 1;
            var width = Math.Min(options.AutoFitWidhtMax, pair.Value + COLUMN_WIDTH_MARGIN);

            Encoding.UTF8.GetBytes(
                @$"<col min=""{id}"" max =""{id}"" width =""{width:0.0}"" bestFit =""1"" customWidth =""1"" />",
                buffer);
            buffer.CopyTo(stream);
            buffer.Clear();
        }
        stream.Write(_colEnd);
    }

    static void WriteSharedStrings(Stream stream, ExcelSerializerWriter writer)
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
