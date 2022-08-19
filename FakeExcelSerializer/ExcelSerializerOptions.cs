using System.Globalization;
using System.Text;

namespace FakeExcelSerializer;

public record ExcelSerializerOptions(IExcelSerializerProvider Provider)
{
    public static ExcelSerializerOptions Default { get; } = new ExcelSerializerOptions(ExcelSerializerProvider.Default);

    public CultureInfo? CultureInfo { get; init; }

    public int MaxDepth { get; init; } = 64;

    public bool AutoFitColumns { get; init; } = false;
    public int AutoFitDepth { get; init; } = 200;
    public int AutoFitWidhtMax { get; init; } = 100;
    public string WorkPath { get; init; } = "";

    public string DateTimeFormat { get; init; } = "yyyy/mm/dd hh:mm;@";
    public string DateFormat { get; init; } = "yyyy/mm/dd;@";
    public string TimeFormat { get; init; } = "hh:mm;@";

    public bool HasHeaderRecord { get; init; } = false;
    public string[]? HeaderTitles { get; init; }

    public IExcelSerializer<T>? GetSerializer<T>()
        => Provider.GetSerializer<T>();

    public IExcelSerializer<T> GetRequiredSerializer<T>()
    {
        var serializer = Provider.GetSerializer<T>();
        if (serializer == null) Throw(typeof(T));
        return serializer!;
    }

    void Throw(Type type)
    {
        throw new InvalidOperationException($"Type is not found in provider. Type:{type}");
    }
}
