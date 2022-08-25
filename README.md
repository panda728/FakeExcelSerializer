# FakeExcelSerializer
Convert object to Excel file (.xlsx) [Open XML SpreadsheetML File Format]

## Getting Started
Supporting platform is .NET 6.

~~~
PM> Install-Package FakeExcelSerializer
~~~

## Usage
You can use `ExcelSerializer.ToFile`.

~~~
ExcelSerializer.ToFile(Users, "test.xlsx", ExcelSerializerOptions.Default);
~~~

## Notice

Folder creation permissions are required since a working folder will be used.

## Examples

If you pass an object, it will be converted to an Excel file.
~~~
ExcelSerializer.ToFile(new string[] { "test", "test2" }, @"c:\test\test.xlsx", ExcelSerializerOptions.Default);
~~~
![image](https://user-images.githubusercontent.com/16958552/185727609-79b574e8-b40c-46dc-83c9-74b078a1f44a.png)

Passing a class expands the property into a column.
~~~
public class Portal
{
    public string Name { get; set; }
    public string Owner { get; set; }
    public int Level { get; set; }
}

var potals = new Portal[] {
    new Portal { Name = "Portal1", Owner = "panda728", Level = 8 },
    new Portal { Name = "Portal2", Owner = "panda728", Level = 1 },
    new Portal { Name = "Portal3", Owner = "panda728", Level = 2 },
};

ExcelSerializer.ToFile(potals, @"c:\test\potals.xlsx", ExcelSerializerOptions.Default);
~~~
![image](https://user-images.githubusercontent.com/16958552/185727657-3e41dea7-1af4-4a52-99bd-1457f895b564.png)


Options can be set to display a title line and automatically adjust column widths.
~~~
var newConfig = ExcelSerializerOptions.Default with
{
    CultureInfo = CultureInfo.InvariantCulture,
    HasHeaderRecord = true,
    HeaderTitles = new string[] { "Name", "Owner", "Level" },
    AutoFitColumns = true,
};
ExcelSerializer.ToFile(potals, @"c:\test\potalsOp.xlsx", newConfig);
~~~
![image](https://user-images.githubusercontent.com/16958552/185727708-18201283-bb0b-46ba-a413-dbe34c20f3a3.png)


## Note

For the method of retrieving values from IEnumerable\<T\>, Cysharp's WebSerializer method is used.

　https://github.com/Cysharp/WebSerializer
  
The following page provides information on how to return to OpenOfficeXml.

　https://gist.github.com/iso2022jp/721df3095f4df512bfe2327503ea1119

　https://docs.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/2c5dee00-eff2-4b22-92b6-0738acd4475## Benckma

## Benchmark
N = 100 lines.

|              Method |   N |        Mean |      Error |     StdDev |      Median | Ratio | RatioSD |      Gen 0 |      Gen 1 |     Gen 2 |  Allocated |
|-------------------- |---- |------------:|-----------:|-----------:|------------:|------:|--------:|-----------:|-----------:|----------:|-----------:|
|           ClosedXml |   1 |    63.63 ms |   1.730 ms |   5.048 ms |    63.00 ms |  1.00 |    0.00 |  1000.0000 |          - |         - |   5,738 KB |
| FakeExcelSerializer |   1 |    19.62 ms |   0.392 ms |   0.677 ms |    19.53 ms |  0.30 |    0.03 |          - |          - |         - |     126 KB |
|                     |     |             |            |            |             |       |         |            |            |           |            |
|           ClosedXml |  10 |   631.47 ms |  12.544 ms |  17.170 ms |   627.75 ms |  1.00 |    0.00 |  9000.0000 |  1000.0000 |         - |  52,660 KB |
| FakeExcelSerializer |  10 |    29.66 ms |   0.592 ms |   1.629 ms |    29.11 ms |  0.05 |    0.00 |   156.2500 |    31.2500 |         - |     661 KB |
|                     |     |             |            |            |             |       |         |            |            |           |            |
|           ClosedXml | 100 | 6,980.32 ms | 127.032 ms | 112.610 ms | 6,941.43 ms |  1.00 |    0.00 | 91000.0000 | 22000.0000 | 5000.0000 | 513,934 KB |
| FakeExcelSerializer | 100 |    82.42 ms |   1.546 ms |   1.655 ms |    82.30 ms |  0.01 |    0.00 |  1428.5714 |   142.8571 |         - |   6,004 KB |

## Sample Extensions
WindowsForm's DataGridView to .xlsx

https://github.com/panda728/DataGridViewDump

## Link
CSV-like File output version
　https://github.com/panda728/FakeCsvSerializer

## License
This library is licensed under the MIT License.
