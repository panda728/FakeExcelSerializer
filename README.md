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

## Benchmark
N = 100 lines.

BenchmarkDotNet=v0.13.1, OS=Windows 10.0.19043.1889 (21H1/May2021Update)
11th Gen Intel Core i7-1165G7 2.80GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK=6.0.301
  [Host]     : .NET 6.0.6 (6.0.622.26707), X64 RyuJIT  [AttachedDebugger]
  DefaultJob : .NET 6.0.6 (6.0.622.26707), X64 RyuJIT


|              Method |   N |         Mean |      Error |     StdDev | Ratio |      Gen 0 |      Gen 1 |     Gen 2 |  Allocated |
|-------------------- |---- |-------------:|-----------:|-----------:|------:|-----------:|-----------:|----------:|-----------:|
|           ClosedXml |   1 |    36.078 ms |  0.3324 ms |  0.3109 ms |  1.00 |   857.1429 |   357.1429 |         - |   5,734 KB |
| FakeExcelSerializer |   1 |     4.587 ms |  0.0397 ms |  0.0332 ms |  0.13 |    15.6250 |     7.8125 |         - |     127 KB |
|                     |     |              |            |            |       |            |            |           |            |
|           ClosedXml |  10 |   343.416 ms |  5.7000 ms |  4.7598 ms |  1.00 |  8000.0000 |  1000.0000 |         - |  52,661 KB |
| FakeExcelSerializer |  10 |     9.067 ms |  0.0850 ms |  0.0709 ms |  0.03 |    93.7500 |    31.2500 |         - |     661 KB |
|                     |     |              |            |            |       |            |            |           |            |
|           ClosedXml | 100 | 3,663.531 ms | 23.3744 ms | 20.7208 ms |  1.00 | 81000.0000 | 22000.0000 | 5000.0000 | 513,936 KB |
| FakeExcelSerializer | 100 |    47.989 ms |  0.9378 ms |  0.8773 ms |  0.01 |   909.0909 |    90.9091 |         - |   6,005 KB |

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
## Sample Extensions
WindowsForm's DataGridView to .xlsx

https://github.com/panda728/DataGridViewDump

## Link
CSV-like File output version
　https://github.com/panda728/FakeCsvSerializer

## License
This library is licensed under the MIT License.
