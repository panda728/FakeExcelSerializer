# FakeExcelSerializer
Convert the object to an Excel readable file.　(.xlsx)

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

　https://docs.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/2c5dee00-eff2-4b22-92b6-0738acd4475e
  
## License
This library is licensed under the MIT License.
