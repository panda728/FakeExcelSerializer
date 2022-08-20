# FakeExcelSerializer
Convert the object to an Excel readable file.　(.xlsx)

### Getting Started
Supporting platform is .NET 6.

~~~
PM> Install-Package FakeExcelSerializer
~~~

### Usage
You can use `FakeExcelSerializer.ToFile`.

~~~
ExcelSerializer.ToFile(Users, "test.xlsx", ExcelSerializerOptions.Default);
~~~

### Note

For the method of retrieving values from IEnumerable\<T\>, Cysharp's WebSerializer method is used.

　https://github.com/Cysharp/WebSerializer
  
The following page provides information on how to return to OpenOfficeXml.

　https://gist.github.com/iso2022jp/721df3095f4df512bfe2327503ea1119

　https://docs.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/2c5dee00-eff2-4b22-92b6-0738acd4475e
  
### License
This library is licensed under the MIT License.
