# FakeExcelSerializer

IEnumerable\<T\>をExcelファイル(.xlsx)に変換します。

### Usage

~~~
ExcelSerializer.ToFile(Users, "test.xlsx", ExcelSerializerOptions.Default);
~~~


### シリアライズの方式について

IEnumerable<T>から値を取り出す方法については
Cysharp様のWebSerializerの方式を使用しています。

効率的でなおかつ拡張性もある非常に素晴らしい構成です。

  https://github.com/Cysharp/WebSerializer
  
### CSV ライクなデータ出力のバージョンもあります
  
  https://github.com/panda728/FakeCsvSerializer
