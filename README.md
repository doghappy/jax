# jax
A bidirectional converter, JSON to XLSX, XLSX to JSON.

## NuGet
```
Install-Package DogHappy.Jax
```

# Sample

## JSON to XLSX

Prepare a JSON file

```json
{
    "Name": "Jax",
    "Age": 18,
    "Category":
    {
        "Id": 1,
        "Name": "Tool"
    }
}
```

Use these codes

```cs
string jsonPath = "config.json";
var cvt = new JsonToXlsxConverter();
using(var stream = await cvt.GetXlsxFileAsync(jsonPath))
{
    File.WriteAllBytes("result.xlsx", stream.ToArray());
}
```

Reuslt shoule be

| key | config.json |
| :- | :- |
| Name | Jax |
| Age | 18 |
| Category.Id | 1 |
| Category.Name | Tool |

## XLSX to JSON

Prepare a XLSX file

| key | config.json |
| :- | :- |
| Name | Jax |
| Age | 18 |
| Category.Id | 1 |
| Category.Name | Tool |

Use these codes

```cs
string xlsxPath = "result.xlsx";
var cvt = new XlsxToJsonConverter();
var list = await cvt.GetJsonFilesAsync(xlsxPath)
{
    foreach (var item in list)
    {
        // each item as a json file, only 1 file in this example.
        File.WriteAllBytes(item.Name, item.Stream.ToArray());
    }
}
```

> Jax supports the conversion of **multiple JSON files to an xlsx file**. with each JSON file as a column, using `master.json`'s data structure as the standard.
