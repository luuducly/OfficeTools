# OfficeTools
A useful cross-platform library to export Word template with repeating data, image, QR code...

Install [OfficeTools package](https://www.nuget.org/packages/OfficeTools) and its dependencies using NuGet Package Manager:
```powershell
Install-Package OfficeTools 
```

**1. Easy to design new template:**<br/>
<p align="left">
    <img alt="select text then insert a bookmark" src="https://github.com/luuducly/OfficeTools/assets/69654714/26426e2e-dd83-420e-a6c0-8a78b1109eb9"/>
</p>
<p align="left">
    <img alt="enter bookmark name" src="https://github.com/luuducly/OfficeTools/assets/69654714/942a1843-1254-4c61-84d6-f3fdcde8a8d5"/>
</p>

**2. Build-in support data type:** Text, Image, QrCode, BarCode, HTML, Document<br/>

   ```csharp
    using (var wordTemplate = new WordTemplate(templateStream))
    {
        //text type is default value, so we dont need to assign
        //wordTemplate.Bookmarks["Text"].DataType = DataType.Text;
        wordTemplate.Bookmarks["HTMLInShape"].DataType = DataType.HTML;
        wordTemplate.Bookmarks["BarcodeInShape"].DataType = DataType.BarCode;
        wordTemplate.Bookmarks["QrcodeInShape"].DataType = DataType.QRCode;
        //wordTemplate.Bookmarks["RepeatNumber"].DataType = DataType.Text;
        //wordTemplate.Bookmarks["RepeatPurpose"].DataType = DataType.Text;
        //wordTemplate.Bookmarks["Order"].DataType = DataType.Text;
        //wordTemplate.Bookmarks["Name"].DataType = DataType.Text;
        //wordTemplate.Bookmarks["DOB"].DataType = DataType.Text;
        //wordTemplate.Bookmarks["FieldName"].DataType = DataType.Text;
        //wordTemplate.Bookmarks["Toggle"].DataType = DataType.Text;
        wordTemplate.Bookmarks["Equation"].DataType = DataType.Document;
    
        using (var exportedStream = wordTemplate.Export(data1))
        {
            using (var output = File.Create("Output1.docx"))
            {
                exportedStream.CopyTo(output);
            }
        }    
    }
  ```

**3. Auto repeating depends on input data object**<br/>
```csharp
{
    "RepeatParaGraph": [
        {
            "RepeatNumber": "1st",
            "RepeatPurpose": "testing first time."
        },
        {
            "RepeatNumber": "2nd",
            "RepeatPurpose": "testing second time."
        },
        {
            "RepeatNumber": "3rd",
            "RepeatPurpose": "testing third time."
        }
    ],
    "RepeatTable": [
        {
            "Order": 1,
            "Name": "Brian",
            "DOB": "20/05/1987"
        },
        {
            "Order": 2,
            "Name": "Wade",
            "DOB": "02/10/1992"
        },
        {
            "Order": 3,
            "Name": "Carlos",
            "DOB": "20/03/1986"
        },
        {
            "Order": 4,
            "Name": "Dave",
            "DOB": "30/09/2014"
        },
        {
            "Order": 5,
            "Name": "Gilbert",
            "DOB": "20/05/2017"
        }
    ]
}
```

**4. Flexible custom data type by creating new IReplacer**<br/>
```csharp
internal class NewBarCodeReplacer : QRCodeReplacer, IReplacer
{
    public override List<OpenXmlElement> GenerateElements(WordprocessingDocument document, Bookmark bookmark)
    {
        var elements = base.GenerateElements(document, bookmark);
        if(RawData != null)
            elements.Add(new Run(new Text(RawData.ToString())));
        return elements;
    }
}

wordTemplate.Bookmarks["QrcodeInShape"].Replacer = new NewBarCodeReplacer();
```
**5. Support both Windows and Linux OS**

<a href='https://github.com/luuducly/OfficeTools/tree/main/src/OfficeTools.Example'>Please visit this link to find more examples!</a>
