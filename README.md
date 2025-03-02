# WordTemplater
A useful cross-platform library to export Word template with formater, repeating data, image, QR code...

Install [WordTemplater package](https://www.nuget.org/packages/WordTemplater) and its dependencies using NuGet Package Manager:
```powershell
Install-Package WordTemplater 
```

**1. Easy to design new template with merge field:**<br/>
<p align="left">
    ![1](https://github.com/user-attachments/assets/14b55b06-09ec-4450-8e35-0fc7108b1aaa)
    <img alt="Word template file sample" src="https://github.com/user-attachments/assets/14b55b06-09ec-4450-8e35-0fc7108b1aaa"/>
</p>
<p align="left">
    ![2](https://github.com/user-attachments/assets/f4d77bb8-1126-4e90-bf33-d80dff2e25cd)
    <img alt="Word template file sample" src="https://github.com/user-attachments/assets/f4d77bb8-1126-4e90-bf33-d80dff2e25cd"/>
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
