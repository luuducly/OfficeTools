# OfficeTools
A library that supports filling data and exporting word from template.

**1. Easy to design new template:**<br/>
<p align="left">
    <img alt="select text then insert a bookmark" src="https://github.com/luuducly/OfficeTools/assets/69654714/26426e2e-dd83-420e-a6c0-8a78b1109eb9"/>
</p>
<p align="left">
    <img alt="enter bookmark name" src="https://github.com/luuducly/OfficeTools/assets/69654714/942a1843-1254-4c61-84d6-f3fdcde8a8d5"/>
</p>

**2. Build-in support data type:** Text, Image, QrCode, BarCode, HTML, Document<br/>
    - Insert Image, QrCode or BarCode inside a Textbox shape for fixed frame size
<p align="left">
    <img alt="insert an image bookmark inside textbox shape" src="https://github.com/luuducly/OfficeTools/assets/69654714/014157bb-3b47-4ab6-bf87-3da3db979ffc"/>
</p>

   ```csharp
  var data = new { Name = "Mr. Smith", Avatar:"Base64StringHere" };
  using (FileStream fileStream = new FileStream("PATH_TO_YOUR_TEMPLATE\\Template.docx",
  FileMode.Open, FileAccess.ReadWrite))
  {
      using (WordTemplate wordTemplate = new WordTemplate(fileStream))
      {
          wordTemplate.Bookmarks["Name"].DataType = OfficeTools.DataType.Text;
          wordTemplate.Bookmarks["Avatar"].DataType = OfficeTools.DataType.Image;
          using (var docStream = wordTemplate.Export(data))
          {
              using (var newFileStream = File.Create("Output.docx"))
              {
                  docStream.Seek(0, SeekOrigin.Begin);
                  docStream.CopyTo(newFileStream);
              }
          }
      }
  }
  ```

**3. Auto repeating depends on input data object**<br/>
```csharp
var data = new { Name = "Mr. Smith", Dependants:["Peter", "Laura"] };
using (FileStream fileStream = new FileStream("PATH_TO_YOUR_TEMPLATE\\Template.docx",
FileMode.Open, FileAccess.ReadWrite))
{
    using (WordTemplate wordTemplate = new WordTemplate(fileStream))
    {
        wordTemplate.Bookmarks["Name"].DataType = OfficeTools.DataType.Text;
        wordTemplate.Bookmarks["Dependants"].DataType = OfficeTools.DataType.Text;
        using (var docStream = wordTemplate.Export(data))
       {
            using (var newFileStream = File.Create("Output.docx"))
            {
                docStream.Seek(0, SeekOrigin.Begin);
                docStream.CopyTo(newFileStream);
            }
       }
    }
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

wordTemplate.Bookmarks["Name"].Replacer = new NewBarCodeReplacer();
```
**5. Support both windows and linux OS**

<a href='https://github.com/luuducly/OfficeTools/tree/main/src/OfficeTools.Example'>Please visit this link to find more examples!</a>
