# OfficeTools
A library that supports filling data and exporting word from template.

**1. Easy to design new template:**
    - Select text then insert a bookmark
<p align="left">
![select text then insert a bookmark](https://github.com/luuducly/OfficeTools/assets/69654714/ef495cb7-7f4e-4bce-99c4-4905783c12ac)
</p>
    - Enter bookmark name, irrespective of lowercase or uppercase
<p align="left">
![enter bookmark name](https://github.com/luuducly/OfficeTools/assets/69654714/4bde70b1-a601-4100-8865-5c59c1e2cc60)
</p>
    - Before or after inserted bookmark, you can continue to format your template as you want
<p align="left">
![continue to format your template](https://github.com/luuducly/OfficeTools/assets/69654714/7ef9f7cf-e8c1-40e9-b3f3-a81c198016e3)
</p>

**2. Build-in support data type:** Text, Image, QrCode, BarCode, HTML, Document
    - Insert Image, QrCode or BarCode inside a Textbox shape for fixed frame size
<p align="left">
![insert an image bookmark inside textbox shape](https://github.com/luuducly/OfficeTools/assets/69654714/014157bb-3b47-4ab6-bf87-3da3db979ffc)
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
          using (var newDoc = wordTemplate.Export(data))
          {
              newDoc.SaveAs("Output.docx");
          }
      }
  }
  ```
**3. Auto repeating depends on input data object**
```csharp
var data = new { Name = "Mr. Smith", Dependants:["Peter", "Laura"] };
using (FileStream fileStream = new FileStream("PATH_TO_YOUR_TEMPLATE\\Template.docx",
FileMode.Open, FileAccess.ReadWrite))
{
    using (WordTemplate wordTemplate = new WordTemplate(fileStream))
    {
        wordTemplate.Bookmarks["Name"].DataType = OfficeTools.DataType.Text;
        wordTemplate.Bookmarks["Dependants"].DataType = OfficeTools.DataType.Text;
        var newDoc = wordTemplate.Export(data);
        newDoc.SaveAs("Output.docx");
        newDoc.Dispose();
    }
}
```
**4. Flexible custom data type by creating new IReplacer**
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
