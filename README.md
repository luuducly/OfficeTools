# WordTemplater
A useful cross-platform library to export Word template with formater, repeating data, image, QR code...

Install [WordTemplater package](https://www.nuget.org/packages/WordTemplater) and its dependencies using NuGet Package Manager:
```powershell
Install-Package WordTemplater 
```

**1. Easy to design new template with merge field:**<br/>
<p align="left">
    <img alt="Word template file sample" src="https://github.com/user-attachments/assets/cf5c858b-4796-4040-840f-955155fa0358"/>
</p>

**2. Build-in many useful function and easy to customize more:** Formating, Repeating, Condition, Image, QrCode, BarCode, HTML, Document...<br/>

   ```csharp
    using (var templateStream = File.OpenRead("Templates\\Template1.docx"))
    {
        using (var wordTemplate = new WordTemplate(templateStream))
        {
            wordTemplate.RegisterEvaluator("customizable", new Number2TextEvaluator());
            wordTemplate.RegisterEvaluator("upperFirstLetter", new UpperCaseFirstLetter());
            using (var exportedStream = wordTemplate.Export(data1))
            {
                using (var output = File.Create("Output1.docx"))
                {
                    exportedStream.CopyTo(output);
                }
            }    
        }
    }
  ```

**3. Support both Windows and Linux OS**

<a href='https://github.com/luuducly/WordTemplater/tree/main/src/WordTemplater.Example'>Find more examples in this link.</a>
