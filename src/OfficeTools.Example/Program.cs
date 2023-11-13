using Newtonsoft.Json.Linq;
using OfficeTools;

var json1 = File.ReadAllText("DataSamples\\Data1.json");
var data1 = JObject.Parse(json1);
var equationFile = File.ReadAllBytes("Templates\\Equation.docx");
data1["Equation"] = Convert.ToBase64String(equationFile);

using (var templateStream = File.OpenRead("Templates\\Template1.docx"))
{
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
}    