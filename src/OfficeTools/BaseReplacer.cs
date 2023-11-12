using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using DRAW = DocumentFormat.OpenXml.Drawing;
using System.Text;
using Newtonsoft.Json.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.VariantTypes;

namespace OfficeTools
{
    public interface IReplacer : IDisposable
    {
        public object RawData { get; internal set; }
        public object FormatedData { get; internal set; }
        public object FormatData(object data);
        public List<OpenXmlElement> GenerateElements(WordprocessingDocument document, Bookmark bookmark);
        public void InsertElements(List<OpenXmlElement> elements, WordprocessingDocument document, Bookmark bookmark);
    }

    public class BaseReplacer : IReplacer
    {
        public object RawData { get; set; }
        public object FormatedData { get; set; }

        public virtual object FormatData(object data)
        {
            return data;
        }

        public virtual List<OpenXmlElement> GenerateElements(WordprocessingDocument document, Bookmark bookmark)
        {
            List<OpenXmlElement> elements = new List<OpenXmlElement>();
            if (FormatedData != null)
            {
                var style = Utils.CloneStyle(bookmark);
                var run = style is null
                    ? new Run(new Text(FormatedData.ToString()))
                    : new Run(style, new Text(FormatedData.ToString()));
                elements.Add(run);
            }
            return elements;
        }

        public virtual void InsertElements(List<OpenXmlElement> elements, WordprocessingDocument document, Bookmark bookmark)
        {
            if (elements != null)
            {
                elements.ForEach(element =>
                {
                    bookmark.BookmarkStart.InsertBeforeSelf(element);
                });
            }
        }

        public void Dispose()
        {
            if(FormatedData is IDisposable)
            {
                ((IDisposable)FormatedData).Dispose();
            }
        }
    }

    public class ImageReplacer : BaseReplacer, IReplacer
    {
        public virtual object FormatData(object data)
        {
            if (data != null)
                return new MemoryStream(Convert.FromBase64String(data.ToString()));
            else
                return null;
        }

        public virtual List<OpenXmlElement> GenerateElements(WordprocessingDocument document, Bookmark bookmark)
        {
            List<OpenXmlElement> elements = new List<OpenXmlElement>();
            var imgElement = Utils.CreateNewPictureElement(Guid.NewGuid().ToString(), 0, 0);
            elements.Add(imgElement);
            return elements;
        }

        public virtual void InsertElements(List<OpenXmlElement> elements, WordprocessingDocument document, Bookmark bookmark)
        {
            if (FormatedData is Stream)
            {
                Stream stream = (Stream)FormatedData;
                if (elements != null)
                {
                    var drawing = bookmark.BookmarkStart.Ancestors<Drawing>().FirstOrDefault();
                    if (drawing != null)
                    {
                        var run = drawing.Ancestors<Run>().FirstOrDefault();

                        DRAW.GraphicData graphicData = null;
                        Size frame = Utils.GetShapeSize(drawing.Descendants<DRAW.Extents>().FirstOrDefault());
                        if (frame == null) frame = Utils.GetImageSize(stream);
                        if (frame != null)
                        {
                            Run? pRun = drawing.Ancestors<Run>().FirstOrDefault();
                            if (pRun != null)
                            {
                                graphicData = drawing?.Descendants<DRAW.GraphicData>().FirstOrDefault();
                                if (graphicData != null)
                                {
                                    graphicData.RemoveAllChildren();
                                    graphicData.Uri = Constant.PICTURE_NAMESPACE;
                                }
                            }
                        }

                        elements.ForEach(element =>
                        {
                            if (element is PIC.Picture)
                            {
                                var imagePart = Utils.AddImagePart(bookmark.ParentPart);
                                var imageId = bookmark.ParentPart.GetIdOfPart(imagePart);
                                imagePart.FeedData(stream);
                                Utils.UpdateImageIdAndSize(element, imageId, frame);
                                if (graphicData != null)
                                {
                                    graphicData.Append(element);
                                }
                                else if(run != null)
                                {
                                    run.Append(element);
                                }
                            }
                            else if (run != null)
                            {
                                run.Append(element);
                            }
                        });
                    }
                    else
                    {
                        elements.ForEach(element =>
                        {
                            if (element is PIC.Picture)
                            {
                                var frame = Utils.GetImageSize(stream);
                                if (frame != null)
                                {
                                    var imagePart = Utils.AddImagePart(bookmark.ParentPart);
                                    var imageId = bookmark.ParentPart.GetIdOfPart(imagePart);
                                    imagePart.FeedData(stream);
                                    Utils.UpdateImageIdAndSize(element, imageId, frame);
                                    element = Utils.CreateNewDrawingElement(element, frame);
                                }
                            }
                            bookmark.BookmarkStart.InsertBeforeSelf(element);
                        });
                    }
                }
            }
        }
    }

    public class BarCodeReplacer : ImageReplacer, IReplacer
    {
        public override object FormatData(object data)
        {
            if (data != null)
            {
                return Utils.GetBarCodeImage(data.ToString());
            }
            return null;
        }
    }

    public class QRCodeReplacer : ImageReplacer, IReplacer
    {
        public override object FormatData(object data)
        {
            if (data != null)
            {
                return Utils.GetQRCodeImage(data.ToString());
            }
            return null;
        }
    }

    public class HTMLReplacer : BaseReplacer, IReplacer
    {
        public override List<OpenXmlElement> GenerateElements(WordprocessingDocument document, Bookmark bookmark)
        {
            var elements = new List<OpenXmlElement>();
            if (FormatedData != null)
            {
                MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(string.Format(Constant.HTML_PATTERN, FormatedData)));
                AlternativeFormatImportPart formatImportPart = null;
                if (bookmark.ParentPart is MainDocumentPart)
                    formatImportPart = ((MainDocumentPart)bookmark.ParentPart).AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Html);
                else if (bookmark.ParentPart is HeaderPart)
                    formatImportPart = ((HeaderPart)bookmark.ParentPart).AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Html);
                else if (bookmark.ParentPart is FooterPart)
                    formatImportPart = ((FooterPart)bookmark.ParentPart).AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Html);

                if (formatImportPart != null)
                {
                    formatImportPart.FeedData(stream);
                    AltChunk altChunk = new AltChunk();
                    altChunk.Id = bookmark.ParentPart.GetIdOfPart(formatImportPart);
                    elements.Add(new Run(altChunk));
                }
                stream.Dispose();                
            }
            return elements;
        }
    }

    public class DocumentReplacer : BaseReplacer, IReplacer
    {
        public override object FormatData(object data)
        {
            if(data != null)
            {
                Stream stream = new MemoryStream(Convert.FromBase64String(data.ToString()));
                return WordprocessingDocument.Open(stream, false);
            }
            return null;
        }

        public override void InsertElements(List<OpenXmlElement> elements, WordprocessingDocument document, Bookmark bookmark)
        {
            if (FormatedData is WordprocessingDocument)
            {
                var startNodeToInsert = bookmark.BookmarkStart.Ancestors().FirstOrDefault(a => a.Parent == document.MainDocumentPart.Document.Body);
                if (startNodeToInsert != null)
                {
                    var wordDocument = (WordprocessingDocument)FormatedData;
                    Dictionary<string, string> mappingRID = new Dictionary<string, string>();
                    foreach (var p in wordDocument.MainDocumentPart.Parts)
                    {
                        //ignore header and footer data
                        if (p.OpenXmlPart is HeaderPart or FooterPart)
                        {
                            continue;
                        }

                        try
                        {
                            var rId = Utils.GetUniqueStringID();
                            document.MainDocumentPart.AddPart(p.OpenXmlPart, rId);
                            mappingRID.Add(p.RelationshipId, rId);
                        }
                        catch { }
                    }

                    foreach (var el in wordDocument.MainDocumentPart.Document.Body.Elements())
                    {
                        if (el is SectionProperties) continue;
                        var newEl = el.CloneNode(true);
                        var subEls = newEl.Descendants().Where(x => { return x.GetAttributes().Where(a => mappingRID.ContainsKey(a.Value)).FirstOrDefault().LocalName != null; });

                        foreach (var x in subEls)
                        {
                            var att = x.GetAttributes().Where(a => mappingRID.ContainsKey(a.Value)).FirstOrDefault();
                            var newAttr = new OpenXmlAttribute(att.Prefix, att.LocalName, att.NamespaceUri, mappingRID[att.Value]);
                            x.SetAttribute(newAttr);
                        }

                        if (startNodeToInsert is Paragraph && newEl is Paragraph)
                        {
                            var paraProp = startNodeToInsert.Elements<ParagraphProperties>().FirstOrDefault();
                            var oldParaProp = newEl.Elements<ParagraphProperties>().FirstOrDefault();
                            if (paraProp != null)
                            {
                                var newParaProp = new ParagraphProperties();
                                if (oldParaProp != null)
                                {
                                    foreach (var item in oldParaProp.ChildElements)
                                    {
                                        if (!(item is Indentation))
                                            newParaProp.Append(item.CloneNode(true));
                                    }

                                    oldParaProp.Remove();
                                }

                                foreach (var item in paraProp.ChildElements)
                                {
                                    if (!newParaProp.Elements().Any(x => x.GetType() == item.GetType()))
                                        newParaProp.Append(item.CloneNode(true));
                                }

                                newEl.InsertAt(newParaProp, 0);
                            }
                        }

                        startNodeToInsert.InsertBeforeSelf(newEl);
                    }

                    wordDocument.Dispose();
                }
            }
        }
    }
}
