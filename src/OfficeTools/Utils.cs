using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using DRAW = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Newtonsoft.Json.Linq;
using SkiaSharp.QrCode;
using BarcodeStandard;
using SkiaSharp;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeTools
{
    internal class Utils
    {
        internal static IReplacer GetDefaultFormatter(DataType type)
        {
            switch (type)
            {
                case DataType.Image: return new ImageReplacer();
                case DataType.BarCode: return new BarCodeReplacer();
                case DataType.QRCode: return new QRCodeReplacer();
                case DataType.HTML: return new HTMLReplacer();
                case DataType.Document: return new DocumentReplacer();
                default: return new BaseReplacer();
            }
        }

        internal static void RemoveFallbackElements(WordprocessingDocument document)
        {
            MainDocumentPart mainPart = document.MainDocumentPart;
            Document documentPart = mainPart.Document;

            foreach (var alternative in documentPart.Body.Descendants<AlternateContent>())
            {
                var choice = alternative.Descendants<AlternateContentChoice>().FirstOrDefault();
                if(choice != null)
                {
                    var clonedNodes = choice.ChildElements.Select(x => x.CloneNode(true)).ToList();
                    clonedNodes.ForEach(node => alternative.InsertBeforeSelf(node));
                    alternative.Remove();
                }
            }

            foreach (HeaderPart headerPart in mainPart.HeaderParts)
            {
                foreach (var alternative in headerPart.Header.Descendants<AlternateContent>())
                {
                    var choice = alternative.Descendants<AlternateContentChoice>().FirstOrDefault();
                    if (choice != null)
                    {
                        var clonedNodes = choice.ChildElements.Select(x => x.CloneNode(true)).ToList();
                        clonedNodes.ForEach(node => alternative.InsertBeforeSelf(node));
                        alternative.Remove();
                    }
                }
            }

            foreach (FooterPart footerPart in mainPart.FooterParts)
            {
                foreach (var alternative in footerPart.Footer.Descendants<AlternateContent>())
                {
                    var choice = alternative.Descendants<AlternateContentChoice>().FirstOrDefault();
                    if (choice != null)
                    {
                        var clonedNodes = choice.ChildElements.Select(x => x.CloneNode(true)).ToList();
                        clonedNodes.ForEach(node => alternative.InsertBeforeSelf(node));
                        alternative.Remove();
                    }
                }
            }
        }

        internal static void CorrectBookmarkTagsPosition(Bookmark bookmark)
        {
            var currentNode = bookmark.BookmarkStart.NextSibling();
            if (currentNode == null) currentNode = bookmark.BookmarkStart.Parent?.NextSibling();
            while (currentNode != null)
            {
                if (currentNode is Run)
                {
                    bookmark.BookmarkStart.Remove();
                    currentNode.InsertBeforeSelf(bookmark.BookmarkStart);
                    break;
                }
                else
                {
                    var firstRun = currentNode.Descendants<Run>().FirstOrDefault();
                    if (firstRun != null)
                    {
                        bookmark.BookmarkStart.Remove();
                        firstRun.InsertBeforeSelf(bookmark.BookmarkStart);
                        break;
                    }
                }
                if (currentNode == bookmark.BookmarkEnd || currentNode.Descendants<BookmarkEnd>().Any(bme => bme == bookmark.BookmarkEnd)) break;

                if (currentNode.NextSibling() != null) currentNode = currentNode.NextSibling();
                else currentNode = currentNode.Parent?.NextSibling();
            }

            currentNode = bookmark.BookmarkEnd.PreviousSibling();
            if (currentNode == null) currentNode = bookmark.BookmarkEnd.Parent?.PreviousSibling();
            while (currentNode != null)
            {
                if (currentNode is Run)
                {
                    bookmark.BookmarkEnd.Remove();
                    currentNode.InsertAfterSelf(bookmark.BookmarkEnd);
                    break;
                }
                else
                {
                    var firstRun = currentNode.Descendants<Run>().LastOrDefault();
                    if (firstRun != null)
                    {
                        bookmark.BookmarkEnd.Remove();
                        firstRun.InsertAfterSelf(bookmark.BookmarkEnd);
                        break;
                    }
                }
                if (currentNode == bookmark.BookmarkStart || currentNode.Descendants<BookmarkStart>().Any(bme => bme == bookmark.BookmarkStart)) break;

                if (currentNode.PreviousSibling() != null) currentNode = currentNode.PreviousSibling();
                else currentNode = currentNode.Parent?.PreviousSibling(); ;
            }
        }

        internal static BookmarkTemplate GetRepeatingTemplate(List<Bookmark> bookmarks)
        {
            BookmarkTemplate bookmarkTemplate = new BookmarkTemplate();
            if (bookmarks.Count == 0) return bookmarkTemplate;

            //find the ascendant of both start and end bookmark node
            BookmarkStart bmStart = bookmarks[0].BookmarkStart;
            OpenXmlElement parentNode = bmStart.Parent;
            for (int i = 0; i < bookmarks.Count; i++)
            {
                while (parentNode != null && !parentNode.Descendants<BookmarkStart>().Any(el => el == bookmarks[i].BookmarkStart))
                {
                    parentNode = parentNode.Parent;
                }

                while (parentNode != null && !parentNode.Descendants<BookmarkEnd>().Any(el => el == bookmarks[i].BookmarkEnd))
                {
                    parentNode = parentNode.Parent;
                }
            }

            //only support repeat TableRow or Paragraph nodes
            if (parentNode is Paragraph or TableRow)
            {
                parentNode = parentNode.Parent;
            }

            if (parentNode != null)
            {
                OpenXmlElement lastChildNode = null;
                OpenXmlElement startNode = null, endNode = null;

                foreach (OpenXmlElement childNode in parentNode.ChildElements)
                {
                    if ((childNode is BookmarkStart && bookmarks.Any(bm => bm.BookmarkStart == childNode)) || childNode.Descendants<BookmarkStart>().Any(el => bookmarks.Any(bm => bm.BookmarkStart == el)))
                    {
                        startNode = childNode;
                        break;
                    }
                }

                foreach (OpenXmlElement childNode in parentNode.ChildElements.Reverse())
                {
                    if ((childNode is BookmarkEnd && bookmarks.Any(bm => bm.BookmarkEnd == childNode)) || childNode.Descendants<BookmarkEnd>().Any(el => bookmarks.Any(bm => bm.BookmarkEnd == el)))
                    {
                        endNode = childNode;
                        break;
                    }
                }

                if (startNode == endNode)
                {
                    bookmarkTemplate.TemplateElements.Add(startNode.CloneNode(true));
                    lastChildNode = endNode;
                }
                else
                {
                    var templateNode = startNode;
                    while (templateNode != null)
                    {
                        bookmarkTemplate.TemplateElements.Add(templateNode.CloneNode(true));
                        lastChildNode = templateNode;

                        if (templateNode == endNode) break;
                        templateNode = templateNode.NextSibling();
                    }
                }

                if (lastChildNode != null)
                {
                    if (lastChildNode.NextSibling() != null)
                        bookmarkTemplate.LastNode = lastChildNode.NextSibling();
                    else
                        bookmarkTemplate.ParentNode = lastChildNode.Parent;
                }
            }

            return bookmarkTemplate;
        }

        internal static ImagePart AddImagePart(TypedOpenXmlPart parentPart)
        {
            ImagePart imagePart = null;
            if (parentPart is HeaderPart)
            {
                imagePart = ((HeaderPart)parentPart).AddImagePart(ImagePartType.Png);
            }
            else if (parentPart is MainDocumentPart)
            {
                imagePart = ((MainDocumentPart)parentPart).AddImagePart(ImagePartType.Png);
            }
            else if (parentPart is FooterPart)
            {
                imagePart = ((FooterPart)parentPart).AddImagePart(ImagePartType.Png);
            }
            return imagePart;
        }

        internal static OpenXmlElement CreateNewDrawingElement(OpenXmlElement image, Size size)
        {
            string name = Guid.NewGuid().ToString();
            var element =
                new Drawing(
                    new Inline(
                        new Extent() { Cx = size.Width, Cy = size.Height },
                        new EffectExtent()
                        {
                            LeftEdge = 0L,
                            TopEdge = 0L,
                            RightEdge = 0L,
                            BottomEdge = 0L
                        },
                        new DocProperties()
                        {
                            Id = GetUintId(),
                            Name = name
                        },
                        new NonVisualGraphicFrameDrawingProperties(
                            new DRAW.GraphicFrameLocks() { NoChangeAspect = true }),
                        new DRAW.Graphic(
                            new DRAW.GraphicData(
                                    image
                                )
                            { Uri = Constant.PICTURE_NAMESPACE })
                    )
                    {
                        DistanceFromTop = 0U,
                        DistanceFromBottom = 0U,
                        DistanceFromLeft = 0U,
                        DistanceFromRight = 0U,
                        EditId = GetRandomHexNumber(8)
                    });
            return element;
        }

        internal static PIC.Picture CreateNewPictureElement(string fileName, long width, long height)
        {
            return new PIC.Picture(
                new PIC.NonVisualPictureProperties(
                    new PIC.NonVisualDrawingProperties()
                    {
                        Id = GetUintId(),
                        Name = fileName + ".png"
                    },
                    new PIC.NonVisualPictureDrawingProperties()),
                new PIC.BlipFill(
                    new DRAW.Blip(
                        new DRAW.BlipExtensionList()
                    )
                    {
                        CompressionState =
                            DRAW.BlipCompressionValues.Print
                    },
                    new DRAW.Stretch(
                        new DRAW.FillRectangle())),
                new PIC.ShapeProperties(
                    new DRAW.Transform2D(
                        new DRAW.Offset() { X = 0L, Y = 0L },
                        new DRAW.Extents() { Cx = width, Cy = height }),
                    new DRAW.PresetGeometry(
                            new DRAW.AdjustValueList()
                        )
                    { Preset = DRAW.ShapeTypeValues.Rectangle }));
        }

        internal static void UpdateImageIdAndSize(OpenXmlElement element, string imageId, Size size)
        {
            if (element != null )
            {
                var blip = element.Descendants<DRAW.Blip>().FirstOrDefault();
                if (blip != null) blip.Embed = imageId;
                var extents = element.Descendants<DRAW.Extents>().FirstOrDefault();
                if (extents != null)
                {
                    extents.Cx = size.Width;
                    extents.Cy = size.Height;
                }
            }
        }

        internal static void GenerateNewIdAndName(OpenXmlElement element)
        {
            if (element != null)
            {
                foreach (var drawing in element.Descendants<Drawing>())
                {
                    var prop = drawing.Descendants<DocProperties>().FirstOrDefault();
                    if (prop != null)
                    {
                        prop.Id = GetUintId();
                        prop.Name = GetUniqueStringID();
                    }
                }
            }
        }

        internal static Stream GetQRCodeImage(string data)
        {
            var generator = new QRCodeGenerator();
            var qr = generator.CreateQrCode(data, ECCLevel.L, quietZoneSize: 1);
            var info = new SKImageInfo(Constant.DEFAULT_QRCODE_SIZE, Constant.DEFAULT_QRCODE_SIZE);
            using var surface = SKSurface.Create(info);
            var canvas = surface.Canvas;
            canvas.Render(qr, info.Width, info.Height, Constant.DEFAULT_DARK_COLOR, Constant.DEFAULT_LIGHT_COLOR);
            using (var image = surface.Snapshot())
            {
                using (var imgData = image.Encode(SKEncodedImageFormat.Png, 100))
                {
                    Stream stream = new MemoryStream();
                    imgData.SaveTo(stream);
                    stream.Position = 0;
                    return stream;
                }
            }
        }

        internal static Stream? GetBarCodeImage(string data)
        {
            var barCode = new Barcode();
            barCode.Height = Constant.DEFAULT_BARCODE_HIGHT;
            var img = barCode.Encode(BarcodeStandard.Type.Code128, data, Constant.DEFAULT_DARK_COLOR, Constant.DEFAULT_LIGHT_COLOR);
            SKData encoded = img.Encode();
            var stream = encoded.AsStream();
            stream.Position = 0;
            return stream;
        }

        internal static Size GetShapeSize(DRAW.Extents extents)
        {
            if (extents != null)
            {
                Int64Value? w = extents.Cx;
                Int64Value? h = extents.Cy;
                if (w.HasValue && h.HasValue)
                    return new Size(w.Value, h.Value);
            }
            return null;
        }

        internal static Size GetImageSize(Stream stream)
        {
            stream.Position = 0;
            var image = SKImage.FromEncodedData(stream);
            stream.Position = 0;
            if (image != null)
            {
                var width = (long)(image.Width / Constant.DEFAULT_DPI * Constant.PIXEL_PER_INCH);
                var height = (long)(image.Height / Constant.DEFAULT_DPI * Constant.PIXEL_PER_INCH);
                return new Size(width, height);
            }
            return null;
        }

        internal static OpenXmlElement? CloneStyle(Bookmark bookmark)
        {
            RunProperties? runProperties = bookmark.BookmarkStart.NextSibling<Run>()
                                                 ?.Descendants<RunProperties>()
                                                 ?.FirstOrDefault();

            if (runProperties == null)
            {
                runProperties = bookmark.BookmarkEnd.PreviousSibling<Run>()
                                     ?.Descendants<RunProperties>()
                                     ?.FirstOrDefault();
            }

            if (runProperties == null)
            {
                runProperties = bookmark.BookmarkStart.Ancestors()
                                       .Select(a => a.Descendants<RunProperties>().FirstOrDefault())
                                       .FirstOrDefault();
            }

            return runProperties?.CloneNode(true);
        }

        internal static BookmarkEnd FindBookmarkEnd(BookmarkStart bmStart)
        {
            if (bmStart == null) throw new ArgumentException(nameof(bmStart));
            if (bmStart.Name == "_GoBack") return null;

            OpenXmlElement curentNode = bmStart.NextSibling();
            if (curentNode == null) curentNode = bmStart.Parent;
            while (curentNode != null)
            {
                //in case bookmart start node and end node are in the same level
                if (curentNode is BookmarkEnd && ((BookmarkEnd)curentNode).Id == bmStart.Id)
                {
                    return (BookmarkEnd)curentNode;
                }
                else //in case bookmark start and bookmark end's ascendant are in the same level
                {
                    var bmEnd = curentNode.Descendants<BookmarkEnd>().FirstOrDefault(end => end.Id == bmStart.Id);
                    if (bmEnd != null)
                    {
                        return bmEnd;
                    }
                }

                //move next
                if (curentNode.NextSibling() != null)
                {
                    curentNode = curentNode.NextSibling();
                }
                else //last node in the same level, move to find with parent level
                {
                    curentNode = curentNode.Parent;
                }
            }
            return null;
        }

        internal static List<string> GetAllRepeatProperties(JObject item)
        {
            if (item == null) return new List<string>();
            var repeatProperties = item.Properties().Select(p => p.Name).ToList();
            foreach (var subItem in item)
            {
                var subItemArr = subItem.Value as JArray;
                if (subItemArr != null)
                {
                    if (subItemArr.Count > 0)
                    {
                        repeatProperties.AddRange(GetAllRepeatProperties(subItemArr[0] as JObject));
                    }

                }
            }
            return repeatProperties;
        }

        internal static void DeleteRedundantElements(Bookmark bookmark)
        {
            OpenXmlElement curentNode = bookmark.BookmarkStart.NextSibling();
            while (curentNode != null)
            {
                if (curentNode == bookmark.BookmarkEnd)
                    break;

                if (curentNode is not (BookmarkStart or BookmarkEnd) && !curentNode.Descendants<BookmarkEnd>().Any(bm => bm == bookmark.BookmarkEnd))
                    curentNode.Remove();

                curentNode = curentNode.NextSibling();
            }

            curentNode = bookmark.BookmarkEnd.PreviousSibling();
            while (curentNode != null)
            {
                if (curentNode == bookmark.BookmarkStart)
                    break;

                if (curentNode is not (BookmarkStart or BookmarkEnd) && !curentNode.Descendants<BookmarkStart>().Any(bm => bm == bookmark.BookmarkStart))
                    curentNode.Remove();

                curentNode = curentNode.PreviousSibling();
            }

            bookmark.BookmarkStart.Remove();
            bookmark.BookmarkEnd.Remove();
        }

        internal static string GetRandomHexNumber(int digits)
        {
            Random random = new Random();
            byte[] buffer = new byte[digits / 2];
            random.NextBytes(buffer);
            string result = string.Concat(buffer.Select(x => x.ToString("X2")).ToArray());
            if (digits % 2 == 0)
                return result;
            return result + random.Next(16).ToString("X");
        }

        private static uint _initUint = 10000U;
        internal static uint GetUintId()
        {
            return _initUint++;
        }

        internal static string GetUniqueStringID()
        {
            return "r" + Guid.NewGuid().ToString().Replace("-", "");
        }
    }

    internal class Size
    {
        internal long Width { get; set; }
        internal long Height { get; set; }

        internal Size(long w, long h)
        {
            Width = w;
            Height = h;
        }
    }
}
