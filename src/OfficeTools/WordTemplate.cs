using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json.Linq;

namespace OfficeTools
{
    public class WordTemplate : IDisposable
    {
        public BookmarkSettingDictionary Bookmarks
        {
            get
            {
                return _bmSettingDictionary;
            }
        }
        private BookmarkSettingDictionary _bmSettingDictionary;

        private Stream _sourceStream;

        public WordTemplate(Stream sourceStream)
        {
            if (sourceStream == null)
            {
                throw new ArgumentNullException(nameof(sourceStream));
            }
            _sourceStream = sourceStream;
            _bmSettingDictionary = new BookmarkSettingDictionary();
        }

        public WordprocessingDocument Export(object data)
        {
            MemoryStream targetStream = new MemoryStream();
            _sourceStream.Position = 0;
            _sourceStream.CopyTo(targetStream);
            WordprocessingDocument targetDocument = WordprocessingDocument.Open(targetStream, true);
            if (data != null)
            {
                //TODO: support docx for word 2007
                Utils.RemoveFallbackElements(targetDocument);
                PrepareBookmarkSettings(targetDocument);

                if (data is not JObject) data = JObject.FromObject(data);
                FillData(targetDocument, data as JObject);
                targetDocument.Save();
            }
            return targetDocument;
        }

        private void PrepareBookmarkSettings(WordprocessingDocument document)
        {
            FindBookmarkNodes(document);
            List<string> redanduntSettings = new List<string>();

            foreach(var item in _bmSettingDictionary.BookmarkSettings)
            {
                var bmSetting = item.Value;
                if (bmSetting.Bookmark.BookmarkStart == null || bmSetting.Bookmark.BookmarkStart == null)
                {
                    redanduntSettings.Add(item.Key);
                    continue;
                }    
                    
                if (bmSetting.Replacer == null)
                    bmSetting.Replacer = Utils.GetDefaultFormatter(bmSetting.DataType);
            }

            foreach(var key in redanduntSettings)
            {
                _bmSettingDictionary.Remove(key);
            }
        }

        private void FindBookmarkNodes(WordprocessingDocument document)
        {
            MainDocumentPart mainPart = document.MainDocumentPart;
            Document documentPart = mainPart.Document;

            //find bookmark templates in body part
            foreach (var bmStart in documentPart.Body.Descendants<BookmarkStart>())
            {
                var bmEnd = Utils.FindBookmarkEnd(bmStart);
                if (bmEnd != null)
                {
                    var bookmarkSetting = _bmSettingDictionary[bmStart.Name];
                    Bookmark bookmark = new Bookmark();
                    bookmark.ParentPart = mainPart;
                    bookmark.BookmarkStart = bmStart;
                    bookmark.BookmarkEnd = bmEnd;
                    bookmarkSetting.Bookmark = bookmark;
                    Utils.CorrectBookmarkTagsPosition(bookmark);
                }
            }

            //find bookmark templates in header parts
            foreach (HeaderPart headerPart in mainPart.HeaderParts)
            {
                foreach (var bmStart in headerPart.Header.Descendants<BookmarkStart>())
                {
                    var bmEnd = Utils.FindBookmarkEnd(bmStart);
                    if (bmEnd != null)
                    {
                        var bookmarkSetting = _bmSettingDictionary[bmStart.Name];
                        Bookmark bookmark = new Bookmark();
                        bookmark.ParentPart = headerPart;
                        bookmark.BookmarkStart = bmStart;
                        bookmark.BookmarkEnd = bmEnd;
                        bookmarkSetting.Bookmark = bookmark;
                        Utils.CorrectBookmarkTagsPosition(bookmark);
                    }
                }
            }

            //find bookmark templates in footer parts
            foreach (FooterPart footerPart in mainPart.FooterParts)
            {
                foreach (var bmStart in footerPart.Footer.Descendants<BookmarkStart>())
                {
                    var bmEnd = Utils.FindBookmarkEnd(bmStart);
                    if (bmEnd != null)
                    {
                        var bookmarkSetting = _bmSettingDictionary[bmStart.Name];
                        Bookmark bookmark = new Bookmark();
                        bookmark.ParentPart = footerPart;
                        bookmark.BookmarkStart = bmStart;
                        bookmark.BookmarkEnd = bmEnd;
                        bookmarkSetting.Bookmark = bookmark;
                        Utils.CorrectBookmarkTagsPosition(bookmark);
                    }
                }
            }
        }

        private void UpdateBookmarkNodes(List<OpenXmlElement> elements, params string[] bookmarkNames)
        {
            if (bookmarkNames == null) return;

            foreach (var bmn in bookmarkNames)
            {
                if (_bmSettingDictionary.ContainsKey(bmn))
                {
                    var bft = _bmSettingDictionary[bmn];
                    var bookmark = bft.Bookmark;
                    foreach (var el in elements)
                    {
                        var bmStart = el.Descendants<BookmarkStart>().Where(bm => string.Compare(bm.Name, bmn, StringComparison.OrdinalIgnoreCase) == 0).FirstOrDefault();
                        if (bmStart != null)
                        {
                            var bmEnd = Utils.FindBookmarkEnd(bmStart);
                            if (bmEnd != null)
                            {
                                bookmark.BookmarkStart = bmStart;
                                bookmark.BookmarkEnd = bmEnd;
                                break;
                            }
                        }
                    }
                }
            }
        }

        private void FillData(WordprocessingDocument document, JObject data)
        {
            foreach (var item in data)
            {
                var propName = item.Key;
                var propData = item.Value;

                //for repeat node
                if (propData is JArray)
                {
                    var arrData = (JArray)propData;
                    if (arrData.Count > 0)
                    {
                        var firstItem = arrData[0];
                        if (firstItem != null)
                        {
                            if (firstItem is JObject)
                            {
                                var repeatProperties = Utils.GetAllRepeatProperties(firstItem as JObject);
                                var bookmarks = _bmSettingDictionary.Bookmarks.Where(x => repeatProperties.Any(p => string.Compare(p, x.Key, true) == 0)).Select(x => x.Value).ToList();
                                var template = Utils.GetRepeatingTemplate(bookmarks);
                                int index = 0;
                                foreach (JObject subItem in arrData)
                                {
                                    FillData(document, subItem);
                                    if (++index < arrData.Count)
                                    {
                                        var newElements = template.CloneAndAppendTemplate();
                                        UpdateBookmarkNodes(newElements, repeatProperties.ToArray());
                                    }
                                }
                            }
                            else if (firstItem is JValue)
                            {
                                if (_bmSettingDictionary.ContainsKey(propName))
                                {
                                    var bookmarkSetting = _bmSettingDictionary[propName];
                                    var bookmarks = new List<Bookmark> { bookmarkSetting.Bookmark };
                                    var template = Utils.GetRepeatingTemplate(bookmarks);
                                    int index = 0;
                                    foreach (JValue subItem in arrData)
                                    {
                                        bookmarkSetting.Replacer.RawData = subItem.Value;
                                        bookmarkSetting.Replacer.FormatedData = bookmarkSetting.Replacer.FormatData(bookmarkSetting.Replacer.RawData);
                                        var elements = bookmarkSetting.Replacer.GenerateElements(document, bookmarkSetting.Bookmark);
                                        bookmarkSetting.Replacer.InsertElements(elements, document, bookmarkSetting.Bookmark);
                                        Utils.DeleteRedundantElements(bookmarkSetting.Bookmark);
                                        if (++index < arrData.Count)
                                        {
                                            var newElements = template.CloneAndAppendTemplate();
                                            UpdateBookmarkNodes(newElements, propName);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    if (_bmSettingDictionary.ContainsKey(propName))
                    {
                        var bookmarkSetting = _bmSettingDictionary[propName];
                        if (propData != null)
                        {
                            if (propData is JValue)
                            {
                                bookmarkSetting.Replacer.RawData = (propData as JValue).Value;
                                bookmarkSetting.Replacer.FormatedData = bookmarkSetting.Replacer.FormatData(bookmarkSetting.Replacer.RawData);
                                var elements = bookmarkSetting.Replacer.GenerateElements(document, bookmarkSetting.Bookmark);
                                bookmarkSetting.Replacer.InsertElements(elements, document, bookmarkSetting.Bookmark);
                            }
                            else
                            {
                                bookmarkSetting.Replacer.RawData = propData;
                                bookmarkSetting.Replacer.FormatedData = bookmarkSetting.Replacer.FormatData(bookmarkSetting.Replacer.RawData);
                                var elements = bookmarkSetting.Replacer.GenerateElements(document, bookmarkSetting.Bookmark);
                                bookmarkSetting.Replacer.InsertElements(elements, document, bookmarkSetting.Bookmark);
                            }
                            Utils.DeleteRedundantElements(bookmarkSetting.Bookmark);
                        }
                        else
                        {
                            Utils.DeleteRedundantElements(bookmarkSetting.Bookmark);
                        }
                    }
                }
            }
        }

        public void Dispose()
        {
            if (_sourceStream != null)
                _sourceStream.Dispose();
        }
    }
}