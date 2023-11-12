using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeTools
{
    public class Bookmark
    {
        public BookmarkStart BookmarkStart { get; internal set; }
        public BookmarkEnd BookmarkEnd { get; internal set; }
        public TypedOpenXmlPart ParentPart { get; internal set; }

        internal Bookmark()
        {
        }
    }
}
