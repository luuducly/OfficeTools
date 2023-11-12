namespace OfficeTools
{
    public class BookmarkSetting
    {
        public DataType DataType { get; set; }        
        public IReplacer Replacer { get; set; }

        internal Bookmark Bookmark;
    }
}
