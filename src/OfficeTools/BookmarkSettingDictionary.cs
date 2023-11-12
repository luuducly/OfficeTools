namespace OfficeTools
{
    public class BookmarkSettingDictionary
    {
        public BookmarkSetting this[string key]
        {
            get
            {
                if (_inputSettings.ContainsKey(key)) return _inputSettings[key];
                var newItem = new BookmarkSetting();
                _inputSettings[key] = newItem;
                return newItem;
            }
            set { _inputSettings[key] = value; }
        }
        private Dictionary<string, BookmarkSetting> _inputSettings;

        internal IEnumerable<KeyValuePair<string, BookmarkSetting>> BookmarkSettings
        {
            get
            {
                return _inputSettings.Select(s => new KeyValuePair<string, BookmarkSetting>(s.Key, s.Value));
            }
        }

        internal IEnumerable<KeyValuePair<string, Bookmark>> Bookmarks
        {
            get
            {
                return _inputSettings.Select(s => new KeyValuePair<string, Bookmark>(s.Key, s.Value.Bookmark));
            }
        }

        internal BookmarkSettingDictionary()
        {
            _inputSettings = new Dictionary<string, BookmarkSetting>(StringComparer.OrdinalIgnoreCase);
        }

        internal bool ContainsKey(string key)
        {
            return _inputSettings.ContainsKey(key);
        }

        internal void Remove(string key) {
            if(_inputSettings.ContainsKey(key))
                _inputSettings.Remove(key);
        }
    }
}
