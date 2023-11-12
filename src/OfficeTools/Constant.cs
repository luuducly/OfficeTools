using SkiaSharp;

namespace OfficeTools
{
    internal static class Constant
    {
        public const string HTML_PATTERN = "<html><head><meta charset=\"UTF-8\"></head><body>{0}</body></html>";
        public const string PICTURE_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/picture";
        public const long PIXEL_PER_INCH = 914400L;
        public const double DEFAULT_DPI = 95.9865952;
        public const int DEFAULT_BARCODE_HIGHT = 70;
        public const int DEFAULT_QRCODE_SIZE = 512;
        public static readonly SKColor DEFAULT_DARK_COLOR = SKColor.Parse("000000");
        public static readonly SKColor DEFAULT_LIGHT_COLOR = SKColor.Empty;
    }
}
