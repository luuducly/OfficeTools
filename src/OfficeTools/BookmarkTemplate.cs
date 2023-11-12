using DocumentFormat.OpenXml;

namespace OfficeTools
{
    internal class BookmarkTemplate
    {
        internal OpenXmlElement LastNode;
        internal OpenXmlElement ParentNode;
        internal List<OpenXmlElement> TemplateElements;

        internal BookmarkTemplate()
        {
            TemplateElements = new List<OpenXmlElement>();
        }

        internal List<OpenXmlElement> CloneAndAppendTemplate()
        {
            var cloneElements = CloneTemplateElements();
            if (LastNode != null)
            {
                foreach (var el in cloneElements)
                {
                    Utils.GenerateNewIdAndName(el);
                    LastNode.InsertBeforeSelf(el);
                }
            }
            else if (ParentNode != null)
            {
                foreach (var el in cloneElements)
                {
                    Utils.GenerateNewIdAndName(el);
                    ParentNode.Append(el);
                }
            }
            return cloneElements;
        }

        private List<OpenXmlElement> CloneTemplateElements()
        {
            List<OpenXmlElement> templateElements = new List<OpenXmlElement>();
            foreach (OpenXmlElement templateElement in TemplateElements)
            {
                templateElements.Add(templateElement.CloneNode(true));
            }
            return templateElements;
        }
    }
}
