using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml;

namespace OfficeTools
{
    internal class RenderContext
    {
        internal string FieldName { get; set; }
        internal MergeField MergeField { get; set; }
        internal IEvaluator Evaluator { get; set; }
        internal string Parameters { get; set; }
        internal string Operator { get; set; }
        internal RenderContext Parent { get; set; }
        internal int Index { get; set; }
        internal List<RenderContext> ChildNodes { get; set; }


        public RenderContext()
        {
            ChildNodes = new List<RenderContext>();
        }
    }

    internal class MergeField
    {
        internal FieldCode StartNode { get; set; }
        internal FieldCode EndNode { get; set; }
        internal TypedOpenXmlPart ParentPart { get; set; }
    }

    internal class MergeFieldTemplate
    {
        internal OpenXmlElement LastNode;
        internal OpenXmlElement ParentNode;
        internal List<OpenXmlElement> TemplateElements;

        internal MergeFieldTemplate()
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
                    GenerateNewIdAndName(el);
                    LastNode.InsertBeforeSelf(el);
                }
            }
            else if (ParentNode != null)
            {
                foreach (var el in cloneElements)
                {
                    GenerateNewIdAndName(el);
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

        private void GenerateNewIdAndName(OpenXmlElement element)
        {
            if (element != null)
            {
                foreach (var drawing in element.Descendants<Drawing>())
                {
                    var prop = drawing.Descendants<DocProperties>().FirstOrDefault();
                    if (prop != null)
                    {
                        prop.Id = Utils.GetUintId();
                        prop.Name = Utils.GetUniqueStringID();
                    }
                }
            }
        }
    }
}
