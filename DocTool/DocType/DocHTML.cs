using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocTool.DocType
{
    public class DocHTML
    {
        public string HTMLStr { get; set; }
        public DocHTML(string HTMLStr)
        {
            this.HTMLStr = HTMLStr;
        }
        public OpenXmlElement ToOpenXml()
        {
            var paragraph = new Paragraph();
            var run = new Run();
            run.Append(new Text(this.HTMLStr));
            paragraph.Append(run);
            return paragraph;
        }
    }
}
