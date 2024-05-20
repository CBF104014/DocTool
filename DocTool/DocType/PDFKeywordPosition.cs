using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocTool.DocType
{
    public class PDFKeywordPosition
    {
        public int pageNum { get; set; }
        public string textStr { get; set; } = "";
        public float positionX { get; set; }
        public float positionY { get; set; }
    }
}
