using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocTool.DocType
{
    public class PDFText
    {
        public PDFText(string newText, float fontSize = 12, float offsetRightX = 0, float offsetRightY = 0)
        {
            this.newText = newText;
            this.fontSize = fontSize;
            this.offsetRightX = offsetRightX;
            this.offsetRightY = offsetRightY;
        }
        public string newText { get; set; }
        public float fontSize { get; set; }
        public float offsetRightX { get; set; }
        public float offsetRightY { get; set; }
    }
}
