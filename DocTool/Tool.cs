using DocTool.Base;
using DocTool.Pdf;
using DocTool.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocTool
{
    public class Tool : IDocBase
    {
        public string libreOfficeAppPath { get; set; }
        public string locationTempPath { get; set; }
        public DocWordTool Word { get; set; }
        public DocPdfTool Pdf { get; set; }
        /// <summary>
        /// 傳入LibreOffice(soffice.exe)路徑
        /// </summary>
        public Tool(string libreOfficeAppPath, string locationTempPath = null)
        {
            this.libreOfficeAppPath = libreOfficeAppPath;
            this.locationTempPath = locationTempPath == null ? $@"C:\TEMP" : locationTempPath;
            Init();
        }
        public void Init()
        {
            this.Word = new DocWordTool(this.libreOfficeAppPath, this.locationTempPath);
            this.Pdf = new DocPdfTool();
        }
    }
}
