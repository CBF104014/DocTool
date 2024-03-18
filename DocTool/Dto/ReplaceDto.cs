using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocTool.Dto
{
    public enum ReplaceType
    {
        //文字
        Text = 0,
        //表格
        Table = 1,
        //圖片
        Image = 2,
        //HTML
        HtmlString = 3,
        //表格-列
        TableRow = 4,
    }
    public class ReplaceDto : FileObj
    {
        public ReplaceType replaceType { get; set; }
        #region 文字
        public string textStr { get; set; }
        #endregion
        #region HTML
        public string htmlStr { get; set; }
        #endregion
        #region 表格
        public DocumentFormat.OpenXml.Wordprocessing.Table tableData { get; set; }
        public List<DocumentFormat.OpenXml.Wordprocessing.TableRow> tableRowDatas { get; set; }
        #endregion
        #region 圖片
        public double imageDpi { get; set; }
        public double imageWidth { get; set; }
        public double imageHeight { get; set; }
        #endregion
    }
}
