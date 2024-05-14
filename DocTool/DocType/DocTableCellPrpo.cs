using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocTool.DocType
{
    public class DocTableCellPrpo
    {
        public DocTableCellPrpo(object cellObj)
        {
            this.cellObj = cellObj == null ? "" : cellObj;
            this.fontSize = 24;
            this.colSpan = 1;
            this.rowSpan = 1;
            this.isBold = false;
            this.fontColor = "000000";
            this.HAlign = JustificationValues.Left;
            this.VAlign = TableVerticalAlignmentValues.Top;
            this.fontName = "標楷體";
            this.cellWidthCM = 0;
        }
        public DocTableCellPrpo(object cellObj, JustificationValues HAlign, TableVerticalAlignmentValues VAlign, decimal fontSize = 24, int colSpan = 1, int rowSpan = 1, bool isBold = false, string textColor = "000000", double cellWidthCM = 0)
        {
            this.cellObj = cellObj == null ? "" : cellObj;
            this.fontSize = fontSize;
            this.colSpan = colSpan;
            this.rowSpan = rowSpan;
            this.isBold = isBold;
            this.fontColor = String.IsNullOrEmpty(textColor) ? "000000" : textColor;
            this.HAlign = HAlign == null ? JustificationValues.Left : HAlign;
            this.VAlign = VAlign == null ? TableVerticalAlignmentValues.Top: VAlign;
            this.fontName = "標楷體";
            this.cellWidthCM = 0;
        }
        public DocTableCellPrpo(object cellObj, decimal fontSize = 24, int colSpan = 1, int rowSpan = 1, bool isBold = false, string textColor = "000000", double cellWidthCM = 0)
        {
            this.cellObj = cellObj == null ? "" : cellObj;
            this.fontSize = fontSize;
            this.colSpan = colSpan;
            this.rowSpan = rowSpan;
            this.isBold = isBold;
            this.fontColor = String.IsNullOrEmpty(textColor) ? "000000" : textColor;
            this.HAlign = JustificationValues.Left;
            this.VAlign = TableVerticalAlignmentValues.Top;
            this.fontName = "標楷體";
            this.cellWidthCM = 0;
        }
        public object cellObj { get; set; }
        /// <summary>
        /// 字體大小
        /// </summary>
        public decimal fontSize { get; set; }
        /// <summary>
        /// 水平合併數
        /// </summary>
        public int colSpan { get; set; }
        /// <summary>
        /// 垂直合併數
        /// </summary>
        public int rowSpan { get; set; }
        /// <summary>
        /// 是否粗體
        /// </summary>
        public bool isBold { get; set; }
        /// <summary>
        /// 字體顏色
        /// </summary>
        public string fontColor { get; set; }
        /// <summary>
        /// 水平對齊
        /// </summary>
        public JustificationValues HAlign { get; set; }
        /// <summary>
        /// 垂直對齊
        /// </summary>
        public TableVerticalAlignmentValues VAlign { get; set; }
        /// <summary>
        /// 字體名稱
        /// </summary>
        public string fontName { get; set; }
        /// <summary>
        /// 寬度公分
        /// ※注意：表格本身也要指定寬度才行
        /// </summary>
        public double cellWidthCM { get; set; }
    }
}
