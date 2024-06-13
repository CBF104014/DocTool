using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocTool.DocType
{
    public class DocTableCellProp
    {
        /// <summary>
        /// 建構子1
        /// </summary>
        public DocTableCellProp(object cellObj, JustificationValues HAlign, TableVerticalAlignmentValues VAlign, decimal fontSize = 24, int colSpan = 1, int rowSpan = 1, bool isBold = false, string textColor = "000000", double cellWidthCM = 0, string bgColor = "", double topMargin = 0, double startMargin = 0, double bottomMargin = 0, double endMargin = 0)
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
            this.bgColor = bgColor;
            this.bgColor = bgColor;
            this.topMargin = topMargin;
            this.startMargin = startMargin;
            this.bottomMargin = bottomMargin;
            this.endMargin = endMargin;
        }
        /// <summary>
        /// 建構子2
        /// </summary>
        public DocTableCellProp(object cellObj) 
            : this(cellObj: cellObj,
                  HAlign: JustificationValues.Left,
                  VAlign: TableVerticalAlignmentValues.Top,
                  fontSize: 24)
        { }
        /// <summary>
        /// 建構子3
        /// </summary>
        public DocTableCellProp(object cellObj, decimal fontSize = 24, int colSpan = 1, int rowSpan = 1, bool isBold = false, string textColor = "000000", double cellWidthCM = 0, string bgColor = "", double topMargin = 0, double startMargin = 0, double bottomMargin = 0, double endMargin = 0)
            : this(cellObj: cellObj,
                  HAlign: JustificationValues.Left,
                  VAlign: TableVerticalAlignmentValues.Top,
                  fontSize: fontSize,
                  colSpan: colSpan,
                  rowSpan: rowSpan,
                  isBold: isBold,
                  textColor: textColor,
                  cellWidthCM: cellWidthCM,
                  bgColor: bgColor,
                  topMargin: topMargin,
                  startMargin: startMargin,
                  bottomMargin: bottomMargin,
                  endMargin: endMargin)
        { }
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
        /// <summary>
        /// 背景顏色
        /// </summary>
        public string bgColor { get; set; }
        /// <summary>
        /// 上邊界(單位Dxa)
        /// </summary>
        public double topMargin { get; set; }
        /// <summary>
        /// 左邊界(單位Dxa)
        /// </summary>
        public double startMargin { get; set; }
        /// <summary>
        /// 下邊界(單位Dxa)
        /// </summary>
        public double bottomMargin { get; set; }
        /// <summary>
        /// 右邊界(單位Dxa)
        /// </summary>
        public double endMargin { get; set; }
    }
}
