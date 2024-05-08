using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocTool.DocType
{
    public class DocTable : DocumentFormat.OpenXml.Wordprocessing.Table
    {
        /// <summary>
        /// 圖片資源集合
        /// </summary>
        public Dictionary<string, DocImage> DocImageDict { get; set; } = new Dictionary<string, DocImage>();
        /// <summary>
        /// 新增欄
        /// </summary>
        public TableRow CreateRow()
        {
            var dataRow = new TableRow();
            return dataRow;
        }
        /// <summary>
        /// 新增儲存格
        /// </summary>
        public TableCell CreateCell(DocTableCellPrpo cellPrpo)
        {
            var cell = new TableCell();
            var paragraph = new Paragraph();
            var run = new Run();
            var runProperties = new RunProperties();
            runProperties.Append(new RunFonts() { Ascii = cellPrpo.fontName });
            runProperties.Append(new FontSize() { Val = cellPrpo.fontSize.ToString() });
            runProperties.Append(new Color() { Val = cellPrpo.fontColor });
            if (cellPrpo.isBold)
                runProperties.Append(new Bold());
            if (cellPrpo.cellObj.GetType() == typeof(DocTable))
            {
                var table = (DocTable)cellPrpo.cellObj;
                if (table.DocImageDict!= null)
                {
                    foreach (var item in table.DocImageDict)
                    {
                        if (!this.DocImageDict.ContainsKey(item.Key))
                        {
                            this.DocImageDict.Add(item.Key, item.Value);
                        }
                    }
                }
                run.Append(runProperties);
                run.Append(table);
            }
            else if (cellPrpo.cellObj.GetType() == typeof(DocImage))
            {
                var image = (DocImage)cellPrpo.cellObj;
                runProperties.AppendChild(new Highlight() { Val = HighlightColorValues.Yellow });
                run.Append(runProperties);
                run.Append(new Text(image.imageName));
                this.DocImageDict.Add(image.imageName, image);
            }
            else
            {
                run.Append(runProperties);
                run.Append(new Text(cellPrpo.cellObj.ToString()));
            }
            paragraph.Append(run);
            cell.Append(paragraph);
            var cellProperties = new TableCellProperties();
            //寬度
            if (cellPrpo.cellWidthCM > 0)
            {
                cellProperties.Append(new TableCellWidth() { Width = CMToDXA(cellPrpo.cellWidthCM), Type = TableWidthUnitValues.Dxa });
            }
            //水平合併(開頭為Restart，其他為Continue)
            if (cellPrpo.colSpan > 1 || cellPrpo.colSpan == 0)
            {
                cellProperties.Append(new HorizontalMerge() { Val = cellPrpo.colSpan > 1 ? MergedCellValues.Restart : MergedCellValues.Continue });
            }
            //垂直合併(開頭為Restart，其他為Continue)
            if (cellPrpo.rowSpan > 1 || cellPrpo.rowSpan == 0)
            {
                cellProperties.Append(new VerticalMerge() { Val = cellPrpo.rowSpan > 1 ? MergedCellValues.Restart : MergedCellValues.Continue });
            }
            cell.Append(cellProperties);
            return cell;
        }
        /// <summary>
        /// 轉XML Table
        /// </summary>
        public DocTable ToXMLTable(List<string> tableHead, List<List<string>> tableRows)
        {
            var row = CreateRow();
            foreach (var item in tableHead)
            {
                var cell = CreateCell(new DocTableCellPrpo(item, JustificationValues.Center, TableVerticalAlignmentValues.Top, isBold: true));
                row.Append(cell);
            }
            this.Append(row);
            foreach (var rowItem in tableRows)
            {
                row = CreateRow();
                foreach (var item in rowItem)
                {
                    var cell = CreateCell(new DocTableCellPrpo(item, JustificationValues.Left, TableVerticalAlignmentValues.Top, isBold: false));
                    row.Append(cell);
                }
                this.Append(row);
            }
            return this;
        }
        /// <summary>
        /// 公分轉DXA
        /// </summary>
        public static string CMToDXA(double centimeters)
        {
            var inches = centimeters / 2.54;
            var twip = (int)(inches * 72.0 * 20);
            return twip.ToString();
        }
    }
}
