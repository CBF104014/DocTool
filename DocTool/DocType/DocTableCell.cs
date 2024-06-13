using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocTool.DocType
{
    public class DocTableCell
    {
        private DocTable SelfTable = new DocTable();
        /// <summary>
        /// 無資料時，是否刪除整欄
        /// </summary>
        public bool IsEmptyRemoveRow { get; set; }
        public List<TableCell> CellDatas { get; set; } = new List<TableCell>();
        public DocTableCell() { }
        public DocTableCell(List<string> sourceDatas)
        {
            foreach (var cellItem in sourceDatas)
            {
                var cellData = CreateCell(new DocTableCellProp(cellItem));
                CellDatas.Add(cellData);
            }
        }
        public TableCell CreateCell(DocTableCellProp cellProp)
        {
            return SelfTable.CreateCell(cellProp);
        }
    }
}
