using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocTool.DocType
{
    public class DocTableRow
    {
        private DocTable SelfTable = new DocTable();
        public List<TableRow> RowDatas { get; set; } = new List<TableRow>();
        public DocTableRow() { }
        public DocTableRow(List<List<string>> sourceDatas)
        {
            foreach (var rowItem in sourceDatas)
            {
                var rowData = CreateRow();
                foreach (var cellItem in rowItem)
                {
                    rowData.Append(CreateCell(new DocTableCellProp(cellItem)));
                }
                RowDatas.Add(rowData);
            }
        }
        public TableRow CreateRow(int rowHeight = 0)
        {
            return SelfTable.CreateRow(rowHeight);
        }
        public TableCell CreateCell(DocTableCellProp cellProp)
        {
            return SelfTable.CreateCell(cellProp);
        }
    }
}
