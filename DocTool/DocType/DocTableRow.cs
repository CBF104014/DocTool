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
                var rowData = SelfTable.CreateRow();
                foreach (var cellItem in rowItem)
                {
                    rowData.Append(SelfTable.CreateCell(new DocTableCellProp(cellItem)));
                }
                RowDatas.Add(rowData);
            }
        }
    }
}
