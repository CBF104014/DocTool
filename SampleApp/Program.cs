using DocTool;
using DocTool.DocType;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace SampleApp
{
    public class Program
    {
        /// <summary>
        /// 主程式
        /// 範例說明：
        /// 1. 確保LibreOffice路徑正確
        /// 2. 請先在C:\TEMP建立「Word_hightlight測試.docx」與「barcode.jpg」檔案，可從範例SampleFile複製
        /// </summary>
        static void Main(string[] args)
        {
            var docData = new MyDocClass();
            var docTool = new Tool(@"E:\PortableApps\LibreOfficePortable\App\libreoffice\program\soffice.exe", docData.AppYData.outFilePath);
            //輸出WORD
            var fileData = docTool.Word
                .Set(docData.AppYData.FileDocPath)
                .ReplaceTag(docData)
                .GetData();
            System.IO.File.WriteAllBytes(Path.Combine(docData.AppYData.outFilePath, $"{DateTime.Now.ToString("yyyyMMddHHmmss")}{fileData.fileName}.{fileData.fileType}"), fileData.fileByteArr);
            //輸出PDF
            fileData = docTool.Word
                .ToPDF()
                .GetData();
            fileData = docTool.Pdf
                .Set(fileData)
                .AddText("=>", new PDFText("iTextSharp 加上去的文字!", offsetRightX: 20))
                .AddImage("=>", new PDFImage(docData.AppYData.FileImgPath, offsetRightX: 20, offsetRightY: -40))
                .SetReadOnly()
                .GetData();
            System.IO.File.WriteAllBytes(Path.Combine(docData.AppYData.outFilePath, $"{DateTime.Now.ToString("yyyyMMddHHmmss")}{fileData.fileName}.{fileData.fileType}"), fileData.fileByteArr);
        }
    }
    /// <summary>
    /// 範例-原始資料-細項
    /// </summary>
    public class AppP
    {
        public string ProductName { get; set; }
        public decimal Quantity { get; set; }
        public decimal Amount { get; set; }
        public decimal TotalAmount { get => this.Quantity * this.Amount; }
    }
    /// <summary>
    /// 範例-原始資料
    /// </summary>
    public class AppY
    {
        public string MyText1 { get; set; } = $"{DateTime.Now.ToString("yyyyMMdd HH:mm:ss")}-Test";
        public string PageEndText1 { get; set; } = $"這是頁尾";
        public string outFilePath { get => "C:\\TEMP"; }
        public string FileDocName { get => "Word_hightlight測試.docx"; }
        public string FileDocPath { get => Path.Combine(outFilePath, this.FileDocName); }
        public string FileImgName { get => "barcode.jpg"; }
        public string FileImgPath { get => Path.Combine(outFilePath, this.FileImgName); }
    }
    /// <summary>
    /// 範例-處理後資料
    /// </summary>
    public class MyDocClass
    {
        public AppY AppYData { get; set; } = new AppY();
        public List<AppP> AppPDatas { get; set; } = new List<AppP>
        {
            new AppP() { ProductName = "品項A", Quantity = 3, Amount = 1000 },
            new AppP() { ProductName = "品項B", Quantity = 77, Amount = 299 },
            new AppP() { ProductName = "品項C", Quantity = 30, Amount = 130 },
        };
        public string MyText1 { get => this.AppYData.MyText1; }
        public string PageEndText1 { get => this.AppYData.PageEndText1; }
        public decimal Total { get => this.AppPDatas == null ? 0 : this.AppPDatas.Sum(x => x.TotalAmount); }
        public string Total_str { get => this.Total.ToString("#,#.##"); }
        public DocTable Table1
        {
            get
            {
                var docTable = new DocTable();
                var dataRow = docTable.CreateRow();
                var imgData = new DocImage(this.AppYData.FileImgPath, imageDpi: 300, imageWidth: 600, imageHeight: 100);
                dataRow.Append(docTable.CreateCell(new DocTableCellProp("欄位A")));
                dataRow.Append(docTable.CreateCell(new DocTableCellProp("欄位B", JustificationValues.Center, TableVerticalAlignmentValues.Center)));
                dataRow.Append(docTable.CreateCell(new DocTableCellProp(imgData, JustificationValues.Center, TableVerticalAlignmentValues.Center, colSpan: 2, topMargin: 50)));
                dataRow.Append(docTable.CreateCell(new DocTableCellProp("", colSpan: 0)));
                docTable.Append(dataRow);
                foreach (var item in this.AppPDatas)
                {
                    dataRow = docTable.CreateRow();
                    dataRow.Append(docTable.CreateCell(new DocTableCellProp(item.ProductName)));
                    dataRow.Append(docTable.CreateCell(new DocTableCellProp(item.Quantity.ToString("#,#.##"))));
                    dataRow.Append(docTable.CreateCell(new DocTableCellProp(item.Amount.ToString("#,#.##"))));
                    dataRow.Append(docTable.CreateCell(new DocTableCellProp(item.TotalAmount.ToString("#,#.##"))));
                    docTable.Append(dataRow);
                }
                return docTable;
            }
        }
        public DocImage Image1 { get { return new DocImage(this.AppYData.FileImgPath, imageDpi: 300, imageWidth: 300, imageHeight: 60); } }
        public DocHTML HTML1 { get { return new DocHTML("<p style='color: #ff0000'>這是紅色HTML字</p>"); } }
        public DocTableRow TableRow1
        {
            get
            {
                return new DocTableRow(this.AppPDatas.Select(x => new List<string>()
                {
                    x.ProductName,
                    x.Quantity.ToString("#,#.##"),
                    x.Amount.ToString("#,#.##"),
                }).ToList());
            }
        }
        public DocTableCell TableCell1
        {
            get
            {
                var cellData = new DocTableCell(new List<string>() { "A_Cell_1", "A_Cell_2" });
                cellData.IsEmptyRemoveRow = true;
                return cellData;
            }
        }
        public DocTableCell TableCell2
        {
            get
            {
                var cellData = new DocTableCell(new List<string>() { "B_Cell_1" });
                cellData.IsEmptyRemoveRow = true;
                return cellData;
            }
        }
        public DocTableCell TableCell3
        {
            get
            {
                //刪除範例
                var cellData = new DocTableCell();
                cellData.IsEmptyRemoveRow = true;
                return cellData;
            }
        }
    }
}
