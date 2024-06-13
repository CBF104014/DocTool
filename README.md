# DocTool
使用OpenXml SDK以及LibreOffice，將word文件轉pdf以及製作word文件，請先於主機上下載LibreOffice。
範例專案為SampleApp(請設為起始專案)，範例附件為SampleApp專案項下SampleFile位置，可自行複製到C:\TEMP資料夾。
>依賴LibreOffice(下載網址：https://zh-tw.libreoffice.org/download/portable-versions/)  
>依賴DocumentFormat.OpenXml 3.0.2  
>依賴iTextSharp 5.5.13.3  

  
### 1.初始化建構
>libreOfficeAppPath：LibreOffice路徑位置(\LibreOfficePortable\App\libreoffice\program\soffice.exe)  
>locationTempPath：指定文件暫時存放位置  
```c#
private Tool docTool { get; set; }
public SampleCode1(string libreOfficeAppPath, string locationTempPath)
{
    this.docTool = new Tool(libreOfficeAppPath, locationTempPath);
}
```
  .
### 2.Word文件替換關鍵字範例
>轉換型態  
>String(文字)  
>DocTable(表格)  
>DocImage(圖片)  
>DocHtml(HTML)  
>DocTableRow(表格-列)  
>DocTableCell(表格-儲存格)  
```c#
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
```
>確保文件內的標籤是{$Tag$}(可指定尋找樣式，預設是Highlight)，如右圖![tag](https://img.shields.io/badge/-{$YourTag$}-fffd00?style=for-the-badge)  
```c#
public FileObj WordReplaceTag()
{
     var docData = new MyDocClass();
     //輸出WORD
     var fileData = docTool.Word
        .Set(docData.FileDocPath)
        .ReplaceTag(docData)
        .GetData();
    return fileData;
}
```
  .
### 3.轉PDF範例
>fileData：附件
**目前只接受.doc、.docx、.odt、.xls、.xlsx、.htm、.html
```c#
public FileObj WordToPdf(FileObj fileData)
{
     var docData = new MyDocClass();
     //輸出PDF
    fileData = docTool.Word
        .Set(docData.FileDocPath)
        .ToPDF()
        .GetData();
    return fileData;
}
```
