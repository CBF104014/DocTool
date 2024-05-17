# DocTool
使用LibreOffice，將word文件轉pdf以及製作word文件，請先於主機上下載LibreOffice。
範例專案為SampleApp(請設為起始專案)，範例附件為SampleApp專案項下SampleFile位置，可自行複製到C:\TEMP資料夾。
>LibreOffice下載網址：https://zh-tw.libreoffice.org/download/portable-versions/
  
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
```c#
    /// <summary>
    /// 範例-細項
    /// </summary>
    public class MyPitem
    {
        public string ProductName { get; set; }
        public decimal Quantity { get; set; }
        public decimal Amount { get; set; }
        public decimal TotalAmount { get => this.Quantity * this.Amount; }
    }
    /// <summary>
    /// 範例-原始資料
    /// </summary>
    public class MyOriData
    {
        public MyOriData()
        {
            this.PONo = $"{DateTime.Now.ToString("yyyyMMdd")}-Test";
            this.PageEndText1 = $"這是頁尾";
            this.PDatas = new List<MyPitem>
            {
                new MyPitem() { ProductName = "品項A", Quantity = 3, Amount = 1000 },
                new MyPitem() { ProductName = "品項B", Quantity = 77, Amount = 299 },
                new MyPitem() { ProductName = "品項C", Quantity = 30, Amount = 130 },
                new MyPitem() { ProductName = "品項D", Quantity = 234, Amount = 1350 },
                new MyPitem() { ProductName = "品項E", Quantity = 10, Amount = 990 },
            };
        }
        public string PONo { get; set; }
        public string PageEndText1 { get; set; }
        public string outFilePath { get => "C:\\TEMP"; }
        public string FileDocName { get => "Word_hightlight測試.docx"; }
        public string FileDocPath { get => Path.Combine(outFilePath, this.FileDocName); }
        public string FileImgName { get => "barcode.jpg"; }
        public string FileImgPath { get => Path.Combine(outFilePath, this.FileImgName); }
        public List<MyPitem> PDatas { get; set; }
        public decimal Total { get => this.PDatas == null ? 0 : this.PDatas.Sum(x => x.TotalAmount); }
        public string Total_str { get => this.Total.ToString("#,#.##"); }
    }
    /// <summary>
    /// 範例-處理後資料
    /// </summary>
    public class MyClass : MyOriData
    {
        public DocTable Table1
        {
            get
            {
                var docTable = DocWordTool.CreateTable();
                var dataRow = docTable.CreateRow();
                var imgData = new DocImage(this.FileImgPath, imageDpi: 300, imageWidth: 600, imageHeight: 100);
                dataRow.Append(docTable.CreateCell(new DocTableCellPrpo("欄位A")));
                dataRow.Append(docTable.CreateCell(new DocTableCellPrpo("欄位B", JustificationValues.Center, TableVerticalAlignmentValues.Center)));
                dataRow.Append(docTable.CreateCell(new DocTableCellPrpo(imgData, JustificationValues.Center, TableVerticalAlignmentValues.Center, colSpan: 2)));
                dataRow.Append(docTable.CreateCell(new DocTableCellPrpo("", colSpan: 0)));
                docTable.Append(dataRow);
                foreach (var item in base.PDatas)
                {
                    dataRow = docTable.CreateRow();
                    dataRow.Append(docTable.CreateCell(new DocTableCellPrpo(item.ProductName)));
                    dataRow.Append(docTable.CreateCell(new DocTableCellPrpo(item.Quantity.ToString("#,#.##"))));
                    dataRow.Append(docTable.CreateCell(new DocTableCellPrpo(item.Amount.ToString("#,#.##"))));
                    dataRow.Append(docTable.CreateCell(new DocTableCellPrpo(item.TotalAmount.ToString("#,#.##"))));
                    docTable.Append(dataRow);
                }
                return docTable;
            }
        }
        public DocImage Image1 { get { return new DocImage(base.FileImgPath, imageDpi: 300, imageWidth: 300, imageHeight: 60); } }
        public DocImage Image2 { get { return new DocImage(base.FileImgPath, imageDpi: 300, imageWidth: 300, imageHeight: 60); } }
        public DocHTML HTML1 { get { return new DocHTML("<p style='color: #ff0000'>這是紅色HTML字</p>"); } }
        public DocTableRow TableRow1
        {
            get
            {
                return new DocTableRow(base.PDatas.Select(x => new List<string>()
                {
                    x.ProductName,
                    x.Quantity.ToString("#,#.##"),
                    x.Amount.ToString("#,#.##"),
                }).ToList());
            }
        }
    }
```
>確保文件內的標籤是{$Tag$}並加上highlight樣式，如右圖![tag](https://img.shields.io/badge/-{$YourTag$}-fffd00?style=for-the-badge)  
```c#
public FileObj WordReplaceTag()
{
     var docData = new MyClass();
     //輸出WORD
     var fileData = docTool.Word
        .ReplaceTag(docData.FileDocName, System.IO.File.ReadAllBytes(docData.FileDocPath), docData)
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
     var docData = new MyClass();
     //輸出PDF
    fileData = docTool.Word
        .ToPDF(fileData)
        .GetData();
    return fileData;
}
```
