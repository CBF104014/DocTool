# DocTool
使用LibreOffice，將word文件轉pdf以及製作word文件，請先於主機上下載LibreOffice
  
### 1.初始化建構
>libreOfficeAppPath：LibreOffice路徑位置(\LibreOfficePortable\App\libreoffice\program\soffice.exe)  
>locationTempPath：指定文件暫時存放位置  
```c#
private DocWordTool wordTool { get; set; }
public SampleCode1(string libreOfficeAppPath, string locationTempPath)
{
    this.wordTool = new DocWordTool(libreOfficeAppPath, locationTempPath);
}
```
  .
### 2.Word文件替換關鍵字範例
>ReplaceType型態  
>0：Text(文字)  
>1：Table(表格)  
>2：Image(圖片)  
>3：HtmlString(HTML)  
>4：TableRow(表格-列)  
```c#
var replaceData = new Dictionary<string, ReplaceDto>()
{
     ["PONo"] = new ReplaceDto() { textStr = "11300000A1" },
     ["BarCode1_Image"] = new ReplaceDto()
     {
         replaceType = ReplaceType.Image,
         fileName = "barcode",
         fileType = "jpg",
         fileByteArr = new byte[0],
         imageDpi = 72,
         imageHeight = 30,
         imageWidth = 150,
     },
     ["PoVenTable"] = new ReplaceDto()
     {
         replaceType = ReplaceType.Table,
         tableData = this.wordTool.ToXMLTable(new List<string>() { "品名", "數量", "總價" }, new List<List<string>>() { new List<string>() { "A", "1", "1500" } }),
     },
     ["AppP_Table"] = new ReplaceDto()
     {
         replaceType = ReplaceType.HtmlString,
         htmlStr = "<p style=\"background-color:Tomato;\">Lorem ipsum...</p>",
     },
     ["AppP2_TableRow"] = new ReplaceDto()
     {
         replaceType = ReplaceType.TableRow,
         tableRowDatas = new List<List<string>>() { new List<string>() { "A", "1", "1500" } }
             .Select(x =>
             {
                  var row = this.wordTool.CreateRow();
                  x.ForEach(item => row.Append(this.wordTool.CreateCell(item)));
                  return row;
             }).ToList()
     }
};
```
>fileData：附件  
>replaceData：取代資料  
>確保文件內的標籤是{$Tag$}並加上highlight樣式，如右圖![tag](https://img.shields.io/badge/-{$YourTag$}-fffd00?style=for-the-badge)  
```c#
public FileObj WordReplaceTag(FileObj fileData, Dictionary<string, ReplaceDto> replaceData)
{
     fileData = wordTool
        .ReplaceTag($"{fileData.fileName}.{fileData.fileType}", fileData.fileByteArr, replaceData)
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
    fileData = wordTool
        .ToPDF($"{fileData.fileName}.{fileData.fileType}", fileData.fileByteArr)
        .GetData();
    return fileData;
}
```
