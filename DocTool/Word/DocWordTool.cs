using DocTool.Dto;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocTool.Base;
using DocTool.DocType;
using DocTool.Pdf;

namespace DocTool.Word
{

    public class DocWordTool : DocBase
    {
        /// <summary>
        /// 建構子
        /// </summary>
        /// <param name="libreOfficeAppPath">libreOffice路徑</param>
        /// <param name="locationTempPath">暫存資料夾</param>
        public DocWordTool(string libreOfficeAppPath, string locationTempPath) : base(libreOfficeAppPath, locationTempPath)
        {
        }
        /// <summary>
        /// 設置資料
        /// </summary>
        public DocWordTool Set(FileObj fileObjData)
        {
            this.fileData = fileObjData;
            return this;
        }
        /// <summary>
        /// 設置資料
        /// </summary>
        public DocWordTool Set(string inputPath)
        {
            this.fileData = new FileObj()
            {
                fileName = Path.GetFileNameWithoutExtension(inputPath),
                fileType = Path.GetExtension(inputPath).Substring(1),
                fileByteArr = File.ReadAllBytes(inputPath),
            };
            return this;
        }
        /// <summary>
        /// 設置資料
        /// </summary>
        public DocWordTool Set(string fileNameWithExtension, byte[] inputBytes)
        {
            this.fileData = new FileObj()
            {
                fileName = Path.GetFileNameWithoutExtension(fileNameWithExtension),
                fileType = Path.GetExtension(fileNameWithExtension).Substring(1),
                fileByteArr = inputBytes,
            };
            return this;
        }
        /// <summary>
        /// 轉PDF
        /// </summary>
        public DocWordTool ToPDF(bool toPDF = true)
        {
            if (toPDF)
                DocConvert(this.fileData.fileNameWithExtension, this.fileData.fileByteArr, fileExtensionType.pdf);
            return this;
        }
        /// <summary>
        /// Html轉Docx
        /// </summary>
        public DocWordTool HtmlToDocx(string htmlBodyStr)
        {
            var htmlContent = $"<html><head><title>table</title></head><body>{htmlBodyStr}</body></html>";
            var outputFolderPath = Path.Combine(this.locationTempPath, NewOutputFolderPath("output"));
            var outputFilePath = Path.Combine(outputFolderPath, $"output_HTML.html");
            try
            {
                Directory.CreateDirectory(outputFolderPath);
                File.WriteAllText(outputFilePath, htmlContent);
                DocConvert(outputFilePath, fileExtensionType.docx);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                ClearDirectory(outputFolderPath);
            }
            return this;
        }
        /// <summary>
        /// 物件轉字典
        /// </summary>
        Dictionary<string, ReplaceObj> ToPropDictionary(object obj)
        {
            var dictionary = new Dictionary<string, ReplaceObj>();
            var type = obj.GetType();
            foreach (var property in type.GetProperties())
            {
                dictionary.Add(property.Name, new ReplaceObj()
                {
                    PropType = property.PropertyType,
                    Value = property.GetValue(obj),
                });
            }
            return dictionary;
        }
        /// <summary>
        /// 尋找並取代關鍵字-前置處理
        /// </summary>
        private bool PreReplaceTag()
        {
            //是否須轉檔
            var isConversion = replaceTypeConversion.Contains($".{this.fileData.fileType}");
            if (isConversion)
            {
                DocConvert(this.fileData.fileNameWithExtension, this.fileData.fileByteArr, fileExtensionType.docx);
            }
            return isConversion;
        }
        /// <summary>
        /// 尋找並取代關鍵字(物件)(預設Highlight)
        /// </summary>
        public DocWordTool ReplaceTag<T>(T replaceData) where T : class
        {
            return ReplaceTag<T, Highlight>(replaceData);
        }
        /// <summary>
        /// 尋找並取代關鍵字(字典)(預設Highlight)
        /// </summary>
        public DocWordTool ReplaceTag(Dictionary<string, ReplaceObj> replaceDatas)
        {
            return ReplaceTag<Highlight>(replaceDatas);
        }
        /// <summary>
        /// 尋找並取代關鍵字(物件)
        /// T1為尋找樣式，可為Highlight
        /// </summary>
        public DocWordTool ReplaceTag<T, T1>(T replaceData)
            where T : class
            where T1 : OpenXmlElement
        {
            var isConversion = PreReplaceTag();
            //讀取
            using (var newDocMS = new MemoryStream())
            {
                using (var oriDocMS = new MemoryStream(this.fileData.fileByteArr))
                {
                    //複製
                    oriDocMS.CopyTo(newDocMS);
                }
                using (var doc = WordprocessingDocument.Open(newDocMS, true))
                {
                    DocReplace<T1>(doc, ToPropDictionary(replaceData));
                    doc.Save();
                    this.fileData.fileByteArr = newDocMS.ToArray();
                    if (isConversion)
                    {
                        //轉回odt
                        DocConvert(this.fileData.fileNameWithExtension, this.fileData.fileByteArr, fileExtensionType.odt);
                    }
                }
            }
            return this;
        }
        /// <summary>
        /// 尋找並取代關鍵字(字典)
        /// T為尋找樣式，可為Highlight
        /// </summary>
        public DocWordTool ReplaceTag<T>(Dictionary<string, ReplaceObj> replaceDatas) where T : OpenXmlElement
        {
            var isConversion = PreReplaceTag();
            //讀取
            using (var newDocMS = new MemoryStream())
            {
                using (var oriDocMS = new MemoryStream(this.fileData.fileByteArr))
                {
                    //複製
                    oriDocMS.CopyTo(newDocMS);
                }
                using (var doc = WordprocessingDocument.Open(newDocMS, true))
                {
                    DocReplace<T>(doc, replaceDatas);
                    doc.Save();
                    this.fileData.fileByteArr = newDocMS.ToArray();
                    if (isConversion)
                    {
                        //轉回odt
                        DocConvert(this.fileData.fileNameWithExtension, this.fileData.fileByteArr, fileExtensionType.odt);
                    }
                }
            }
            return this;
        }
        /// <summary>
        /// 取代的主要功能方法
        /// </summary>
        private void DocReplace<T>(WordprocessingDocument doc, Dictionary<string, ReplaceObj> replaceDatas) where T : OpenXmlElement
        {
            var maxCount = 5;
            var indexCount = 0;
            while (true)
            {
                var isMapSuccess = StyleReplace<T>(doc, replaceDatas, GetAllStyleRuns<T>(doc));
                if (!isMapSuccess)
                {
                    break;
                }
                if (indexCount > maxCount)
                {
                    break;
                }
                indexCount++;
            }
        }
        private IEnumerable<Run> GetAllStyleRuns<T>(WordprocessingDocument doc) where T : OpenXmlElement
        {
            var styleRuns = doc.MainDocumentPart.Document
                .Descendants<Run>()
                .Concat(doc.MainDocumentPart.HeaderParts.SelectMany(h => h.Header.Descendants<Run>()))
                .Concat(doc.MainDocumentPart.FooterParts.SelectMany(f => f.Footer.Descendants<Run>()))
                .Where(x => x.RunProperties?.Elements<T>().Any() ?? false);
            return styleRuns;
        }
        /// <summary>
        /// 尋找指定樣式並取代
        /// </summary>
        private bool StyleReplace<T>(WordprocessingDocument doc, Dictionary<string, ReplaceObj> replaceDatas, IEnumerable<Run> docStyleRuns) where T : OpenXmlElement
        {
            var tempPool = new List<Run>();
            var matchText = string.Empty;
            var isMapSuccess = false;
            foreach (var itemRun in docStyleRuns)
            {
                var text = itemRun.InnerText;
                if (text.StartsWith(this.prefixStr))
                {
                    tempPool = new List<Run>() { itemRun };
                    matchText = text;
                }
                else
                {
                    matchText = matchText + text;
                    tempPool.Add(itemRun);
                }
                if (text.EndsWith(this.suffixStr))
                {
                    var m = Regex.Match(matchText, $@"{String.Join("", this.prefixStr.Select(x => "\\" + x))}(?<n>\w+){String.Join("", this.suffixStr.Select(x => "\\" + x))}");
                    var key = m.Groups["n"].Value;
                    if (m.Success && replaceDatas.ContainsKey(key))
                    {
                        isMapSuccess = true;
                        var firstRun = tempPool.First();
                        firstRun.RemoveAllChildren<Text>();
                        if (firstRun.RunProperties == null)
                            continue;
                        firstRun.RunProperties.RemoveAllChildren<T>();
                        var keyProp = replaceDatas[key];
                        //用型態區分
                        if (keyProp.PropType == typeof(DocTable))
                        {
                            AppendTableToElement((DocTable)keyProp.Value, firstRun, doc);
                        }
                        else if (keyProp.PropType == typeof(DocImage))
                        {
                            AppendImageToElement((DocImage)keyProp.Value, firstRun, doc);
                        }
                        else if (keyProp.PropType == typeof(DocHTML))
                        {
                            AppendHTMLToElement((DocHTML)keyProp.Value, firstRun);
                        }
                        else if (keyProp.PropType == typeof(DocTableRow))
                        {
                            AppendTableRowToElement((DocTableRow)keyProp.Value, firstRun);
                        }
                        else if (keyProp.PropType == typeof(DocTableCell))
                        {
                            AppendTableCellToElement((DocTableCell)keyProp.Value, firstRun);
                        }
                        else
                        {
                            var firstLine = true;
                            foreach (var line in Regex.Split((keyProp.Value ?? "").ToString(), @"\r?\n"))
                            {
                                if (firstLine) firstLine = false;
                                else firstRun.Append(new Break());
                                firstRun.Append(new Text(line));
                            }
                        }
                        tempPool.Skip(1).ToList().ForEach(o => o.Remove());
                    }
                }
            }
            return isMapSuccess;
        }
        /// <summary>
        /// 新增表格到指定位置
        /// </summary>
        private void AppendTableToElement(DocTable tableData, OpenXmlElement element, WordprocessingDocument doc)
        {
            if (tableData == null)
                return;
            if (tableData.DocImageDict.Any())
            {
                foreach (var cell in tableData.Descendants<TableCell>())
                {
                    if (cell.Descendants<Highlight>().Any())
                    {
                        foreach (var run in cell.Descendants<Run>())
                        {
                            foreach (var text in run.Descendants<Text>())
                            {
                                //動態表格內圖片渲染
                                if (tableData.DocImageDict.ContainsKey(text.Text))
                                {
                                    if (!cell.Descendants<Table>().Any())
                                    {
                                        run.RunProperties.RemoveAllChildren<Highlight>();
                                        run.RemoveAllChildren<Text>();
                                        AppendImageToElement(tableData.DocImageDict[text.Text], run, doc);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            element.Parent.Append(tableData.CloneNode(true));
        }
        /// <summary>
        /// 新增圖片到指定位置
        /// </summary>
        private void AppendImageToElement(DocImage replaceData, OpenXmlElement element, WordprocessingDocument doc)
        {
            if (replaceData == null)
                return;
            var relationshipId = replaceData.FeedImgData(doc);
            var imgElement = replaceData.GetImageElement(relationshipId);
            element.Append(imgElement);
        }
        /// <summary>
        /// 新增HTML到指定位置
        /// </summary>
        private void AppendHTMLToElement(DocHTML htmlData, OpenXmlElement element)
        {
            if (htmlData == null)
                return;
            //html轉docx
            var DocxData = new DocWordTool(this.libreOfficeAppPath, this.locationTempPath)
                .HtmlToDocx(htmlData.HTMLStr)
                .GetData();
            if (DocxData == null)
                return;
            //讀取
            using (var newDocMS = new MemoryStream())
            {
                using (var oriDocMS = new MemoryStream(DocxData.fileByteArr))
                {
                    //複製
                    oriDocMS.CopyTo(newDocMS);
                }
                using (var htmlDoc = WordprocessingDocument.Open(newDocMS, true))
                {
                    var bodyElement = htmlDoc.MainDocumentPart.Document.Descendants<Body>().FirstOrDefault();
                    if (bodyElement != null && bodyElement.FirstChild != null)
                    {
                        var newContent = bodyElement.FirstChild
                            .Where(x => x is Run);
                        foreach (var runItem in newContent)
                        {
                            element.Parent.AppendChild(runItem.CloneNode(true));
                        }
                    }
                }
            }
        }
        /// <summary>
        /// 新增多個表格列到指定表格位置(一整欄)
        /// </summary>
        private void AppendTableRowToElement(DocTableRow tableRowData, OpenXmlElement element)
        {
            if (tableRowData == null)
                return;
            var tempElement = element;
            TableRow currentTableRow = null;
            while (tempElement != null && !(tempElement is DocTable) && !(tempElement is Table))
            {
                tempElement = tempElement.Parent;
                if (tempElement is TableRow)
                {
                    if (currentTableRow == null)
                        currentTableRow = tempElement as TableRow;
                }
            }
            var currentTable = tempElement as Table;
            if (currentTable != null)
            {
                try
                {
                    tableRowData.RowDatas.Reverse();
                    foreach (var item in tableRowData.RowDatas)
                        currentTable.InsertAfter(item.CloneNode(true), currentTableRow);
                    currentTable.RemoveChild(currentTableRow);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                finally
                {
                    tableRowData.RowDatas.Reverse();
                }
            }
        }
        /// <summary>
        /// 新增多個表格儲存格到指定表格位置
        /// </summary>
        private void AppendTableCellToElement(DocTableCell tableCellCollectionData, OpenXmlElement element)
        {
            var targetCell = FindParentElement<TableCell>(element);
            var targetRow = FindParentElement<TableRow>(element);
            if (targetCell == null || targetRow == null)
                return;
            int targetCellIndex = 0;
            int oriCellIndex = 0;
            var allCellCnt = targetRow.Elements<TableCell>().Count();
            if (tableCellCollectionData.IsEmptyRemoveRow
                && (tableCellCollectionData == null || tableCellCollectionData.CellDatas == null || tableCellCollectionData.CellDatas.Count == 0))
            {
                //刪除整欄
                var targetTable = FindParentElement<Table>(element);
                if (targetTable != null)
                {
                    targetTable.RemoveChild(targetRow);
                }
            }
            else
            {
                //找cell位置
                foreach (var cell in targetRow.Elements<TableCell>())
                {
                    if (cell == targetCell)
                    {
                        break;
                    }
                    targetCellIndex++;
                }
                //取代
                for (int i = 0; i < allCellCnt; i++)
                {
                    if (i < targetCellIndex)
                        continue;
                    if (oriCellIndex + 1 > tableCellCollectionData.CellDatas.Count)
                        break;
                    var currentCell = targetRow.Elements<TableCell>().ElementAt(i);
                    currentCell.RemoveAllChildren();
                    foreach (var item in tableCellCollectionData.CellDatas[oriCellIndex].Elements<Paragraph>())
                    {
                        currentCell.Append(item.CloneNode(true));
                    }
                    oriCellIndex++;
                }
            }
        }

        public T FindParentElement<T>(OpenXmlElement element, int maxDeep = 20) where T : OpenXmlElement
        {
            var i = 0;
            while (element != null)
            {
                if (i > maxDeep)
                    break;
                if (element is T t)
                {
                    return t;
                }
                element = element.Parent;
                i++;
            }
            return null;
        }
    }
}
