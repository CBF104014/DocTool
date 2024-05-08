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

namespace DocTool.Word
{

    public class DocWordTool : DocBase
    {
        public DocWordTool(string libreOfficeAppPath, string locationTempPath) : base(libreOfficeAppPath, locationTempPath)
        {
        }

        public DocWordTool ToPDF()
        {
            DocConvert($"{this.fileData.fileName}.{this.fileData.fileType}", this.fileData.fileByteArr, fileExtensionType.pdf);
            return this;
        }
        public DocWordTool ToPDF(string inputPath)
        {
            DocConvert(inputPath, fileExtensionType.pdf);
            return this;
        }
        public DocWordTool ToPDF(string fileNameWithExtension, byte[] inputBytes)
        {
            DocConvert(fileNameWithExtension, inputBytes, fileExtensionType.pdf);
            return this;
        }
        public DocWordTool ToPDF(FileObj fileObjData)
        {
            DocConvert($"{fileObjData.fileName}.{fileObjData.fileType}", fileObjData.fileByteArr, fileExtensionType.pdf);
            return this;
        }
        public DocWordTool HtmlToDocx(string htmlBosyStr)
        {
            var htmlContent = $"<html><head><title>table</title></head><body>{htmlBosyStr}</body></html>";
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
        public class ReplaceObj
        {
            public Type PropType { get; set; }
            public object Value { get; set; }
        }
        /// <summary>
        /// 創立表格
        /// </summary>
        /// <param name="tableWidthCM">欄位公分加總</param>
        public DocTable CreateTable(double tableWidthCM = 0)
        {
            var table = new DocTable();
            var tableProperties = new TableProperties(
                new TableBorders(
                    new TopBorder() { Val = BorderValues.Single, Color = "000000", Size = 12 },
                    new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = 12 },
                    new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = 12 },
                    new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = 12 },
                    new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = 12 },
                    new InsideVerticalBorder() { Val = BorderValues.Single, Color = "000000", Size = 12 }
                )
            //new TableLayout { Type = TableLayoutValues.Autofit },
            );
            //表格指定寬度
            if (tableWidthCM > 0)
            {
                tableProperties.Append(new TableWidth() { Width = DocTable.CMToDXA(tableWidthCM), Type = TableWidthUnitValues.Dxa });
            }
            table.AppendChild(tableProperties);
            return table;
        }
        Dictionary<string, ReplaceObj> ToPropDictionary(object obj)
        {
            var dictionary = new Dictionary<string, ReplaceObj>();
            var type = obj.GetType();
            foreach (var property in type.GetProperties())
            {
                string propertyName = property.Name;
                //object propertyValue = property.GetValue(obj);
                dictionary.Add(propertyName, new ReplaceObj()
                {
                    PropType = property.PropertyType,
                    Value = property.GetValue(obj),
                });
            }

            return dictionary;
        }
        public DocWordTool ReplaceTag<T>(string fileNameWithExtension, byte[] inputBytes, T replaceData)
        {
            //讀取
            using (var newDocMS = new MemoryStream())
            {
                using (var oriDocMS = new MemoryStream(inputBytes))
                {
                    //複製
                    oriDocMS.CopyTo(newDocMS);
                }
                using (var doc = WordprocessingDocument.Open(newDocMS, true))
                {
                    DocReplace(doc, ToPropDictionary(replaceData));
                    doc.Save();
                    this.fileData = new FileObj()
                    {
                        fileName = Path.GetFileNameWithoutExtension(fileNameWithExtension),
                        fileType = Path.GetExtension(fileNameWithExtension).Substring(1),
                        fileByteArr = newDocMS.ToArray(),
                    };
                }
            }
            return this;
        }
        public DocWordTool ReplaceTag(string fileNameWithExtension, byte[] inputBytes, Dictionary<string, ReplaceObj> replaceDatas)
        {
            //讀取
            using (var newDocMS = new MemoryStream())
            {
                using (var oriDocMS = new MemoryStream(inputBytes))
                {
                    //複製
                    oriDocMS.CopyTo(newDocMS);
                }
                using (var doc = WordprocessingDocument.Open(newDocMS, true))
                {
                    DocReplace(doc, replaceDatas);
                    doc.Save();
                    this.fileData = new FileObj()
                    {
                        fileName = Path.GetFileNameWithoutExtension(fileNameWithExtension),
                        fileType = Path.GetExtension(fileNameWithExtension).Substring(1),
                        fileByteArr = newDocMS.ToArray(),
                    };
                }
            }
            return this;
        }
        private void DocReplace(WordprocessingDocument doc, Dictionary<string, ReplaceObj> replaceDatas)
        {
            //渲染
            HighlightReplace(doc, replaceDatas, doc.MainDocumentPart.Document.Descendants<Run>()
                    .Concat(doc.MainDocumentPart.HeaderParts.SelectMany(h => h.Header.Descendants<Run>()))
                    .Concat(doc.MainDocumentPart.FooterParts.SelectMany(f => f.Footer.Descendants<Run>()))
                    .Where(x => x.RunProperties?.Elements<Highlight>().Any() ?? false));
        }
        private void HighlightReplace(WordprocessingDocument doc, Dictionary<string, ReplaceObj> replaceDatas, IEnumerable<Run> docHighlightRuns)
        {
            var tempPool = new List<Run>();
            var matchText = string.Empty;
            foreach (var highlightRun in docHighlightRuns)
            {
                var text = highlightRun.InnerText;
                if (text.StartsWith(this.prefixStr))
                {
                    tempPool = new List<Run>() { highlightRun };
                    matchText = text;
                }
                else
                {
                    matchText = matchText + text;
                    tempPool.Add(highlightRun);
                }
                if (text.EndsWith(this.suffixStr))
                {
                    var m = Regex.Match(matchText, $@"{String.Join("", this.prefixStr.Select(x => "\\" + x))}(?<n>\w+){String.Join("", this.suffixStr.Select(x => "\\" + x))}");
                    var key = m.Groups["n"].Value;
                    if (m.Success && replaceDatas.ContainsKey(key))
                    {
                        var firstRun = tempPool.First();
                        firstRun.RemoveAllChildren<Text>();
                        if (firstRun.RunProperties == null)
                            continue;
                        firstRun.RunProperties.RemoveAllChildren<Highlight>();
                        var keyProp = replaceDatas[key];
                        var isRemove = true;
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
                            isRemove = false;
                            AppendTableRowToElement((DocTableRow)keyProp.Value, firstRun);
                        }
                        else
                        {
                            var firstLine = true;
                            foreach (var line in Regex.Split(keyProp.Value.ToString(), @"\\n"))
                            {
                                if (firstLine) firstLine = false;
                                else firstRun.Append(new Break());
                                firstRun.Append(new Text(line));
                            }
                        }
                        if (isRemove)
                            tempPool.Skip(1).ToList().ForEach(o => o.Remove());
                    }
                }
            }

        }
        //private TableCell CreateTableCell(string text, int gridSpan)
        //{
        //    TableCell cell = new TableCell(new Paragraph(new Run(new Text(text))));
        //    TableCellProperties cellProp = new TableCellProperties(new GridSpan { Val = gridSpan });
        //    cell.AppendChild(cellProp);
        //    return cell;
        //}
        private void AppendTableToElement(DocTable tableData, OpenXmlElement element, WordprocessingDocument doc)
        {
            if (tableData != null)
            {
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
        }
        private void AppendImageToElement(DocImage replaceData, OpenXmlElement element, WordprocessingDocument doc)
        {
            if (replaceData != null)
            {
                var relationshipId = replaceData.FeedImgData(doc);
                var imgElement = replaceData.GetImageElement(relationshipId);
                element.Append(imgElement);
            }
        }
        private void AppendHTMLToElement(DocHTML htmlData, OpenXmlElement element)
        {
            if (htmlData != null)
            {
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
        }
        private void AppendTableRowToElement(DocTableRow tableRowDatas, OpenXmlElement element)
        {
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
                foreach (var item in tableRowDatas.RowDatas)
                    currentTable.InsertAfter(item, currentTableRow);
                currentTable.RemoveChild(currentTableRow);
            }
        }
    }
}
