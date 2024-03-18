using DocTool.Dto;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Experimental;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace DocTool
{
    public class DocWordTool : DocBase
    {
        private int _imageCounter { get; set; } = 0;
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
        public Table ToXMLTable(List<string> tableHead, List<List<string>> tableRows)
        {
            var table = CreateTable();
            var row = CreateRow();
            foreach (var item in tableHead)
            {
                var cell = CreateCell(item, isBold: true);
                row.Append(cell);
            }
            table.Append(row);
            foreach (var rowItem in tableRows)
            {
                row = CreateRow();
                foreach (var item in rowItem)
                {
                    var cell = CreateCell(item, isBold: true);
                    row.Append(cell);
                }
                table.Append(row);
            }
            return table;
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
        public DocWordTool ReplaceTag(string fileNameWithExtension, byte[] inputBytes, Dictionary<string, ReplaceDto> replaceDatas)
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

        private void DocReplace(WordprocessingDocument doc, Dictionary<string, ReplaceDto> replaceDatas)
        {
            //第一次渲染
            HighlightReplace(doc, replaceDatas,
                doc.MainDocumentPart.Document.Descendants<Run>()
                .Concat(doc.MainDocumentPart.HeaderParts.SelectMany(h => h.Header.Descendants<Run>()))
                .Concat(doc.MainDocumentPart.FooterParts.SelectMany(f => f.Footer.Descendants<Run>()))
                .Where(x => x.RunProperties?.Elements<Highlight>().Any() ?? false));
            //第二次渲染(表格內的)
            HighlightReplace(doc, replaceDatas,
                doc.MainDocumentPart.Document.Descendants<Run>()
                .Concat(doc.MainDocumentPart.HeaderParts.SelectMany(h => h.Header.Descendants<Run>()))
                .Concat(doc.MainDocumentPart.FooterParts.SelectMany(f => f.Footer.Descendants<Run>()))
                .Where(x => x.RunProperties?.Elements<Highlight>().Any() ?? false));
        }
        private void HighlightReplace(WordprocessingDocument doc, Dictionary<string, ReplaceDto> replaceDatas, IEnumerable<Run> docHighlightRuns)
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
                        var keyObj = replaceDatas[key];
                        //用型態區分
                        switch (keyObj.replaceType)
                        {
                            case ReplaceType.Table:
                                {
                                    AppendTableToElement(keyObj.tableData, firstRun);
                                    break;
                                }
                            case ReplaceType.Image:
                                {
                                    _imageCounter++;
                                    AppendImageToElement(keyObj, firstRun, doc);
                                    break;
                                }
                            case ReplaceType.HtmlString:
                                {
                                    AppendHTMLToElement(keyObj.htmlStr, firstRun);
                                    break;
                                }
                            case ReplaceType.TableRow:
                                {
                                    AppendTableRowToElement(keyObj.tableRowDatas, firstRun);
                                    break;
                                }
                            default:
                                {
                                    var firstLine = true;
                                    foreach (var line in Regex.Split(keyObj.textStr, @"\\n"))
                                    {
                                        if (firstLine) firstLine = false;
                                        else firstRun.Append(new Break());
                                        firstRun.Append(new Text(line));
                                    }
                                    break;
                                }
                        }
                        tempPool.Skip(1).ToList().ForEach(o => o.Remove());
                    }
                }
            }

        }
        private void AppendTableRowToElement(List<TableRow> tableRowDatas, OpenXmlElement element)
        {
            var tempElement = element;
            TableRow currentTableRow = null;
            while (tempElement != null && !(tempElement is Table))
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
                foreach (var item in tableRowDatas)
                    currentTable.InsertAfter(item, currentTableRow);
                currentTable.RemoveChild(currentTableRow);
            }
        }
        private void AppendHTMLToElement(string htmlBodyStr, OpenXmlElement element)
        {
            //html轉docx
            HtmlToDocx(htmlBodyStr);
            if (GetData() == null)
                return;
            //讀取
            using (var newDocMS = new MemoryStream())
            {
                using (var oriDocMS = new MemoryStream(GetData().fileByteArr))
                {
                    //複製
                    oriDocMS.CopyTo(newDocMS);
                }
                using (var doc = WordprocessingDocument.Open(newDocMS, true))
                {
                    var bodyElement = doc.MainDocumentPart.Document.Descendants<Body>().FirstOrDefault();
                    if (bodyElement != null && bodyElement.FirstChild != null)
                        element.Append(bodyElement.FirstChild.CloneNode(true));
                }
            }
        }

        private void AppendTableToElement(Table table, OpenXmlElement element)
        {
            element.Append(table.CloneNode(true));
        }

        private void AppendImageToElement(ReplaceDto replaceData, OpenXmlElement element, WordprocessingDocument doc)
        {
            if (replaceData.imageDpi <= 0)
                replaceData.imageDpi = 72;
            var imageUri = new Uri($"/word/media/{replaceData.fileName}{_imageCounter}.{replaceData.fileType}", UriKind.Relative);
            var packageImagePart = doc.GetPackage().CreatePart(imageUri, $"Image/{replaceData.fileType}", CompressionOption.Normal);
            packageImagePart.GetStream(FileMode.Open, FileAccess.Write).Write(replaceData.fileByteArr, 0, replaceData.fileByteArr.Length);
            var documentPackagePart = doc.GetPackage().GetPart(new Uri("/word/document.xml", UriKind.Relative));
            var imageRelationshipPart = documentPackagePart.Relationships.Create(
                new Uri($"media/{replaceData.fileName}{_imageCounter}.{replaceData.fileType}", UriKind.Relative),
                TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
            var drawing = GetImageElement(imageRelationshipPart.Id, replaceData.fileName, "picture", replaceData.imageWidth, replaceData.imageHeight, replaceData.imageDpi);
            element.AppendChild(drawing);
        }

        private Drawing GetImageElement(
            string imagePartId,
            string fileName,
            string pictureName,
            double width,
            double height,
            double ppi)
        {
            double englishMetricUnitsPerInch = 914400;
            double pixelsPerInch = ppi;
            //calculate size in emu
            double emuWidth = width * englishMetricUnitsPerInch / pixelsPerInch;
            double emuHeight = height * englishMetricUnitsPerInch / pixelsPerInch;

            var element = new Drawing(
                new DW.Inline(
                    new DW.Extent { Cx = (Int64Value)emuWidth, Cy = (Int64Value)emuHeight },
                    new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                    new DW.DocProperties { Id = (UInt32Value)1U, Name = pictureName + _imageCounter },
                    new DW.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks { NoChangeAspect = true }),
                    new A.Graphic(
                        new A.GraphicData(
                            new PIC.Picture(
                                new PIC.NonVisualPictureProperties(
                                    new PIC.NonVisualDrawingProperties { Id = (UInt32Value)0U, Name = fileName },
                                    new PIC.NonVisualPictureDrawingProperties()),
                                new PIC.BlipFill(
                                    new A.Blip(
                                        new A.BlipExtensionList(
                                            new A.BlipExtension { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" }))
                                    {
                                        Embed = imagePartId,
                                        CompressionState = A.BlipCompressionValues.Print,
                                    },
                                            new A.Stretch(new A.FillRectangle())),
                                new PIC.ShapeProperties(
                                    new A.Transform2D(
                                        new A.Offset { X = 0L, Y = 0L },
                                        new A.Extents { Cx = (Int64Value)emuWidth, Cy = (Int64Value)emuHeight }),
                                    new A.PresetGeometry(new A.AdjustValueList())
                                    {
                                        Preset = A.ShapeTypeValues.Rectangle
                                    })))
                        {
                            Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
                        }))
                {
                    DistanceFromTop = (UInt32Value)0U,
                    DistanceFromBottom = (UInt32Value)0U,
                    DistanceFromLeft = (UInt32Value)0U,
                    DistanceFromRight = (UInt32Value)0U,
                    EditId = "50D07946"
                });
            return element;
        }

        public Table CreateTable(decimal tableWidth = 100)
        {
            var table = new Table();
            //表格寬度
            var width = new TableWidth() { Width = $"{tableWidth}%", Type = TableWidthUnitValues.Pct };
            //表格框線
            var borders = new TableBorders(
                new TopBorder() { Val = BorderValues.Single, Color = "000000" },
                new BottomBorder() { Val = BorderValues.Single, Color = "000000" },
                new LeftBorder() { Val = BorderValues.Single, Color = "000000" },
                new RightBorder() { Val = BorderValues.Single, Color = "000000" }
            );
            var tableProperties = new TableProperties(borders, width);
            table.AppendChild(tableProperties);
            return table;
        }
        public TableRow CreateRow()
        {
            var row = new TableRow();
            return row;
        }
        public TableCell CreateCell(string cellText, decimal fontSize = 24, bool isBold = false, int rowspan = 1, int colspan = 1)
        {
            #region cell
            var cell = new TableCell();
            if (rowspan > 1)
            {
                cell.AppendChild(new TableCellProperties(new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }));
            }

            if (colspan > 1)
            {
                cell.AppendChild(new TableCellProperties(new GridSpan() { Val = colspan }));
            }
            cell.AppendChild(new TableCellProperties(new TableCellBorders(
                new TopBorder() { Val = BorderValues.Single, Color = "000000" },
                new BottomBorder() { Val = BorderValues.Single, Color = "000000" },
                new LeftBorder() { Val = BorderValues.Single, Color = "000000" },
                new RightBorder() { Val = BorderValues.Single, Color = "000000" }
            )));
            #endregion
            #region Paragraph
            var paragraph = new Paragraph();
            #endregion
            #region Run
            var run = new Run();
            var runProperties = new RunProperties();
            runProperties.Append(new RunFonts() { Ascii = "標楷體", HighAnsi = "標楷體", EastAsia = "標楷體" });
            runProperties.Append(new FontSize() { Val = fontSize.ToString() });
            if (isBold)
                runProperties.Append(new Bold());
            run.Append(runProperties);
            #endregion
            #region Text
            var text = new Text(cellText);
            #endregion
            //組合
            run.Append(text);
            paragraph.Append(run);
            cell.Append(paragraph);
            //TODO:合併問題
            // rowspan
            //if (rowspan > 1)
            //{
            //    var vMerge = new VerticalMerge() { Val = MergedCellValues.Restart };
            //    cell.AppendChild(new TableCellProperties(vMerge));
            //    for (int i = 1; i < rowspan; i++)
            //    {
            //        var mergedCell = new TableCell();
            //        mergedCell.AppendChild(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Continue }));
            //        currentRow.Append(mergedCell);
            //    }
            //}
            return cell;
        }
    }
}
