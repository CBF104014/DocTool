using DocTool.DocType;
using DocTool.Dto;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocTool.Pdf
{
    public class DocPdfTool
    {
        /// <summary>
        /// 附件資料
        /// </summary>
        protected FileObj fileData { get; set; }
        /// <summary>
        /// 建構子
        /// </summary>
        public DocPdfTool() { }
        /// <summary>
        /// 設置資料
        /// </summary>
        public DocPdfTool Set(FileObj fileObjData)
        {
            this.fileData = fileObjData;
            return this;
        }
        /// <summary>
        /// 設置資料
        /// </summary>
        public DocPdfTool Set(string inputPath)
        {
            this.fileData = new FileObj()
            {
                fileName = System.IO.Path.GetFileNameWithoutExtension(inputPath),
                fileType = System.IO.Path.GetExtension(inputPath).Substring(1),
                fileByteArr = File.ReadAllBytes(inputPath),
            };
            return this;
        }
        /// <summary>
        /// 設置資料
        /// </summary>
        public DocPdfTool Set(string fileNameWithExtension, byte[] inputBytes)
        {
            this.fileData = new FileObj()
            {
                fileName = System.IO.Path.GetFileNameWithoutExtension(fileNameWithExtension),
                fileType = System.IO.Path.GetExtension(fileNameWithExtension).Substring(1),
                fileByteArr = inputBytes,
            };
            return this;
        }
        /// <summary>
        /// 合併PDF
        /// </summary>
        public byte[] Merge(List<byte[]> pdfByteContent)
        {
            using (var ms = new MemoryStream())
            {
                using (var doc = new Document(PageSize.A4, 0, 0, 0, 0))
                {
                    using (var copy = new PdfSmartCopy(doc, ms))
                    {
                        doc.Open();
                        var isline = true;
                        var isinit = false;
                        foreach (var p in pdfByteContent)
                        {
                            PdfReader.unethicalreading = true;
                            using (var reader = new PdfReader(p))
                            {
                                int n = reader.NumberOfPages; //取得頁數
                                //判斷第一頁是橫的還直的
                                var pageSize1 = reader.GetPageSizeWithRotation(1);
                                if (!isinit)
                                {
                                    isline = pageSize1.Width < pageSize1.Height;
                                    isinit = true;
                                }
                                PdfDictionary pd;
                                for (int j = 1; j <= n; j++)
                                {
                                    pd = reader.GetPageN(j);
                                    var pageSize = reader.GetPageSizeWithRotation(j);
                                    var islinetmp = pageSize.Width < pageSize.Height;
                                    if (isline != islinetmp)
                                    {
                                        pd.Put(PdfName.ROTATE, new PdfNumber(0));
                                    }
                                }
                                for (int page = 0; page < n;)
                                {
                                    copy.AddPage(copy.GetImportedPage(reader, ++page));
                                }
                                reader.Close();//關閉
                            }
                        }
                        doc.Close();
                    }
                }
                return ms.ToArray();
            }
        }
        /// <summary>
        /// 分割PDF文件
        /// </summary>
        public DocPdfTool Split(int startPage, int endPage)
        {
            var pdfByteList = new List<byte[]>();
            using (var reader = new PdfReader(this.fileData.fileByteArr))
            {
                for (int i = startPage; i <= endPage; i++)
                {
                    using (var ms = new MemoryStream())
                    {
                        var document = new Document();
                        var pdfCopyProvider = new PdfCopy(document, ms);
                        document.Open();
                        pdfCopyProvider.AddPage(pdfCopyProvider.GetImportedPage(reader, i));
                        document.Close();
                        pdfByteList.Add(ms.ToArray());
                    }
                }
            }
            this.fileData.fileByteArr = Merge(pdfByteList);
            return this;
        }
        /// <summary>
        /// 找出關鍵字位置
        /// </summary>
        public List<PDFKeywordPosition> FindKeyword(byte[] pdfByteContent, string keyword)
        {
            var AllPositionDatas = new List<PDFKeywordPosition>();
            using (var pdfReader = new PdfReader(pdfByteContent))
            {
                var numberOfPages = pdfReader.NumberOfPages;
                for (int page = 1; page <= numberOfPages; page++)
                {
                    var parser = new PdfReaderContentParser(pdfReader);
                    var strategy = new PDFTextExtractionStrategy(page, keyword);
                    parser.ProcessContent(page, strategy);
                    AllPositionDatas = AllPositionDatas.Concat(strategy.GetResult()).ToList();
                }
            }
            return AllPositionDatas;
        }
        /// <summary>
        /// 插入文字到指定位置
        /// </summary>
        public DocPdfTool AddText(string keyword, PDFText pdfText)
        {
            var positionDatas = FindKeyword(this.fileData.fileByteArr, keyword);
            using (var pdfReader = new PdfReader(this.fileData.fileByteArr))
            using (var outStream = new MemoryStream())
            {
                using (var pdfStamper = new PdfStamper(pdfReader, outStream))
                {
                    foreach (var positionItem in positionDatas)
                    {
                        var x = positionItem.positionX + pdfText.offsetRightX;
                        var y = positionItem.positionY + pdfText.offsetRightY;
                        var cb = pdfStamper.GetOverContent(positionItem.pageNum);
                        cb.BeginText();
                        cb.SetFontAndSize(BaseFont.CreateFont(), pdfText.fontSize);
                        cb.SetTextMatrix(x, y);
                        cb.ShowText(pdfText.newText);
                        cb.EndText();
                    }
                }
                this.fileData.fileByteArr = outStream.ToArray();
                return this;
            }
        }
        /// <summary>
        /// 插入圖片到指定位置
        /// </summary>
        public DocPdfTool AddImage(string keyword, PDFImage pdfImage)
        {
            var positionDatas = FindKeyword(this.fileData.fileByteArr, keyword);
            using (var pdfReader = new PdfReader(this.fileData.fileByteArr))
            using (var outStream = new MemoryStream())
            {
                using (var pdfStamper = new PdfStamper(pdfReader, outStream))
                {
                    foreach (var positionItem in positionDatas)
                    {
                        var x = positionItem.positionX + pdfImage.offsetRightX;
                        var y = positionItem.positionY + pdfImage.offsetRightY;
                        var cb = pdfStamper.GetOverContent(positionItem.pageNum);
                        cb.AddImage(pdfImage.GetImage(x, y));
                    }
                }
                this.fileData.fileByteArr = outStream.ToArray();
                return this;
            }
        }
        /// <summary>
        /// 設置PDF為保護中的檔案(唯讀)
        /// ※注意：鎖定後則不可編輯!!
        /// </summary>
        public DocPdfTool SetReadOnly(bool isReadOnly = true)
        {
            using (var pdfReader = new PdfReader(this.fileData.fileByteArr))
            using (var outStream = new MemoryStream())
            {
                PdfReader.unethicalreading = true;
                using (var stamper = new PdfStamper(pdfReader, outStream))
                {
                    stamper.MoreInfo = pdfReader.Info;
                    if (isReadOnly)
                    {
                        //唯讀
                        stamper.SetEncryption(PdfWriter.STRENGTH40BITS
                            , null
                            , null
                            , PdfWriter.AllowScreenReaders | PdfWriter.ALLOW_PRINTING);
                    }
                    else
                    {
                        //可編輯
                        stamper.SetEncryption(PdfWriter.STRENGTH40BITS
                            , null
                            , null
                            , PdfWriter.ALLOW_PRINTING | PdfWriter.ALLOW_MODIFY_CONTENTS | PdfWriter.AllowCopy | PdfWriter.ALLOW_ASSEMBLY | PdfWriter.ALLOW_MODIFY_ANNOTATIONS);
                    }
                }
                this.fileData.fileByteArr = outStream.ToArray();
            }
            return this;
        }
        /// <summary>
        /// 取得附件資料
        /// </summary>
        public FileObj GetData()
        {
            return this.fileData;
        }
    }
}
