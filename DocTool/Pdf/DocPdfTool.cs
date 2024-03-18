using iTextSharp.text;
using iTextSharp.text.pdf;
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
        public byte[] Split(byte[] pdfByteContent, int startPage, int endPage)
        {
            var pdfByteList = new List<byte[]>();
            using (var reader = new PdfReader(pdfByteContent))
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
            return Merge(pdfByteList);
        }
    }
}
