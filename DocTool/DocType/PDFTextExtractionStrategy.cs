using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocTool.DocType
{
    public class PDFTextExtractionStrategy : ITextExtractionStrategy
    {
        /// <summary>
        /// 目前頁數
        /// </summary>
        public int pageNum { get; set; }
        /// <summary>
        /// 找尋的關鍵字
        /// </summary>
        public string keyword { get; set; }
        /// <summary>
        /// 是否模糊搜尋
        /// </summary>
        public bool isFuzzySearch { get; set; }
        /// <summary>
        /// 尋找到的位置資訊
        /// </summary>
        public List<PDFKeywordPosition> keywordLocations { get; set; }
        /// <summary>
        /// 建構子
        /// </summary>
        /// <param name="keyword">關鍵字</param>
        public PDFTextExtractionStrategy(int pageNum, string keyword)
        {
            this.pageNum = pageNum;
            this.keyword = keyword;
            this.isFuzzySearch = true;
            this.keywordLocations = new List<PDFKeywordPosition>();
        }
        public void BeginTextBlock() { }
        public void EndTextBlock() { }
        public void RenderImage(ImageRenderInfo renderInfo) { }
        public void RenderText(TextRenderInfo renderInfo)
        {
            if (isFuzzySearch)
            {
                if (renderInfo.GetText().Contains(keyword))
                {
                    AddToList(renderInfo);
                }
            }
            else
            {
                if (renderInfo.GetText().Equals(keyword))
                {
                    AddToList(renderInfo);
                }
            }
        }
        public void AddToList(TextRenderInfo renderInfo)
        {
            this.keywordLocations.Add(new PDFKeywordPosition()
            {
                pageNum = this.pageNum,
                textStr = renderInfo.GetText(),
                positionX = renderInfo.GetBaseline().GetBoundingRectange().X,
                positionY = renderInfo.GetBaseline().GetBoundingRectange().Y,
            });
        }
        public string GetResultantText()
        {
            if (this.keywordLocations == null)
                return null;
            return String.Join("、", this.keywordLocations);
        }
        public List<PDFKeywordPosition> GetResult()
        {
            return this.keywordLocations;
        }
    }
}
