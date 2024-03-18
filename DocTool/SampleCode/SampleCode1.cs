using DocTool.Dto;
using DocTool.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocTool.SampleCode
{
    public class SampleCode1
    {
        private DocWordTool wordTool { get; set; }
        public SampleCode1(string libreOfficeAppPath, string locationTempPath)
        {
            this.wordTool = new DocWordTool(libreOfficeAppPath, locationTempPath);
        }
        /// <summary>
        /// 轉PDF範例
        /// </summary>
        public FileObj WordToPdf(FileObj fileData)
        {
            fileData = wordTool
                .ToPDF($"{fileData.fileName}.{fileData.fileType}", fileData.fileByteArr)
                .GetData();
            return fileData;
        }
        /// <summary>
        /// 替換關鍵字
        /// </summary>
        public FileObj WordReplaceTag(FileObj fileData, Dictionary<string, ReplaceDto> replaceData = null)
        {
            if (replaceData == null)
                replaceData = GetSampleReplaceData();
             fileData = wordTool
                .ReplaceTag($"{fileData.fileName}.{fileData.fileType}", fileData.fileByteArr, replaceData)
                .GetData();
            return fileData;
        }
        private Dictionary<string, ReplaceDto> GetSampleReplaceData()
        {
            var cout = 10;
            var rowDatas = new List<List<string>>();
            for (int i = 0; i < cout; i++)
            {
                rowDatas.Add(new List<string>() { "AAA" + i, (i + i).ToString(), ((i + i) * 3).ToString() });
            }
            var tableRowDatas = rowDatas.Select(x => {
                var row = wordTool.CreateRow();
                x.ForEach(item => row.Append(wordTool.CreateCell(item)));
                return row;
            }).ToList();
            //測試
            var Dict = new Dictionary<string, ReplaceDto>()
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
                    tableData = this.wordTool.ToXMLTable(new List<string>() {
                        "品名", "數量", "總價"
                    }, new List<List<string>>() { new List<string>() { "A", "1", "1500" } }),
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
                    .Select(x => {
                        var row = wordTool.CreateRow();
                        x.ForEach(item => row.Append(wordTool.CreateCell(item)));
                        return row;
                    }).ToList()
                }
            };
            return Dict;
        }
    }
}
