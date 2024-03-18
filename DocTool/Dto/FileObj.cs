using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocTool.Dto
{
    public class FileObj
    {
        /// <summary>
        /// 附件名稱
        /// </summary>
        public string fileName { get; set; }
        /// <summary>
        /// 附件附檔名
        /// </summary>
        public string fileType { get; set; }
        /// <summary>
        /// 附件資料
        /// </summary>
        public byte[] fileByteArr { get; set; }
    }
}
