using DocTool.Dto;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace DocTool.Base
{
    public class DocBase : IDocBase
    {
        /// <summary>
        /// 附件類型
        /// </summary>
        public enum fileExtensionType
        {
            pdf = 101,
            docx = 201,
            doc = 202,
            odt = 203,
            html = 301,
            htm = 302,
            xlsx = 401,
            xls = 402,
        }
        /// <summary>
        /// 取代時，需要先轉換成docx的副檔名
        /// </summary>
        protected List<string> replaceTypeConversion = new List<string>() { $".{fileExtensionType.odt}" };
        /// <summary>
        /// 是否刪除暫存附件
        /// </summary>
        public bool deleteTempFile { get; set; } = true;
        private int count { get; set; } = 0;
        public string prefixStr { get; set; } = "{$";
        public string suffixStr { get; set; } = "$}";
        /// <summary>
        /// 附件資料
        /// </summary>
        protected FileObj fileData { get; set; }
        /// <summary>
        /// LibreOffice路徑
        /// </summary>
        public string libreOfficeAppPath { get; set; }
        /// <summary>
        /// 暫存檔案路徑
        /// </summary>
        public string locationTempPath { get; set; }
        /// <summary>
        /// 建構
        /// </summary>
        public DocBase(string libreOfficeAppPath, string locationTempPath)
        {
            this.libreOfficeAppPath = libreOfficeAppPath;
            this.locationTempPath = locationTempPath;
            Directory.CreateDirectory(this.locationTempPath);
        }
        /// <summary>
        /// 轉換主檔
        /// </summary>
        private void DocConvertMain(string inputPath, string outputFolderPath, fileExtensionType outputType)
        {
            try
            {
                var commandArgs = new List<string>() { "--convert-to" };
                var fileExtension = Path.GetExtension(inputPath);
                //轉pdf
                if (outputType == fileExtensionType.pdf)
                {
                    var acceptTypeDatas = new List<string>()
                    {
                        $".{fileExtensionType.pdf}",
                        $".{fileExtensionType.htm}",
                        $".{fileExtensionType.html}",
                        $".{fileExtensionType.xls}",
                        $".{fileExtensionType.xlsx}",
                        $".{fileExtensionType.doc}",
                        $".{fileExtensionType.docx}",
                        $".{fileExtensionType.odt}",
                    };
                    if (acceptTypeDatas.Contains(fileExtension))
                    {
                        commandArgs.Add("pdf:writer_pdf_Export");
                    }
                    else
                    {
                        throw new Exception($"outputType:{outputType}. file extension \"{fileExtension}\" not supported!");
                    }
                }
                //轉html
                else if (outputType == fileExtensionType.html || outputType == fileExtensionType.htm)
                {
                    var acceptTypeDatas = new List<string>()
                    {
                        $".{fileExtensionType.htm}",
                        $".{fileExtensionType.html}",
                        $".{fileExtensionType.doc}",
                        $".{fileExtensionType.docx}",
                        $".{fileExtensionType.odt}",
                    };
                    if (acceptTypeDatas.Contains(fileExtension))
                    {
                        commandArgs.Add("html:HTML:EmbedImages");
                    }
                    else
                    {
                        throw new Exception($"outputType:{outputType}. file extension \"{fileExtension}\" not supported!");
                    }
                }
                //轉docx
                else if (outputType == fileExtensionType.doc || outputType == fileExtensionType.docx)
                {
                    var acceptTypeDatas = new List<string>()
                    {
                        $".{fileExtensionType.doc}",
                        $".{fileExtensionType.docx}",
                        $".{fileExtensionType.odt}",
                        $".{fileExtensionType.htm}",
                        $".{fileExtensionType.html}",
                    };
                    if (acceptTypeDatas.Contains(fileExtension))
                    {
                        commandArgs.Add("docx");
                    }
                    else
                    {
                        throw new Exception($"outputType:{outputType}. file extension \"{fileExtension}\" not supported!");
                    }
                }
                //轉odt
                else if (outputType == fileExtensionType.odt)
                {
                    var acceptTypeDatas = new List<string>()
                    {
                        $".{fileExtensionType.doc}",
                        $".{fileExtensionType.docx}",
                        $".{fileExtensionType.odt}",
                        $".{fileExtensionType.htm}",
                        $".{fileExtensionType.html}",
                    };
                    if (acceptTypeDatas.Contains(fileExtension))
                    {
                        commandArgs.Add("odt");
                    }
                    else
                    {
                        throw new Exception($"outputType:{outputType}. file extension \"{fileExtension}\" not supported!");
                    }
                }
                commandArgs.AddRange(new[] { inputPath, "--norestore", "--writer", "--headless", "--outdir", outputFolderPath });
                var procStartInfo = new ProcessStartInfo(this.libreOfficeAppPath);
                procStartInfo.Arguments = String.Join(" ", commandArgs);
                procStartInfo.RedirectStandardOutput = true;
                procStartInfo.UseShellExecute = false;
                procStartInfo.CreateNoWindow = true;
                procStartInfo.WorkingDirectory = Environment.CurrentDirectory;
                var process = new Process() { StartInfo = procStartInfo };
                var pname = Process.GetProcessesByName("soffice");
                //無法執行：可先開啟工作管理員 => 詳細資料 => 找到soffice => 結束工作
                var waitTimeCount = 0;
                while (pname.Length > 0)
                {
                    Thread.Sleep(5000);
                    pname = Process.GetProcessesByName("soffice");
                    waitTimeCount++;
                    if (waitTimeCount > 5)
                        throw new Exception("time out");
                }
                process.Start();
                process.WaitForExit(20000);
                this.fileData = new FileObj()
                {
                    fileName = $"{Path.GetFileNameWithoutExtension(inputPath)}",
                    fileByteArr = File.ReadAllBytes(Path.Combine(outputFolderPath, $"{Path.GetFileNameWithoutExtension(inputPath)}.{outputType.ToString()}")),
                    fileType = outputType.ToString(),
                };
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        /// <summary>
        /// 轉換-位元組陣列
        /// </summary>
        protected void DocConvert(string inputFileNameWithExtension, byte[] inputBytes, fileExtensionType outputType)
        {
            var inputFolderPath = Path.Combine(this.locationTempPath, NewOutputFolderPath("input"));
            var inputPath = Path.Combine(inputFolderPath, inputFileNameWithExtension);
            try
            {
                Directory.CreateDirectory(this.locationTempPath);
                Directory.CreateDirectory(inputFolderPath);
                File.WriteAllBytes(inputPath, inputBytes);
                DocConvert(inputPath, outputType);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                ClearDirectory(inputFolderPath);
            }
        }
        /// <summary>
        /// 轉換-路徑
        /// </summary>
        protected void DocConvert(string inputPath, fileExtensionType outputType)
        {
            var outputFolderPath = Path.Combine(this.locationTempPath, NewOutputFolderPath("output"));
            try
            {
                if (!Directory.Exists(this.locationTempPath))
                    Directory.CreateDirectory(this.locationTempPath);
                Directory.CreateDirectory(outputFolderPath);
                DocConvertMain(inputPath, outputFolderPath, outputType);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                ClearDirectory(outputFolderPath);
            }
        }
        /// <summary>
        /// 清除暫存資料夾
        /// </summary>
        protected void ClearDirectory(string folderName)
        {
            if (this.deleteTempFile)
            {
                var dir = new DirectoryInfo(folderName);
                foreach (FileInfo fi in dir.GetFiles())
                    fi.Delete();
                foreach (DirectoryInfo di in dir.GetDirectories())
                {
                    ClearDirectory(di.FullName);
                    di.Delete();
                }
                Directory.Delete(folderName);
            }
        }
        /// <summary>
        /// 取得附件資料
        /// </summary>
        public FileObj GetData()
        {
            return this.fileData;
        }
        protected string NewOutputFolderPath(string frontName = "")
        {
            var folderSubName = DateTime.Now.Ticks.ToString();
            return $"{frontName}_{this.count++}_{folderSubName}";
        }
    }
}
