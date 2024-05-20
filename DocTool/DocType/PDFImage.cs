using iTextSharp.text;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocTool.DocType
{
    public class PDFImage
    {
        public PDFImage(byte[] imageByteArr, float imgPercent = 20, float offsetRightX = 0, float offsetRightY = 0)
        {
            this.imageByteArr = imageByteArr;
            this.imgPercent = imgPercent;
            this.offsetRightX = offsetRightX;
            this.offsetRightY = offsetRightY;
        }
        public PDFImage(string imagePath, float imgPercent = 20, float offsetRightX = 0, float offsetRightY = 0)
        {
            this.imageByteArr = File.ReadAllBytes(imagePath);
            this.imgPercent = imgPercent;
            this.offsetRightX = offsetRightX;
            this.offsetRightY = offsetRightY;
        }
        public byte[] imageByteArr { get; set; }
        public float imgPercent { get; set; }
        public float offsetRightX { get; set; }
        public float offsetRightY { get; set; }
        public Image GetImage(float x, float y)
        {
            var image = Image.GetInstance(this.imageByteArr);
            image.ScalePercent(this.imgPercent);
            image.SetAbsolutePosition(x, y);
            return image;
        }
    }
}
