using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocTool.Dto;
using DocumentFormat.OpenXml.Drawing;

namespace DocTool.DocType
{
    public class DocImage
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
        public double imageDpi { get; set; }
        public double imageWidth { get; set; }
        public double imageHeight { get; set; }
        public string imageName { get; set; }
        public PartTypeInfo imageType
        {
            get
            {
                switch (fileType.ToLower())
                {
                    case "jpg":
                        return ImagePartType.Jpeg;
                    case "png":
                        return ImagePartType.Png;
                    case "gif":
                        return ImagePartType.Gif;
                    case "bmp":
                        return ImagePartType.Bmp;
                    default:
                        return ImagePartType.Jpeg;
                }
            }
        }
        private const decimal INCH_TO_CM = 2.54M;
        private const decimal CM_TO_EMU = 360000M;
        public decimal Width;
        public decimal Height;
        public long WidthInEMU => Convert.ToInt64(Width * CM_TO_EMU);
        public long HeightInEMU => Convert.ToInt64(Height * CM_TO_EMU);
        public DocImage(string fileName, string fileType, byte[] fileByteArr, double imageDpi = 300, double imageWidth = 0, double imageHeight = 0)
        {
            this.fileName = fileName;
            this.fileType = fileType;
            this.fileByteArr = fileByteArr;
            this.imageDpi = imageDpi;
            this.imageName = $"IMG_{Guid.NewGuid().ToString().Substring(0, 8)}";
            this.imageWidth = imageWidth;
            this.imageHeight = imageHeight;
            if (this.imageDpi < 1)
                this.imageDpi = 300;
            if (this.imageWidth < 1)
                this.imageWidth = 150;
            if (this.imageHeight < 1)
                this.imageHeight = 30;
        }
        public DocImage(string filePath, double imageDpi = 72, double imageWidth = 0, double imageHeight = 0)
        {
            this.fileName = System.IO.Path.GetFileNameWithoutExtension(filePath);
            this.fileType = System.IO.Path.GetExtension(filePath).Substring(1);
            this.fileByteArr = File.ReadAllBytes(filePath);
            this.imageDpi = imageDpi;
            this.imageName = $"IMG_{Guid.NewGuid().ToString().Substring(0, 8)}";
            this.imageWidth = imageWidth;
            this.imageHeight = imageHeight;
            if (imageDpi < 1)
                this.imageDpi = 72;
            if (this.imageWidth < 1)
                this.imageWidth = 150;
            if (this.imageHeight < 1)
                this.imageHeight = 30;
        }
        public Drawing GetImageElement(string relationshipId)
        {
            double englishMetricUnitsPerInch = 914400;
            // 設置圖片的寬度和高度
            long widthEmu = (long)(this.imageWidth * englishMetricUnitsPerInch / this.imageDpi);
            long heightEmu = (long)(this.imageHeight * englishMetricUnitsPerInch / this.imageDpi);
            var imgElement = new Drawing(
                new DW.Inline(
                    new DW.Extent { Cx = widthEmu, Cy = heightEmu },
                    new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                    new DW.DocProperties { Id = (UInt32Value)1U, Name = this.imageName },
                    new DW.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks { NoChangeAspect = true }),
                    new A.Graphic(
                        new A.GraphicData(
                            new PIC.Picture(
                                new PIC.NonVisualPictureProperties(
                                    new PIC.NonVisualDrawingProperties { Id = (UInt32Value)0U, Name = this.fileName },
                                    new PIC.NonVisualPictureDrawingProperties()),
                                new PIC.BlipFill(
                                    new A.Blip(
                                        new A.BlipExtensionList(
                                            new A.BlipExtension { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" }))
                                    {
                                        Embed = relationshipId,
                                        CompressionState = A.BlipCompressionValues.Print,
                                    },
                                            new A.Stretch(new A.FillRectangle())),
                                new PIC.ShapeProperties(
                                    new A.Transform2D(
                                        new A.Offset { X = 0L, Y = 0L },
                                        new A.Extents { Cx = widthEmu, Cy = heightEmu }),
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
            return imgElement;
        }
        public Graphic GetImageGraphic(string relationshipId = "")
        {
            double englishMetricUnitsPerInch = 914400;
            // 設置圖片的寬度和高度
            long widthEmu = (long)(this.imageWidth * englishMetricUnitsPerInch / this.imageDpi);
            long heightEmu = (long)(this.imageHeight * englishMetricUnitsPerInch / this.imageDpi);
            var graphic = new Graphic();
            graphic.GraphicData = new GraphicData(new PIC.Picture(
                new PIC.NonVisualPictureProperties(
                    new PIC.NonVisualDrawingProperties() { Id = 1, Name = this.fileName },
                    new PIC.NonVisualPictureDrawingProperties()),
                new PIC.BlipFill(
                    new A.Blip() { Embed = "rId1", CompressionState = A.BlipCompressionValues.Print },
                    new A.Stretch(new A.FillRectangle())),
                new PIC.ShapeProperties(new A.Transform2D(
                    new A.Offset() { X = 0L, Y = 0L },
                    new A.Extents() { Cx = widthEmu, Cy = heightEmu }),
                    new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle })));
            return graphic;
        }
        public string FeedImgData(WordprocessingDocument doc)
        {
            var mainPart = doc.MainDocumentPart;
            var imagePart = doc.MainDocumentPart.AddImagePart(this.imageType);
            imagePart.FeedData(new MemoryStream(this.fileByteArr));
            return mainPart.GetIdOfPart(imagePart);
        }
    }
}
