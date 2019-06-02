using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Presentations.Saving
{
    class AddImageFromSVGObject
    {
        public static void Run() {

            //ExStart:AddImageFromSVGObject
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_PresentationSaving();
            string svgPath = dataDir + "sample.svg";
            string outPptxPath = dataDir + "presentation.pptx";
            using (var p = new Presentation())
            {
                string svgContent = File.ReadAllText(svgPath);
                ISvgImage svgImage = new SvgImage(svgContent);
                IPPImage ppImage = p.Images.AddImage(svgImage);
                p.Slides[0].Shapes.AddPictureFrame(ShapeType.Rectangle, 0, 0, ppImage.Width, ppImage.Height, ppImage);
                p.Save(outPptxPath, SaveFormat.Pptx);
            }

            //ExEnd:AddImageFromSVGObject
        }
    }
}
