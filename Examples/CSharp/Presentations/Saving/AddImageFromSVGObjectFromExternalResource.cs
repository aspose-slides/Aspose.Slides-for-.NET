using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using Aspose.Slides.Import;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Presentations.Saving
{
    class AddImageFromSVGObjectFromExternalResource
    {
        public static void Run() {

            //ExStart:AddImageFromSVGObjectFromExternalResource
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_PresentationSaving();
            string outPptxPath = dataDir + "presentation_external.pptx";

            using (var p = new Presentation())
            {
                string svgContent = File.ReadAllText(new Uri(new Uri(dataDir), "image1.svg").AbsolutePath);
                ISvgImage svgImage = new SvgImage(svgContent, new ExternalResourceResolver(), dataDir);
                IPPImage ppImage = p.Images.AddImage(svgImage);
                p.Slides[0].Shapes.AddPictureFrame(ShapeType.Rectangle, 0, 0, ppImage.Width, ppImage.Height, ppImage);
                p.Save(outPptxPath, SaveFormat.Pptx);
            }

            //ExEnd:AddImageFromSVGObjectFromExternalResource
        }
    }
}
