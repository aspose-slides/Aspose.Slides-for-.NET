using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Presentations.Saving
{
    class ConvertSvgImageObjectIntoGroupOfShapes
    {
        public static void Run() {

            //ExStart:ConvertSvgImageObjectIntoGroupOfShapes
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_PresentationSaving();

            using (Presentation pres = new Presentation(dataDir+ "image.pptx"))
            {
                PictureFrame pFrame = pres.Slides[0].Shapes[0] as PictureFrame;
                ISvgImage svgImage = pFrame.PictureFormat.Picture.Image.SvgImage;
                if (svgImage != null)
                {
                    // Convert svg image into group of shapes
                    IGroupShape groupShape = pres.Slides[0].Shapes.AddGroupShape(svgImage, pFrame.Frame.X, pFrame.Frame.Y,
                        pFrame.Frame.Width, pFrame.Frame.Height);
                    // remove source svg image from presentation
                    pres.Slides[0].Shapes.Remove(pFrame);
                }

                pres.Save(dataDir + "image_group.pptx", SaveFormat.Pptx);
            }
            //ExEnd:ConvertSvgImageObjectIntoGroupOfShapes
        }

    }
}
