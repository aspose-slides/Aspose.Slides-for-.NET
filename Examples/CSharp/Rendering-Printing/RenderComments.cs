using Aspose.Slides.Export;
using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Rendering.Printing
{
    class RenderComments
    {
        public static void Run()
        {
            //ExStart:RenderComments
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Rendering();


            Presentation _document = new Presentation(dataDir+"test.pptx");
            ISlide slide = _document.Slides[pageNumber - 1];
            Size size = _document.SlideSize.Size.ToSize();

            using (System.Drawing.Bitmap image = new System.Drawing.Bitmap(size.Width, size.Height))
            {
                using (System.Drawing.Graphics graphics = System.Drawing.Graphics.FromImage(image))
                {
                    slide.RenderToGraphics(RenderNotes, graphics);
                    graphics.Save();
                }

                ImageToStream(dataDir+"image", outputStream);
            }

            }
            //ExEnd:RenderComments
        }
    }


