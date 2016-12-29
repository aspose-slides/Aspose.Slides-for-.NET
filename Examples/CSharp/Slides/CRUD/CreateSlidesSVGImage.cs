using System.IO;
using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Slides.CRUD
{
    public class CreateSlidesSVGImage
    {
        public static void Run()
        {
            //ExStart:CreateSlidesSVGImage
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations_CRUD();

            // Instantiate a Presentation class that represents the presentation file

            using (Presentation pres = new Presentation(dataDir + "CreateSlidesSVGImage.pptx"))
            {

                // Access the first slide
                ISlide sld = pres.Slides[0];

                // Create a memory stream object
                MemoryStream SvgStream = new MemoryStream();

                // Generate SVG image of slide and save in memory stream
                sld.WriteAsSvg(SvgStream);
                SvgStream.Position = 0;

                // Save memory stream to file
                using (Stream fileStream = System.IO.File.OpenWrite(dataDir + "Aspose_out.svg"))
                {
                    byte[] buffer = new byte[8 * 1024];
                    int len;
                    while ((len = SvgStream.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        fileStream.Write(buffer, 0, len);
                    }

                }
                SvgStream.Close();
            }
            //ExEnd:CreateSlidesSVGImage
        }
    }
}