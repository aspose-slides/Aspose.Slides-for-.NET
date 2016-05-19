// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using Aspose.Slides;
using System.IO;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, install it and then add its reference to this project. For any issues, questions or suggestions please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace Aspose.Plugins.AsposeVSOpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"..\..\..\Sample Files\";
            string srcFileName = FilePath + "Conversion.pptx";
            string destFileName = FilePath + "Creating Slide SVG Image.svg";
            
            //Instantiate a Presentation class that represents the presentation file
            using (Presentation pres = new Presentation(srcFileName))
            {

                //Access the second slide
                ISlide sld = pres.Slides[1];

                //Create a memory stream object
                MemoryStream SvgStream = new MemoryStream();

                //Generate SVG image of slide and save in memory stream
                sld.WriteAsSvg(SvgStream);
                SvgStream.Position = 0;

                //Save memory stream to file
                using (Stream fileStream = System.IO.File.OpenWrite(destFileName))
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

        }
    }
}
