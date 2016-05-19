using Aspose.Slides;
using System;
using System.Collections.Generic;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, install it and then add its reference to this project. For any issues, questions or suggestions please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace Aspose.Plugins.AsposeVSOpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"..\..\..\..\Sample Files\";
            string FileName = FilePath + "Get all the External Eyperlinks.pptx";
            
            //Instantiate a Presentation object that represents a PPTX file
            Presentation pres = new Presentation(FileName);

            //Get the hyperlinks from presentation
            IList<IHyperlinkContainer> links = pres.HyperlinkQueries.GetAnyHyperlinks();

            foreach (IHyperlinkContainer link in links)
                Console.WriteLine(link.HyperlinkClick.ExternalUrl);

        }
    }
}
