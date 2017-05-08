using System;
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
    class PrintPreview
    {
        public static void Run()
        {
            //ExStart:PrintPreview
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Rendering();
           using (Presentation pres = new Presentation())
       {
          PrinterSettings printerSettings = new PrinterSettings();
          printerSettings.Copies = 2;
          printerSettings.DefaultPageSettings.Landscape = true;
          printerSettings.DefaultPageSettings.Margins.Left = 10;
   //...etc
            pres.Print(printerSettings);
            }
                       
            }
        //ExEnd:PrintPreview
    }
}


