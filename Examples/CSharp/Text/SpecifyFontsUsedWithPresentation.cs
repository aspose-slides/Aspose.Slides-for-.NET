using System;
using Aspose.Slides.Export;
using Aspose.Slides.Charts;
using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Text
{
    class SpecifyFontsUsedWithPresentation
    {
        public static void Run()
        {
            // ExStart:SpecifyFontsUsedWithPresentation
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();

            byte[] memoryFont1 = File.ReadAllBytes("customfonts\\CustomFont1.ttf");
            byte[] memoryFont2 = File.ReadAllBytes("customfonts\\CustomFont2.ttf");

            ILoadOptions loadOptions = new LoadOptions();
            loadOptions.DocumentLevelFontSources.FontFolders = new string[] { "assets\\fonts", "global\\fonts" };
            loadOptions.DocumentLevelFontSources.MemoryFonts = new byte[][] { memoryFont1, memoryFont2 };
            using (IPresentation presentation = CreatePresentation("MyPresentation.pptx", loadOptions))
            {
                //work with the presentation
                //CustomFont1, CustomFont2 as well as fonts from assets\fonts & global\fonts folders and their subfolders are available to the presentation
            }
            // ExEnd:SpecifyFontsUsedWithPresentation
        }
    }
}
