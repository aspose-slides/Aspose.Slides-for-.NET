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
    class GetFontsFolders
    {
        public static void Run()
        {
            //ExStart:GetFontsFolders
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();

            //The following line shall return folders where font files are searched.
            //Those are folders that have been added with LoadExternalFonts method as well as system font folders.
            string[] fontFolders = FontsLoader.GetFontFolders();

            //ExEnd:GetFontsFolders
        }
    }
}
