using Aspose.Slides.Export;
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
    class ReplaceFontsExplicitly
    {
        public static void Run()
        {
            // ExStart:ReplaceFontsExplicitly
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();
           
            // ExStart:ReplaceFontsExplicitly
            // Load presentation
            Presentation presentation = new Presentation(dataDir + "Fonts.pptx");

            // Load source font to be replaced
            IFontData sourceFont = new FontData("Arial");

            // Load the replacing font
            IFontData destFont = new FontData("Times New Roman");

            // Replace the fonts
            presentation.FontsManager.ReplaceFont(sourceFont, destFont);

            // ExEnd:ReplaceFontsExplicitly
            // Save the presentation
            presentation.Save(dataDir + "UpdatedFont_out.pptx", SaveFormat.Pptx);
            // ExEnd:ReplaceFontsExplicitly
        }
    }
}