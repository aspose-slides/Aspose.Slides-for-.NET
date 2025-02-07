using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using System;
using System.IO;
using Aspose.Slides.Export;
using Aspose.Slides.LowCode;

/*
The following code sample demonstrates how to use the DisableFontLigatures property to disable ligatures during export.
*/

namespace CSharp.Text
{
    class DisableFontLigaturesExample
    {
        public static void Run()
        {
            string presentationName = Path.Combine(RunExamples.GetDataDir_Text(), "TextLigatures.pptx");
            string outPathEnabled = Path.Combine(RunExamples.OutPath, "EnableLigatures-out.html");
            string outPathDisabled = Path.Combine(RunExamples.OutPath, "DisableLigatures-out.html");

            using (Presentation pres = new Presentation(presentationName))
            {
                // Save with enabled ligatures
                pres.Save(outPathEnabled, SaveFormat.Html);
                
                // Configure export options
                HtmlOptions options = new HtmlOptions
                {
                    DisableFontLigatures = true // Disable ligatures in rendered text
                };

                // Export presentation to HTML with disabled ligatures
                pres.Save(outPathDisabled, SaveFormat.Html, options);
            }
        }
    }
}
