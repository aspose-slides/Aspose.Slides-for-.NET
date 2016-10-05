using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace Aspose.Slides.Examples.CSharp.Presentations
{
    public class ManageEmbeddedFonts
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Presentations();

            // Instantiate a Presentation object that represents a presentation file
            using (Presentation presentation = new Presentation(dataDir + "EmbeddedFonts.pptx"))
            {
                // render a slide that contains a text frame that uses embedded "FunSized"
                presentation.Slides[0].GetThumbnail(new Size(960, 720)).Save(dataDir + "picture1_out.png", ImageFormat.Png);

                IFontsManager fontsManager = presentation.FontsManager;

                // get all embedded fonts
                IFontData[] embeddedFonts = fontsManager.GetEmbeddedFonts();

                // find "FunSized" font
                IFontData funSizedEmbeddedFont = Array.Find(embeddedFonts, delegate(IFontData data)
                {
                    return data.FontName == "Calibri";
                });

                // remove "Calibri" font
                fontsManager.RemoveEmbeddedFont(funSizedEmbeddedFont);

                // render the presentation; removed "Calibri" font is replaced to an existing one
                presentation.Slides[0].GetThumbnail(new Size(960, 720)).Save(dataDir + "picture2_out.png", ImageFormat.Png);

                // save the presentation without embedded "Calibri" font
                presentation.Save(dataDir + "WithoutManageEmbeddedFonts_out.ppt", SaveFormat.Ppt);
            }
        }
    }
}