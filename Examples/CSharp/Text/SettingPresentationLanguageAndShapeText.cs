using System.IO;

using Aspose.Slides;

namespace Aspose.Slides.Examples.CSharp.Text
{
    public class TextBoxOnSlideProgram
    {
        public static void Run()
        {
            // ExStart:SettingPresentationLanguageAndShapeText
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();

         using (Presentation pres = new Presentation(dataDir+"test0.pptx"))
            {
                IAutoShape shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 50);
                shape.AddTextFrame("Text to apply spellcheck language");
                shape.TextFrame.Paragraphs[0].Portions[0].PortionFormat.LanguageId = "en-EN";

                pres.Save(dataDir+"test1.pptx",Aspose.Slides.Export.SaveFormat.Pptx);
           
            }
          
            }
            // ExEnd:SettingPresentationLanguageAndShapeText
        }
    }
}