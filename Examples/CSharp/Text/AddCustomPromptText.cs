using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Text
{
    class AddCustomPromptText
    {
        public static void Run() {

            //ExStart:AddCustomPromptText
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();

            using (Presentation pres = new Presentation(dataDir + "Presentation2.pptx"))
            {
                ISlide slide = pres.Slides[0];
                foreach (IShape shape in slide.Slide.Shapes) // iterate through the slide
                {
                    if (shape.Placeholder != null && shape is AutoShape)
                    {
                        string text = "";
                        if (shape.Placeholder.Type == PlaceholderType.CenteredTitle) // title - the text is empty, PowerPoint displays "Click to add title". 
                        {
                            text = "Click to add custom title";
                        }
                        else if (shape.Placeholder.Type == PlaceholderType.Subtitle) // the same for subtitle.
                        {
                            text = "Click to add custom subtitle";
                        }

                        ((IAutoShape)shape).TextFrame.Text = text;

                        Console.WriteLine("Placeholder with text: {0}", text);
                    }
                }

                pres.Save(dataDir + "Placeholders_PromptText.pptx", SaveFormat.Pptx);
            }



            //ExEnd:AddCustomPromptText

        }
    }
}
