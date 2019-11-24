using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Slides.Layout
{
    class ManageHeaderFooterText
    {
        public static void Run() {

            //ExStart:ManageHeaderFooterText

            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations_Layout();

            // Load Presentation
            Presentation pres = new Presentation(dataDir + "headerTest.pptx");

            // Setting Footer
            pres.HeaderFooterManager.SetAllFootersText("My Footer text");
            pres.HeaderFooterManager.SetAllFootersVisibility(true);

            // Access and Update Header
            IMasterNotesSlide masterNotesSlide = pres.MasterNotesSlideManager.MasterNotesSlide;
            if (null != masterNotesSlide)
            {
                UpdateHeaderFooterText(masterNotesSlide);
            }

            // Save presentation
            pres.Save(dataDir + "HeaderFooterJava.pptx", SaveFormat.Pptx);

            //ExEnd:ManageHeaderFooterText

        }

        //ExStart:UpdateHeaderFooterText
        // Method to set Header/Footer Text
        public static void UpdateHeaderFooterText(IBaseSlide master)
        {
            foreach (IShape shape in master.Shapes)
            {
                if (shape.Placeholder != null)
                {
                    if (shape.Placeholder.Type == PlaceholderType.Header)
                    {
                        ((IAutoShape)shape).TextFrame.Text = "HI there new header";
                    }
                }
            }
        }
        //ExEnd:UpdateHeaderFooterText
    }
}
