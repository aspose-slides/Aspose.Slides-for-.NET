using System;
using System.IO;

/*
The following code sample demonstrates how to use the SplitTextByColumns method.
*/

namespace Aspose.Slides.Examples.CSharp.Text
{
    public class SplitTextByColumnsExample
    {
        public static void Run()
        {
            string presentationName = Path.Combine(RunExamples.GetDataDir_Text(), "MultiColumnText.pptx");

            using (Presentation pres = new Presentation(presentationName))
            {
                // Get the first shape on the slide
                IAutoShape shape = pres.Slides[0].Shapes[0] as AutoShape;
                // Get textFrame
                ITextFrame textFrame = shape.TextFrame;
                // Split the text frame content into columns
                string[] columnsText = textFrame.SplitTextByColumns();
                // Print each column's text to the console
                foreach (string column in columnsText)
                    Console.WriteLine(column + '\n');
            }
        }
    }
}