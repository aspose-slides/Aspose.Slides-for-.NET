using System;

/*
PowerPoint behaves specifically for links and their corresponding text in a portion. It allows to create text for the hyperlink 
in the form of a valid URL, different from the real address of the link. In this case, when you view the link in the edit window, 
it will be changed to match the text portion. 

Next example shows how to get as valid so real hyperlinks.
*/

namespace Aspose.Slides.Examples.CSharp.Slides.Hyperlinks
{
    class ExternalUrlOriginal
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations_Hyperlink();

            // Instantiate Presentation class
            using (var presentation = new Presentation(dataDir + "ExternalUrlOriginal.pptx"))
            {
                var portion = ((AutoShape)presentation.Slides[0].Shapes[1]).TextFrame.Paragraphs[0].Portions[0];

                // Shows fake hyperlink
                Console.WriteLine("Fake External Hyperlink : {0}", portion.PortionFormat.AsIHyperlinkContainer.HyperlinkClick.ExternalUrl);

                // Shows real hyperlink
                Console.WriteLine("Real External Hyperlink : {0}", portion.PortionFormat.AsIHyperlinkContainer.HyperlinkClick.ExternalUrlOriginal);
            }
        }
    }
}