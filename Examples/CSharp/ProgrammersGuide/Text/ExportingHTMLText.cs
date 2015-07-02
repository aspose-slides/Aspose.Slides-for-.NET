//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;
using System.Text;

namespace CSharp.Text
{
    public class ExportingHTMLText
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();


            //Load the presentation file
            using (Presentation pres = new Presentation(dataDir + "ExportingHTMLText.pptx"))
            {

                //Acesss the default first slide of presentation
                ISlide slide = pres.Slides[0];

                //Desired index
                int index = 0;

                //Accessing the added shape
                IAutoShape ashape = (IAutoShape)slide.Shapes[index];

                // Extracting first paragraph as HTML
                StreamWriter sw = new StreamWriter(dataDir + "output.html", false, Encoding.UTF8);

                //Writing Paragraphs data to HTML by providing paragraph starting index, total paragraphs to be copied
                sw.Write(ashape.TextFrame.Paragraphs.ExportToHtml(0, ashape.TextFrame.Paragraphs.Count, null));

                sw.Close();
            }

            
        }
    }
}