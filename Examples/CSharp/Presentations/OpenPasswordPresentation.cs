 
using System.IO;

using Aspose.Slides;

namespace Aspose.Slides.Examples.CSharp.Presentations
{
    public class OpenPasswordPresentation
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Presentations();

            // creating instance of load options to set the presentation access password
            LoadOptions loadOptions = new LoadOptions();

            // Setting the access password
            loadOptions.Password = "pass";

            // Opening the presentation file by passing the file path and load options to the constructor of Presentation class
            Presentation pres = new Presentation(dataDir + "OpenPasswordPresentation.pptx", loadOptions);

            // Printing the total number of slides present in the presentation
            System.Console.WriteLine(pres.Slides.Count.ToString());
        }
    }
}