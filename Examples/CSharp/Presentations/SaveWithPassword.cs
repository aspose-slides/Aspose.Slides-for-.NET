using System.IO;

using Aspose.Slides;

namespace Aspose.Slides.Examples.CSharp.Presentations
{
    public class SaveWithPassword
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Presentations();

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            //Instantiate a Presentation object that represents a PPT file
            Presentation pres = new Presentation();

            //....do some work here.....

            //Setting Password
            pres.ProtectionManager.Encrypt("pass");

            //Save your presentation to a file
            pres.Save(dataDir + "demoPass.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
        }
    }
}