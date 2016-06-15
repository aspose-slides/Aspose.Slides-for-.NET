using System.IO;
using Aspose.Slides;

namespace Aspose.Slides.Examples.CSharp.Presentations
{
    public class SaveAsReadOnly
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
            Presentation presentation = new Presentation();

            //....do some work here.....

            //Setting Write protection Password
            presentation.ProtectionManager.SetWriteProtection("test");

            //Save your presentation to a file
            presentation.Save(dataDir + "WriteProtected.pptx", Aspose.Slides.Export.SaveFormat.Pptx);

        }
    }
}