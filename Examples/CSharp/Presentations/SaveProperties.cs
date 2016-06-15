
using System.IO;
using Aspose.Slides;

namespace Aspose.Slides.Examples.CSharp.Presentations
{
    public class SaveProperties
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Presentations();

            //Instantiate a Presentation object that represents a PPT file
            Presentation presentation = new Presentation();

            //....do some work here.....

            //Setting access to document properties in password protected mode
            presentation.ProtectionManager.EncryptDocumentProperties = false;

            //Setting Password
            presentation.ProtectionManager.Encrypt("pass");

            //Save your presentation to a file
            presentation.Save(dataDir + "Password Protected Presentation.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
        }
    }
}