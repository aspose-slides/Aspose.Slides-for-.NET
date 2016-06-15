using System.IO;
using Aspose.Slides;

namespace Aspose.Slides.Examples.CSharp.Presentations
{
    public class AccessModifyingProperties
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Presentations();

            //Instanciate the Presentation class that represents the PPTX
            Presentation presentation = new Presentation(dataDir + "AccessModifyingProperties.pptx");

            //Create a reference to DocumentProperties object associated with Prsentation
            IDocumentProperties documentProperties = presentation.DocumentProperties;

            //Access and modify custom properties
            for (int i = 0; i < documentProperties.CountOfCustomProperties; i++)
            {
                //Display names and values of custom properties
                System.Console.WriteLine("Custom Property Name : " + documentProperties.GetCustomPropertyName(i));
                System.Console.WriteLine("Custom Property Value : " + documentProperties[documentProperties.GetCustomPropertyName(i)]);

                //Modify values of custom properties
                documentProperties[documentProperties.GetCustomPropertyName(i)] = "New Value " + (i + 1);
            }

            //Save your presentation to a file
            presentation.Save(dataDir + "CustomDemoModified.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
        }
    }
}