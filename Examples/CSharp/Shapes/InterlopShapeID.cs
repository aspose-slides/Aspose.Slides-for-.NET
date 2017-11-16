using System.IO;

using Aspose.Slides;
using Aspose.Slides.Export;

namespace Aspose.Slides.Examples.CSharp.Shapes
{
  class InterlopShapeID
   {  
        //ExStart:InterlopShapeID
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Shapes();

            using (Presentation presentation = new Presentation("Presentation.pptx"))
         {
            // Getting unique shape identifier in slide scope
            long officeInteropShapeId = presentation.Slides[0].Shapes[0].OfficeInteropShapeId;
   
            //ExEnd:InterlopShapeID
            }
            }
        }
    }
