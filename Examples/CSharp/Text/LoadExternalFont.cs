using System.IO;

using Aspose.Slides;
using System.Drawing;
using Aspose.Slides.Export;

namespace Aspose.Slides.Examples.CSharp.Text
{
    public class LoadExternalFont

    {
        public static void Run()
        {
            // ExStart:LoadExternalFont
            
            // The path to the documents directory.
          
            string dataDir = RunExamples.GetDataDir_Text();

           
           // loading presentation uses SomeFont which is not installed on the system
              using(Presentation pres = new Presentation("pres.pptx")
            {
         // load SomeFont from file into the byte array
              byte[] fontData = File.ReadAllBytes(@"fonts\SomeFont.ttf");

       // load font represented as byte array
             FontsLoader.LoadExternalFont(fontData);

       // font SomeFont will be available during the rendering or other operations
}
           
            // ExEnd:LoadExternalFont
       
        }
    }
}
