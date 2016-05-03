using Aspose.Slides;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Working_With_Size_and_Layout
{
    class Program
    {
        static void Main(string[] args)
        {
            string Path = @"E:\Aspose\Aspose Vs OpenXML\Files\";

            //Instantiate a Presentation object that represents a presentation file 
            Presentation presentation = new Presentation(Path + "render.pptx");
            Presentation auxPresentation = new Presentation();

            ISlide slide = presentation.Slides[0];

            //Set the slide size of generated presentations to that of source
            auxPresentation.SlideSize.Type = presentation.SlideSize.Type;
            auxPresentation.SlideSize.Size = presentation.SlideSize.Size;

            auxPresentation.Slides.InsertClone(0, slide);
            auxPresentation.Slides.RemoveAt(0);

            //Save Presentation to disk
            auxPresentation.Save(Path + "size.pptx", Aspose.Slides.Export.SaveFormat.Pptx);

        }
    }
}
