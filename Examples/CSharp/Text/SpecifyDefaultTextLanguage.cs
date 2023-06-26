using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides;

/*
The example below demonstrates using load options to define the default text culture
*/

namespace CSharp.Text
{
    class SpecifyDefaultTextLanguage
    {
        public static void Run()
        {
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.DefaultTextLanguage = "en-US";
            using (Presentation pres = new Presentation(loadOptions))
            {
                // Add new rectangle shape with text
                IAutoShape shp = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 150, 50);
                shp.TextFrame.Text = "New Text";

                // Check the first portion language
                Console.WriteLine(shp.TextFrame.Paragraphs[0].Portions[0].PortionFormat.LanguageId);
            }
        }
    }
}
