using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;

/*
This code example demonstrates how the SetMacroHyperlinkClick method is used to set a macro hyperlink click on a shape:
*/

namespace Aspose.Slides.Examples.CSharp.Slides.Hyperlinks
{
    public class MacroHyperlink
    {
        public static void Run()
        {
            string macroName = "TestMacro";
            using (Presentation presentation = new Presentation())
            {
                IAutoShape shape = presentation.Slides[0].Shapes.AddAutoShape(ShapeType.BlankButton, 20, 20, 80, 30);
                shape.HyperlinkManager.SetMacroHyperlinkClick(macroName);

                Console.WriteLine("External URL is {0}", shape.HyperlinkClick.ExternalUrl);
                Console.WriteLine("Shape action type is {0}", shape.HyperlinkClick.ActionType);
            }
        }
    }
}
