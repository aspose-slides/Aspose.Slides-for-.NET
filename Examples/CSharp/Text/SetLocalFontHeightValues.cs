using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Text
{
    class SetLocalFontHeightValues
    {
        public static void Run() {


            //ExStart:SetLocalFontHeightValues
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();

            using (Presentation pres = new Presentation())
            {
                IAutoShape newShape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 400, 75, false);
                newShape.AddTextFrame("");
                newShape.TextFrame.Paragraphs[0].Portions.Clear();

                IPortion portion0 = new Portion("Sample text with first portion");
                IPortion portion1 = new Portion(" and second portion.");

                newShape.TextFrame.Paragraphs[0].Portions.Add(portion0);
                newShape.TextFrame.Paragraphs[0].Portions.Add(portion1);

                Console.WriteLine("Effective font height just after creation:");
                Console.WriteLine("Portion #0: " + portion0.PortionFormat.GetEffective().FontHeight);
                Console.WriteLine("Portion #1: " + portion1.PortionFormat.GetEffective().FontHeight);

                pres.DefaultTextStyle.GetLevel(0).DefaultPortionFormat.FontHeight = 24;

                Console.WriteLine("Effective font height after setting entire presentation default font height:");
                Console.WriteLine("Portion #0: " + portion0.PortionFormat.GetEffective().FontHeight);
                Console.WriteLine("Portion #1: " + portion1.PortionFormat.GetEffective().FontHeight);

                newShape.TextFrame.Paragraphs[0].ParagraphFormat.DefaultPortionFormat.FontHeight = 40;

                Console.WriteLine("Effective font height after setting paragraph default font height:");
                Console.WriteLine("Portion #0: " + portion0.PortionFormat.GetEffective().FontHeight);
                Console.WriteLine("Portion #1: " + portion1.PortionFormat.GetEffective().FontHeight);

                newShape.TextFrame.Paragraphs[0].Portions[0].PortionFormat.FontHeight = 55;

                Console.WriteLine("Effective font height after setting portion #0 font height:");
                Console.WriteLine("Portion #0: " + portion0.PortionFormat.GetEffective().FontHeight);
                Console.WriteLine("Portion #1: " + portion1.PortionFormat.GetEffective().FontHeight);

                newShape.TextFrame.Paragraphs[0].Portions[1].PortionFormat.FontHeight = 18;

                Console.WriteLine("Effective font height after setting portion #1 font height:");
                Console.WriteLine("Portion #0: " + portion0.PortionFormat.GetEffective().FontHeight);
                Console.WriteLine("Portion #1: " + portion1.PortionFormat.GetEffective().FontHeight);

                pres.Save(dataDir + "SetLocalFontHeightValues.pptx",SaveFormat.Pptx);
            }

            //ExEnd:SetLocalFontHeightValues

        }
    }
}
