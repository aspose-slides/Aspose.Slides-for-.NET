using System;
using System.Drawing;
using System.IO;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace Aspose.Slides.Examples.CSharp.Text
{

    // This example demonstrates retrieving bullet's fill effective data.

    class BulletFillFormatEffective
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_Text();
            string pptxFile = Path.Combine(dataDir, "BulletData.pptx");

            using (Presentation pres = new Presentation(pptxFile))
            {
                AutoShape autoShape = (AutoShape) pres.Slides[0].Shapes[0];
                foreach (Paragraph para in autoShape.TextFrame.Paragraphs)
                {
                    IBulletFormatEffectiveData bulletFormatEffective = para.ParagraphFormat.Bullet.GetEffective();
                    Console.WriteLine("Bullet type: " + bulletFormatEffective.Type);
                    if (bulletFormatEffective.Type != BulletType.None)
                    {
                        Console.WriteLine("Bullet fill type: " + bulletFormatEffective.FillFormat.FillType);
                        switch (bulletFormatEffective.FillFormat.FillType)
                        {
                            case FillType.Solid:
                                Console.WriteLine(
                                    "Solid fill color: " + bulletFormatEffective.FillFormat.SolidFillColor);
                                break;
                            case FillType.Gradient:
                                Console.WriteLine("Gradient stops count: " +
                                                  bulletFormatEffective.FillFormat.GradientFormat.GradientStops.Count);
                                foreach (IGradientStopEffectiveData gradStop in bulletFormatEffective.FillFormat
                                    .GradientFormat.GradientStops)
                                    Console.WriteLine(gradStop.Position + ": " + gradStop.Color);
                                break;
                            case FillType.Pattern:
                                Console.WriteLine("Pattern style: " +
                                                  bulletFormatEffective.FillFormat.PatternFormat.PatternStyle);
                                Console.WriteLine("Fore color: " +
                                                  bulletFormatEffective.FillFormat.PatternFormat.ForeColor);
                                Console.WriteLine("Back color: " +
                                                  bulletFormatEffective.FillFormat.PatternFormat.BackColor);
                                break;
                        }
                    }

                    Console.WriteLine();
                }
            }
        }
    }
}
