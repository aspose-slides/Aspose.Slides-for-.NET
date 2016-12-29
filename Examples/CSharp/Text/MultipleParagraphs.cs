using System.IO;

using Aspose.Slides;
using System.Drawing;
using Aspose.Slides.Export;

namespace Aspose.Slides.Examples.CSharp.Text
{
    public class MultipleParagraphs
    {
        public static void Run()
        {
            // ExStart:MultipleParagraphs
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            // Instantiate a Presentation class that represents a PPTX file
            using (Presentation pres = new Presentation())
            {
                // Accessing first slide
                ISlide slide = pres.Slides[0];

                // Add an AutoShape of Rectangle type
                IAutoShape ashp = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 150, 300, 150);

                // Access TextFrame of the AutoShape
                ITextFrame tf = ashp.TextFrame;

                // Create Paragraphs and Portions with different text formats
                IParagraph para0 = tf.Paragraphs[0];
                IPortion port01 = new Portion();
                IPortion port02 = new Portion();
                para0.Portions.Add(port01);
                para0.Portions.Add(port02);

                IParagraph para1 = new Paragraph();
                tf.Paragraphs.Add(para1);
                IPortion port10 = new Portion();
                IPortion port11 = new Portion();
                IPortion port12 = new Portion();
                para1.Portions.Add(port10);
                para1.Portions.Add(port11);
                para1.Portions.Add(port12);

                IParagraph para2 = new Paragraph();
                tf.Paragraphs.Add(para2);
                IPortion port20 = new Portion();
                IPortion port21 = new Portion();
                IPortion port22 = new Portion();
                para2.Portions.Add(port20);
                para2.Portions.Add(port21);
                para2.Portions.Add(port22);

                for (int i = 0; i < 3; i++)
                    for (int j = 0; j < 3; j++)
                    {
                        tf.Paragraphs[i].Portions[j].Text = "Portion0" + j.ToString();
                        if (j == 0)
                        {
                            tf.Paragraphs[i].Portions[j].PortionFormat.FillFormat.FillType = FillType.Solid;
                            tf.Paragraphs[i].Portions[j].PortionFormat.FillFormat.SolidFillColor.Color = Color.Red;
                            tf.Paragraphs[i].Portions[j].PortionFormat.FontBold = NullableBool.True;
                            tf.Paragraphs[i].Portions[j].PortionFormat.FontHeight = 15;
                        }
                        else if (j == 1)
                        {
                            tf.Paragraphs[i].Portions[j].PortionFormat.FillFormat.FillType = FillType.Solid;
                            tf.Paragraphs[i].Portions[j].PortionFormat.FillFormat.SolidFillColor.Color = Color.Blue;
                            tf.Paragraphs[i].Portions[j].PortionFormat.FontItalic = NullableBool.True;
                            tf.Paragraphs[i].Portions[j].PortionFormat.FontHeight = 18;
                        }
                    }

                //Write PPTX to Disk
                pres.Save(dataDir + "multiParaPort_out.pptx", SaveFormat.Pptx);
            }
            // ExEnd:MultipleParagraphs
        }
    }
}