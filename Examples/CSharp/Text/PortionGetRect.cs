using System;
using System.Diagnostics;
using System.Drawing;
using Aspose.Slides.Export;
using Aspose.Slides.Charts;
using Aspose.Slides;
using CSharp.Tables;


namespace Aspose.Slides.Examples.CSharp.Text
{
    class PortionGetRect
    {
        public static void Run()
        {
            string outPath = RunExamples.OutPath;

            using (Presentation pres = new Presentation())
            {
                // Create table
                ITable tbl = pres.Slides[0].Shapes.AddTable(50, 50, new double[] { 50, 70 }, new double[] { 50, 50, 50 });

                // Create paragraths
                IParagraph paragraph0 = new Paragraph();
                paragraph0.Portions.Add(new Portion("Text "));
                paragraph0.Portions.Add(new Portion("in0"));
                paragraph0.Portions.Add(new Portion(" Cell"));

                IParagraph paragraph1 = new Paragraph();
                paragraph1.Text = "On0";

                IParagraph paragraph2 = new Paragraph();
                paragraph2.Portions.Add(new Portion("Hi there "));
                paragraph2.Portions.Add(new Portion("col0"));

                ICell cell = tbl.Rows[1][1];

                // Add text into the table cell
                cell.TextFrame.Paragraphs.Clear();
                cell.TextFrame.Paragraphs.Add(paragraph0);
                cell.TextFrame.Paragraphs.Add(paragraph1);
                cell.TextFrame.Paragraphs.Add(paragraph2);

                // Add TextFrame
                IAutoShape autoShape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 400, 100, 60, 120);
                autoShape.TextFrame.Text = "Text in shape";
                autoShape.TextFrame.Paragraphs[0].ParagraphFormat.Alignment = TextAlignment.Left;

                // Getting coordinates of the left top corner of the table cell.
                double x = tbl.X + cell.OffsetX;
                double y = tbl.Y + cell.OffsetY;

                // Using IParagrap.GetRect() and IPortion.GetRect() methods in order to add frame to portions and paragraphs.
                foreach (IParagraph para in cell.TextFrame.Paragraphs)
                {
                    if (para.Text == "")
                        continue;

                    RectangleF rect = para.GetRect();
                    IAutoShape shape =
                        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle,
                            rect.X + (float)x, rect.Y + (float)y, rect.Width, rect.Height);

                    shape.FillFormat.FillType = FillType.NoFill;
                    shape.LineFormat.FillFormat.SolidFillColor.Color = Color.Yellow;
                    shape.LineFormat.FillFormat.FillType = FillType.Solid;


                    foreach (IPortion portion in para.Portions)
                    {
                        if (portion.Text.Contains("0"))
                        {
                            rect = portion.GetRect();
                            shape =
                                pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle,
                                    rect.X + (float)x, rect.Y + (float)y, rect.Width, rect.Height);

                            shape.FillFormat.FillType = FillType.NoFill;
                        }
                    }
                }

                // Add frame to AutoShape paragraphs.
                foreach (IParagraph para in autoShape.TextFrame.Paragraphs)
                {
                    RectangleF rect = para.GetRect();
                    IAutoShape shape =
                        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle,
                            rect.X + autoShape.X, rect.Y + autoShape.Y, rect.Width, rect.Height);

                    shape.FillFormat.FillType = FillType.NoFill;
                    shape.LineFormat.FillFormat.SolidFillColor.Color = Color.Yellow;
                    shape.LineFormat.FillFormat.FillType = FillType.Solid;

                }

                pres.Save(outPath + "GetRect_Out.pptx", SaveFormat.Pptx);
                Process.Start(outPath + "GetRect_Out.pptx");
            }
        }
    }
}
