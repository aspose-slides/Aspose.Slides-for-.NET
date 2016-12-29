using System.Drawing;
using Aspose.Slides.Export;
using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Tables
{
    public class MergeCells
    {
        public static void Run()
        {
            // ExStart:MergeCells
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Tables();

            // Instantiate Presentation class that represents PPTX file
            using (Presentation presentation = new Presentation())
            {
                // Access first slide
                ISlide sld = presentation.Slides[0];

                // Define columns with widths and rows with heights
                double[] dblCols = { 70, 70, 70, 70 };
                double[] dblRows = { 70, 70, 70, 70 };

                // Add table shape to slide
                ITable tbl = sld.Shapes.AddTable(100, 50, dblCols, dblRows);

                // Set border format for each cell
                foreach (IRow row in tbl.Rows)
                {
                    foreach (ICell cell in row)
                    {
                        cell.BorderTop.FillFormat.FillType = FillType.Solid;
                        cell.BorderTop.FillFormat.SolidFillColor.Color = Color.Red;
                        cell.BorderTop.Width = 5;

                        cell.BorderBottom.FillFormat.FillType = FillType.Solid;
                        cell.BorderBottom.FillFormat.SolidFillColor.Color = Color.Red;
                        cell.BorderBottom.Width = 5;

                        cell.BorderLeft.FillFormat.FillType = FillType.Solid;
                        cell.BorderLeft.FillFormat.SolidFillColor.Color = Color.Red;
                        cell.BorderLeft.Width = 5;

                        cell.BorderRight.FillFormat.FillType = FillType.Solid;
                        cell.BorderRight.FillFormat.SolidFillColor.Color = Color.Red;
                        cell.BorderRight.Width = 5;
                    }
                }

                // Merging cells (1, 1) x (2, 1)
                tbl.MergeCells(tbl[1, 1], tbl[2, 1], false);

                // Merging cells (1, 2) x (2, 2)
                tbl.MergeCells(tbl[1, 2], tbl[2, 2], false);

                presentation.Save(dataDir + "MergeCells_out.pptx", SaveFormat.Pptx);
                // ExEnd:MergeCells

            }
        }
    }
}

