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
    public class MergeCell
    {
        public static void Run()
        {
            // ExStart:MergeCell
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Tables();

            // Instantiate Presentation class that represents PPTX file
            using (Presentation presentation = new Presentation())
            {

                // Access first slide
                ISlide slide = presentation.Slides[0];

                // Define columns with widths and rows with heights
                double[] dblCols = { 70, 70, 70, 70 };
                double[] dblRows = { 70, 70, 70, 70 };

                // Add table shape to slide
                ITable table = slide.Shapes.AddTable(100, 50, dblCols, dblRows);

                // Set border format for each cell
                foreach (IRow row in table.Rows)
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
                table.MergeCells(table[1, 1], table[2, 1], false);

                // Merging cells (1, 2) x (2, 2)
                table.MergeCells(table[1, 2], table[2, 2], false);

                // Merging cells (1, 2) x (2, 2)
                table.MergeCells(table[1, 1], table[1, 2], true);

                //Write PPTX to Disk
                presentation.Save(dataDir + "MergeCells1_out.pptx", SaveFormat.Pptx);
            }
            // ExEnd:MergeCell
        }
    }
}



