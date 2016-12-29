using System.Drawing;
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
    public class CloningInTable
    {
        public static void Run()
        {
            // ExStart:CloningInTable
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Tables();

            // Instantiate presentationentation class that representationents PPTX file
            using (Presentation presentation = new Presentation())
            {
                // Access first slide
                ISlide sld = presentation.Slides[0];

                // Define columns with widths and rows with heights
                double[] dblCols = { 50, 50, 50 };
                double[] dblRows = { 50, 30, 30, 30, 30 };

                // Add table shape to slide
                ITable table = sld.Shapes.AddTable(100, 50, dblCols, dblRows);

                // Set border format for each cell
                foreach (IRow row in table.Rows)
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

                // Merge cells 1 & 2 of row 1
                table.MergeCells(table[0, 0], table[1, 0], false);

                // Add text to the merged cell
                table[0, 0].TextFrame.Text = "Merged Cells";

                // Write PPTX to Disk
                presentation.Save(dataDir + "table_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
            }
            // ExEnd:CloningInTable
        }
    }
}