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

namespace CSharp.Tables
{
    public class CloningInTable
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Tables();

            // Creating empty presentation
            Presentation presentation = new Presentation();

            // Access first slide
            ISlide slide = presentation.Slides[0];

            // Define columns with widths and rows with heights
            double[] dblCols = { 50, 50, 50 };
            double[] dblRows = { 50, 30, 30, 30 };

            // Add table shape to slide
            ITable table = slide.Shapes.AddTable(100, 50, dblCols, dblRows);

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

            table[0, 0].TextFrame.Text = "00";
            table[0, 1].TextFrame.Text = "01";
            table[0, 2].TextFrame.Text = "02";
            table[0, 3].TextFrame.Text = "03";
            table[1, 0].TextFrame.Text = "10";
            table[2, 0].TextFrame.Text = "20";
            table[1, 1].TextFrame.Text = "11";
            table[2, 1].TextFrame.Text = "21";

            // AddClone adds a row in the end of the table
            table.Rows.AddClone(table.Rows[0], false);

            // InsertClone adds a row at specific position in a table
            table.Rows.InsertClone(2, table.Rows[0], false);

            // AddClone adds a column in the end of the table
            table.Columns.AddClone(table.Columns[0], false);

            // InsertClone adds a column at specific position in a table
            table.Columns.InsertClone(2, table.Columns[0], false);

            presentation.Save(dataDir + "CloneRow.pptx", SaveFormat.Pptx);

        }
    }
}