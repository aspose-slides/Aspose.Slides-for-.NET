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
            //ExStart:CloningInTable
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Tables();

             // Instantiate presentationentation class that representationents PPTX file
            using (Presentation presentation = new Presentation(dataDir+"Test.pptx"))
            {
                // Access first slide
                ISlide sld = presentation.Slides[0];

                // Define columns with widths and rows with heights
                double[] dblCols = { 50, 50, 50 };
                double[] dblRows = { 50, 30, 30, 30, 30 };

                // Add table shape to slide
                ITable table = sld.Shapes.AddTable(100, 50, dblCols, dblRows);


                // Add text to the row 1 cell 1
                table[0, 0].TextFrame.Text = "Row 1 Cell 1";

                // Add text to the row 1 cell 2
                table[1, 0].TextFrame.Text = "Row 1 Cell 2";

                // Clone Row 1 at end of table
                table.Rows.AddClone(table.Rows[0], false);

                // Add text to the row 2 cell 1
                table[0, 1].TextFrame.Text = "Row 2 Cell 1";

                // Add text to the row 2 cell 2
                table[1, 1].TextFrame.Text = "Row 2 Cell 2";


                // Clone Row 2 as 4th row of table
                table.Rows.InsertClone(3,table.Rows[1], false);

                //Cloning first column at end
                table.Columns.AddClone(table.Columns[0], false);

                //Cloning 2nd column at 4th column index
                table.Columns.InsertClone(3,table.Columns[1], false);
                

                // Write PPTX to Disk
                presentation.Save(dataDir + "table_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
            }
            }
            }
            //ExEnd:CloningInTable
        }
   