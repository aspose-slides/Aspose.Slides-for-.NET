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
    public class VerticallyAlignText
    {
        public static void Run()
        {
            // ExStart:IdentifyingTheMergedCellsinTable
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Tables();


            using (Presentation pres = new Presentation(dataDir + "SomePresentationWithTable.pptx"))
            {
                ITable table = pres.Slides[0].Shapes[0] as ITable; // assuming that Slide#0.Shape#0 is a table
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    for (int j = 0; j < table.Columns.Count; j++)
                    {
                        ICell currentCell = table.Rows[i][j];
                        if (currentCell.IsMergedCell)
                        {
                            Console.WriteLine(string.Format("Cell {0};{1} is a part of merged cell with RowSpan={2} and ColSpan={3} starting from Cell {4};{5}.",
                                              i, j, currentCell.RowSpan, currentCell.ColSpan, currentCell.FirstRowIndex, currentCell.FirstColumnIndex));


                        }
                    }

                }
            }


            // ExEnd:IdentifyingTheMergedCellsinTable
         }
    }
}

