using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Tables
{
    class CreateATableFromScratchInASlide
    {
        public static void Run() {

            //ExStart:CreateATableFromScratchInASlide

            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Tables();

            // Instantiate Presentation class that represents PPTX file
            Presentation pres = new Presentation();

            // Access first slide
            ISlide sld = pres.Slides[0];

            // Define columns with widths and rows with heights
            double[] dblCols = { 50, 50, 50 };
            double[] dblRows = { 50, 30, 30, 30, 30 };

            // Add table shape to slide
            ITable tbl = sld.Shapes.AddTable(100, 50, dblCols, dblRows);

            // Set border format for each cell
            for (int row = 0; row < tbl.Rows.Count; row++)
            {
                for (int cell = 0; cell < tbl.Rows[row].Count; cell++)
                {
                    tbl.Rows[row][cell].CellFormat.BorderTop.FillFormat.FillType = FillType.Solid;
                    tbl.Rows[row][cell].CellFormat.BorderTop.FillFormat.SolidFillColor.Color = Color.Red;
                    tbl.Rows[row][cell].CellFormat.BorderTop.Width = 5;

                    tbl.Rows[row][cell].CellFormat.BorderBottom.FillFormat.FillType = (FillType.Solid);
                    tbl.Rows[row][cell].CellFormat.BorderBottom.FillFormat.SolidFillColor.Color= Color.Red;
                    tbl.Rows[row][cell].CellFormat.BorderBottom.Width =5;

                    tbl.Rows[row][cell].CellFormat.BorderLeft.FillFormat.FillType = FillType.Solid;
                    tbl.Rows[row][cell].CellFormat.BorderLeft.FillFormat.SolidFillColor.Color =Color.Red;
                    tbl.Rows[row][cell].CellFormat.BorderLeft.Width = 5;

                    tbl.Rows[row][cell].CellFormat.BorderRight.FillFormat.FillType = FillType.Solid;
                    tbl.Rows[row][cell].CellFormat.BorderRight.FillFormat.SolidFillColor.Color = Color.Red;
                    tbl.Rows[row][cell].CellFormat.BorderRight.Width = 5;
                }
            }
            // Merge cells 1 & 2 of row 1
            tbl.MergeCells(tbl.Rows[0][0], tbl.Rows[1][1], false);

            // Add text to the merged cell
            tbl.Rows[0][0].TextFrame.Text = "Merged Cells";

            // Save PPTX to Disk
            pres.Save(dataDir + "table.pptx", SaveFormat.Pptx);
            //ExEnd:CreateATableFromScratchInASlide


        }
    }
}
