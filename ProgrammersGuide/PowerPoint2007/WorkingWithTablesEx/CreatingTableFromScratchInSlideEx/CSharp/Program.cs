//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;
using Aspose.Slides.Pptx;
using System.Drawing;

namespace CreatingTableFromScratchInSlideEx
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Instantiate PresentationEx class that represents PPTX file
            PresentationEx pres = new PresentationEx();

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            //Access first slide
            SlideEx sld = pres.Slides[0];

            //Define columns with widths and rows with heights
            double[] dblCols = { 50, 50, 50 };
            double[] dblRows = { 50, 30, 30, 30, 30 };

            //Add table shape to slide
            int idx = sld.Shapes.AddTable(100, 50, dblCols, dblRows);
            TableEx tbl = (TableEx)sld.Shapes[idx];

            //Set border format for each cell
            foreach (RowEx row in tbl.Rows)
                foreach (CellEx cell in row)
                {
                    cell.BorderTop.FillFormat.FillType = FillTypeEx.Solid;
                    cell.BorderTop.FillFormat.SolidFillColor.Color = Color.Red;
                    cell.BorderTop.Width = 5;

                    cell.BorderBottom.FillFormat.FillType = FillTypeEx.Solid;
                    cell.BorderBottom.FillFormat.SolidFillColor.Color = Color.Red;
                    cell.BorderBottom.Width = 5;

                    cell.BorderLeft.FillFormat.FillType = FillTypeEx.Solid;
                    cell.BorderLeft.FillFormat.SolidFillColor.Color = Color.Red;
                    cell.BorderLeft.Width = 5;

                    cell.BorderRight.FillFormat.FillType = FillTypeEx.Solid;
                    cell.BorderRight.FillFormat.SolidFillColor.Color = Color.Red;
                    cell.BorderRight.Width = 5;
                }

            //Merge cells 1 & 2 of row 1
            tbl.MergeCells(tbl[0, 0], tbl[1, 0], false);

            //Add text to the merged cell
            tbl[0, 0].TextFrame.Text = "Merged Cells";

            //Write PPTX to Disk
            pres.Write(dataDir + "table.pptx");

        }
    }
}