//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;
using System.Drawing;

namespace CreateTableFromScratch
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Instantiate a Presentation object that represents a PPT file
            Presentation pres = new Presentation(dataDir + "demo.ppt");


            //Accessing a slide using its slide position
            Slide slide = pres.GetSlideByPosition(2);


            //Setting table parameters
            int xPosition = 880;
            int yPosition = 1400;
            int tableWidth = 4000;
            int tableHeight = 500;
            int columns = 4;
            int rows = 4;
            double borderWidth = 2;


            //Adding a new table to the slide using specified table parameters
            Table table = slide.Shapes.AddTable(xPosition, yPosition, tableWidth, tableHeight,
                                           columns, rows, borderWidth, Color.Blue);


            //Setting the alternative text for the table that can help in future to find
            //and identify this specific table
            table.AlternativeText = "myTable";


            //Merging two cells
            table.MergeCells(table.GetCell(0, 0), table.GetCell(1, 0));


            //Accessing the text frame of the first cell of the first row in the table
            TextFrame tf = table.GetCell(0, 0).TextFrame;


            //If text frame is not null then add some text to the cell
            if (tf != null)
            {
                tf.Paragraphs[0].Portions[0].Text = "Welcome to Aspose.Slides ";
            }


            //Writing the presentation as a PPT file
            pres.Write(dataDir + "modified.ppt");

        }
    }
}