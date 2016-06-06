//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

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
    public class TableWithCellBorders
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Tables();

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            //Instantiate Presentation class that represents PPTX file
            using (Presentation pres = new Presentation())
            {

                //Access first slide
                Slide sld = (Slide)pres.Slides[0];

                //Define columns with widths and rows with heights
                double[] dblCols = { 50, 50, 50, 50 };
                double[] dblRows = { 50, 30, 30, 30, 30 };

                //Add table shape to slide

                //Add table shape to slide
                ITable tbl = sld.Shapes.AddTable(100, 50, dblCols, dblRows);

                //Set border format for each cell
                foreach (IRow row in tbl.Rows)
                    foreach (ICell cell in row)
                    {
                        cell.BorderTop.FillFormat.FillType = FillType.NoFill;
                        cell.BorderBottom.FillFormat.FillType = FillType.NoFill;
                        cell.BorderLeft.FillFormat.FillType = FillType.NoFill;
                        cell.BorderRight.FillFormat.FillType = FillType.NoFill;
                    }



                //Write PPTX to Disk
                pres.Save(dataDir + "table.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
            }

        }
    }
}