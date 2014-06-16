//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;

namespace RemovingRowColumn
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            Presentation pres = new Presentation();

            ISlide slide = pres.Slides[0];
            double[] colWidth = { 100, 50, 30 };
            double[] rowHeight = { 30, 50, 30 };

            ITable table = slide.Shapes.AddTable(100, 100, colWidth, rowHeight);

            table.Rows.RemoveAt(1, false);
            table.Columns.RemoveAt(1, false);


            pres.Save(dataDir + "TestTable.pptx", Aspose.Slides.Export.SaveFormat.Pptx);


        }
    }
}