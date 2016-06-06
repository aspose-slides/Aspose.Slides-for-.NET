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
    public class RemovingRowColumn
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Tables();

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