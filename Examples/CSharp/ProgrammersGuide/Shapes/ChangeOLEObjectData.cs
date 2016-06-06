//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;
using Aspose.Slides.Export;
using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/


namespace CSharp.Shapes 
{
    public class ChangeOLEObjectData
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Shapes();

            Presentation pres = new Presentation(dataDir+ "ChangeOLEObjectData.pptx");

            ISlide slide = pres.Slides[0];

            OleObjectFrame ole = null;

            //Traversing all shapes for Ole frame
            foreach (IShape shape in slide.Shapes)
            {
                if (shape is OleObjectFrame)
                {
                    ole = (OleObjectFrame)shape;
                }
            }

            if (ole != null)
            {
                //Reading object data in Workbook
                Aspose.Cells.Workbook Wb;

                using (System.IO.MemoryStream msln = new System.IO.MemoryStream(ole.ObjectData))
                {
                    Wb = new Aspose.Cells.Workbook(msln);

                    using (System.IO.MemoryStream msout = new System.IO.MemoryStream())
                    {
                        //Modifying the workbook data
                        Wb.Worksheets[0].Cells[0, 4].PutValue("E");
                        Wb.Worksheets[0].Cells[1, 4].PutValue(12);
                        Wb.Worksheets[0].Cells[2, 4].PutValue(14);
                        Wb.Worksheets[0].Cells[3, 4].PutValue(15);

                        Aspose.Cells.OoxmlSaveOptions so1 = new Aspose.Cells.OoxmlSaveOptions(Aspose.Cells.SaveFormat.Xlsx);

                        Wb.Save(msout, so1);

                        //Changing Ole frame object data
                        msout.Position = 0;
                        ole.ObjectData = msout.ToArray();
                    }
                }
            }

            pres.Save(dataDir+ "OleEdit.pptx", SaveFormat.Pptx);

 
            
        }
    }
}