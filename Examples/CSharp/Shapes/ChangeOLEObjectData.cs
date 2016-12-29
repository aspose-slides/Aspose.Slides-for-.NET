
using System.IO;

using Aspose.Slides;
using Aspose.Slides.Export;

namespace Aspose.Slides.Examples.CSharp.Shapes 
{
    public class ChangeOLEObjectData
    {
        public static void Run()
        {
            //ExStart:ChangeOLEObjectData
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Shapes();

            Presentation pres = new Presentation(dataDir+ "ChangeOLEObjectData.pptx");

            ISlide slide = pres.Slides[0];

            OleObjectFrame ole = null;

            // Traversing all shapes for Ole frame
            foreach (IShape shape in slide.Shapes)
            {
                if (shape is OleObjectFrame)
                {
                    ole = (OleObjectFrame)shape;
                }
            }

            if (ole != null)
            {
                // Reading object data in Workbook
                Aspose.Cells.Workbook Wb;

                using (System.IO.MemoryStream msln = new System.IO.MemoryStream(ole.ObjectData))
                {
                    Wb = new Aspose.Cells.Workbook(msln);

                    using (System.IO.MemoryStream msout = new System.IO.MemoryStream())
                    {
                        // Modifying the workbook data
                        Wb.Worksheets[0].Cells[0, 4].PutValue("E");
                        Wb.Worksheets[0].Cells[1, 4].PutValue(12);
                        Wb.Worksheets[0].Cells[2, 4].PutValue(14);
                        Wb.Worksheets[0].Cells[3, 4].PutValue(15);

                        Aspose.Cells.OoxmlSaveOptions so1 = new Aspose.Cells.OoxmlSaveOptions(Aspose.Cells.SaveFormat.Xlsx);

                        Wb.Save(msout, so1);

                        // Changing Ole frame object data
                        msout.Position = 0;
                        ole.ObjectData = msout.ToArray();
                    }
                }
            }

            pres.Save(dataDir + "OleEdit_out.pptx", SaveFormat.Pptx);
            //ExEnd:ChangeOLEObjectData
        }
    }
}