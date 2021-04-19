
using System.IO;
using Aspose.Cells;
using Aspose.Slides;
using Aspose.Slides.DOM.Ole;
using SaveFormat = Aspose.Slides.Export.SaveFormat;

namespace Aspose.Slides.Examples.CSharp.Shapes 
{
    public class ChangeOLEObjectData
    {
        public static void Run()
        {
            //ExStart:ChangeOLEObjectData
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Shapes();

            using (Presentation pres = new Presentation(dataDir + "ChangeOLEObjectData.pptx"))
            {
                ISlide slide = pres.Slides[0];

                OleObjectFrame ole = null;

                // Traversing all shapes for Ole frame
                foreach (IShape shape in slide.Shapes)
                {
                    if (shape is OleObjectFrame)
                    {
                        ole = (OleObjectFrame) shape;
                    }
                }

                if (ole != null)
                {
                    using (MemoryStream msln = new MemoryStream(ole.EmbeddedData.EmbeddedFileData))
                    {
                        // Reading object data in Workbook
                        Workbook Wb = new Workbook(msln);

                        using (MemoryStream msout = new MemoryStream())
                        {
                            // Modifying the workbook data
                            Wb.Worksheets[0].Cells[0, 4].PutValue("E");
                            Wb.Worksheets[0].Cells[1, 4].PutValue(12);
                            Wb.Worksheets[0].Cells[2, 4].PutValue(14);
                            Wb.Worksheets[0].Cells[3, 4].PutValue(15);

                            OoxmlSaveOptions so1 = new OoxmlSaveOptions(Aspose.Cells.SaveFormat.Xlsx);
                            Wb.Save(msout, so1);

                            // Changing Ole frame object data
                            IOleEmbeddedDataInfo newData = new OleEmbeddedDataInfo(msout.ToArray(), ole.EmbeddedData.EmbeddedFileExtension);
                            ole.SetEmbeddedData(newData);
                        }
                    }
                }
                pres.Save(dataDir + "OleEdit_out.pptx", SaveFormat.Pptx);
            }

            //ExEnd:ChangeOLEObjectData
        }
    }
}