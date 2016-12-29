using System.IO;
using Aspose.Slides;

namespace Aspose.Slides.Examples.CSharp.Tables
{
    public class UpdateExistingTable
    {
        public static void Run()
        {
            // ExStart:UpdateExistingTable
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Tables();

            // Instantiate Presentation class that represents PPTX// Instantiate Presentation class that represents PPTX
            using (Presentation pres = new Presentation(dataDir + "UpdateExistingTable.pptx"))
            {

                // Access the first slide
                ISlide sld = pres.Slides[0];

                // Initialize null TableEx
                ITable tbl = null;

                // Iterate through the shapes and set a reference to the table found
                foreach (IShape shp in sld.Shapes)
                    if (shp is ITable)
                        tbl = (ITable)shp;

                // Set the text of the first column of second row
                tbl[0, 1].TextFrame.Text = "New";

                //Write the PPTX to Disk
                pres.Save(dataDir + "table1_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
            }
            // ExEnd:UpdateExistingTable
        }
    }
}