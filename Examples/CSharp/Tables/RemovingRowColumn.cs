using System.IO;
using Aspose.Slides;

namespace Aspose.Slides.Examples.CSharp.Tables
{
    public class RemovingRowColumn
    {
        public static void Run()
        {
            // ExStart:RemovingRowColumn
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
            pres.Save(dataDir + "TestTable_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
            // ExEnd:RemovingRowColumn
        }
    }
}