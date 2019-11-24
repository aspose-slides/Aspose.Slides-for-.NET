using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Tables
{
    class SetFirstRowAsHeader
    {

        public static void Run() {

            //ExStart:SetFirstRowAsHeader

            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Tables();

            // Instantiate Presentation class that represents PPTX
            Presentation pres = new Presentation(dataDir + "table.pptx");

            // Access the first slide
            ISlide sld = pres.Slides[0];

            // Initialize null TableEx
            ITable tbl = null;

            // Iterate through the shapes and set a reference to the table found
            foreach (IShape shp in sld.Shapes)
            {
                if (shp is ITable) {
                tbl = (ITable)shp;
            }
        }

       
           //Set the first row of a table as header with a special formatting.
           tbl.FirstRow = true;
            

           
            //ExEnd:SetFirstRowAsHeader

        }
    }
}
