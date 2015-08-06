using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Slides
{
    class Program
    {
        static void Main(string[] args)
        {
            string FileName = @"E:\Aspose\Aspose Vs VSTO\Aspose.Slides Vs VSTO Presentations v 1.1\Sample Files\Removing Row Or Column in Table.pptx";
            Presentation MyPresentation = new Presentation(FileName);

            //Get First Slide
            ISlide sld = MyPresentation.Slides[0];

            foreach (IShape shp in sld.Shapes)
                if (shp is ITable)
                {
                    ITable tbl = (ITable)shp;
                    tbl.Rows.RemoveAt(0, false);
                }

            MyPresentation.Save(FileName,Export.SaveFormat.Pptx);
        }
    }
}
