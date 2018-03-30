using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Aspose.Slides;
using Aspose.Cells;
using Aspose.Slides.Examples;
using Aspose.Slides.Examples.CSharp;
using Aspose.Cells.Rendering;
namespace CSharp.Shapes
{
   public  class ImageAsEMF
    {
   public static void Run()
   {
       string dataDir = RunExamples.GetDataDir_Shapes();
      //ExStart:ImageAsEMF
    Workbook book = new Workbook(dataDir + "chart.xlsx");
    Worksheet sheet = book.Worksheets[0];
    Aspose.Cells.Rendering.ImageOrPrintOptions options = new Aspose.Cells.Rendering.ImageOrPrintOptions();
    options.HorizontalResolution = 200;
    options.VerticalResolution = 200;
    options.ImageFormat = System.Drawing.Imaging.ImageFormat.Emf;

    //Save the workbook to stream
    SheetRender sr = new SheetRender(sheet, options);
    Presentation pres = new Presentation();
    pres.Slides.RemoveAt(0);

    String EmfSheetName="";
    for (int j = 0; j < sr.PageCount; j++)
    {

        EmfSheetName=dataDir + "test" + sheet.Name + " Page" + (j + 1) + ".out.emf";
        sr.ToImage(j, EmfSheetName);
     
        var bytes = File.ReadAllBytes(EmfSheetName);
        var emfImage = pres.Images.AddImage(bytes);
        ISlide slide= pres.Slides.AddEmptySlide(pres.LayoutSlides.GetByType(SlideLayoutType.Blank));
        var m = slide.Shapes.AddPictureFrame(ShapeType.Rectangle, 0, 0, pres.SlideSize.Size.Width, pres.SlideSize.Size.Height, emfImage);
    }
    
    pres.Save(dataDir+"Saved.pptx", Aspose.Slides.Export.SaveFormat.Pptx);

       //ExEnd:ImageAsEMF
   }
   }
}
