using System.Drawing;
using Aspose.Slides.Export;
using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace Aspose.Slides.Examples.CSharp.Tables
{
    public class AddImageinsideTableCells                                                                            
    {
        public static void Run()
        {
            // ExStart:AddImageinsideTableCell
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Tables();

            // Instantiate Presentation class object
            Presentation presentation = new Presentation();

            // Access first slide
            ISlide islide = presentation.Slides[0];

            // Define columns with widths and rows with heights
            double[] dblCols = { 150, 150, 150, 150 };
            double[] dblRows = { 100, 100, 100, 100, 90 };

            // Add table shape to slide
            ITable tbl = islide.Shapes.AddTable(50, 50, dblCols, dblRows);

            // Creating a Bitmap Image object to hold the image file
            IImage image = Images.FromFile(dataDir + "aspose-logo.jpg");

            // Create an IPPImage object using the bitmap object
            IPPImage imgx1 = presentation.Images.AddImage(image);

            // Add image to first table cell
            tbl[0, 0].CellFormat.FillFormat.FillType = FillType.Picture;
            tbl[0, 0].CellFormat.FillFormat.PictureFillFormat.PictureFillMode = PictureFillMode.Stretch;
            tbl[0, 0].CellFormat.FillFormat.PictureFillFormat.Picture.Image = imgx1;
            tbl[0, 0].CellFormat.FillFormat.PictureFillFormat.CropRight = 20;
            tbl[0, 0].CellFormat.FillFormat.PictureFillFormat.CropLeft = 20;
            tbl[0, 0].CellFormat.FillFormat.PictureFillFormat.CropTop = 20;
            tbl[0, 0].CellFormat.FillFormat.PictureFillFormat.CropBottom = 20;
            // Save PPTX to Disk
            presentation.Save(dataDir + "Image_In_TableCell_out.pptx", SaveFormat.Pptx);
            // ExEnd:AddImageinsideTableCell
        }
    }
}