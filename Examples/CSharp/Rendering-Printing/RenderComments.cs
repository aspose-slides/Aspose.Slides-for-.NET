using Aspose.Slides.Export;
using Aspose.Slides;
using System.Drawing;
using System.Drawing.Imaging;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Rendering.Printing
{
    class RenderComments
    {
        public static void Run()
        {
            //ExStart:RenderComments
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Rendering();
            Presentation pres = new Presentation(dataDir + "presentation.pptx");
            Bitmap bmp = new Bitmap(740, 960);

            NotesCommentsLayoutingOptions opts = new NotesCommentsLayoutingOptions();
            opts.CommentsAreaColor = Color.Red;

            opts.CommentsAreaWidth = 200;
            opts.CommentsPosition = CommentsPositions.Right;
            opts.NotesPosition = NotesPositions.BottomTruncated;

            using (Graphics graphics = Graphics.FromImage(bmp))
            {
                pres.Slides[0].RenderToGraphics(opts, graphics);
            }

            bmp.Save(dataDir + "OutPresBitmap.png", ImageFormat.Png);
            System.Diagnostics.Process.Start(dataDir + "OutPresBitmap.png");

        }

    }

    //ExEnd:RenderComments
}
    

