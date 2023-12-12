using Aspose.Slides.Export;
using Aspose.Slides;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

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
            string resultPath = Path.Combine(RunExamples.OutPath, "OutPresBitmap_Comments.png");

            Presentation pres = new Presentation(dataDir + "presentation.pptx");
            Bitmap bmp = new Bitmap(740, 960);

            IRenderingOptions renderOptions = new RenderingOptions();
            NotesCommentsLayoutingOptions notesOptions = new NotesCommentsLayoutingOptions();
            notesOptions.CommentsAreaColor = Color.Red;
            notesOptions.CommentsAreaWidth = 200;
            notesOptions.CommentsPosition = CommentsPositions.Right;
            notesOptions.NotesPosition = NotesPositions.BottomTruncated;
            renderOptions.SlidesLayoutOptions = notesOptions;

            using (Graphics graphics = Graphics.FromImage(bmp))
            {
                pres.Slides[0].RenderToGraphics(renderOptions, graphics);
            }

            bmp.Save(resultPath, ImageFormat.Png);
            System.Diagnostics.Process.Start(resultPath);

        }

    }

    //ExEnd:RenderComments
}
    

