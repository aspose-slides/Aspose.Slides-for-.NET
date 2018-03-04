using Aspose.Slides.Export;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Presentations.Conversion
{
    class AddBlobToPresentation
    {



        public static void Run()
        {
            //ExStart:AddBlobToPresentation
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Conversion();
            const string pathToVeryLargeVideo = "veryLargeVideo.avi";

            // create a new presentation which will contain this video
            using (Presentation pres = new Presentation())
            {
                using (FileStream fileStream = new FileStream(pathToVeryLargeVideo, FileMode.Open))
                {
                    // let's add the video to the presentation - we choose KeepLocked behavior, because we not
                    // have an intent to access the "veryLargeVideo.avi" file.
                    IVideo video = pres.Videos.AddVideo(fileStream, LoadingStreamBehavior.KeepLocked);
                    pres.Slides[0].Shapes.AddVideoFrame(0, 0, 480, 270, video);

                    // save the presentation. Despite that the output presentation will be very large, the memory
                    // consumption will be low the whole lifetime of the pres object
                    pres.Save("presentationWithLargeVideo.pptx", SaveFormat.Pptx);
                }
            }
        }
        //ExEnd:AddBlobToPresentation
    }
  }

