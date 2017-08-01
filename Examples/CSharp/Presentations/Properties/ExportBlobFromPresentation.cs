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
    class ExportBlobFromPresentation
    {
        public static void Run()
        {
            //ExStart:ExportBlobFromPresentation
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Conversion();
         const string hugePresentationWithAudiosAndVideosFile = @"c:\bin\aspose\Tasks\020, 38595\orig\Large  Video File Test1.pptx";

         LoadOptions loadOptions = new LoadOptions
    {
        BlobManagementOptions =
        {
            // lock the source file and don't load it into memory
            PresentationLockingBehavior = PresentationLockingBehavior.KeepLocked,
        }
      };

    // create the Presentation's instance, lock the "hugePresentationWithAudiosAndVideos.pptx" file.
    using (Presentation pres = new Presentation(hugePresentationWithAudiosAndVideosFile, loadOptions))
    {
        // let's save each video to a file. to prevent memory usage we need a buffer which will be used
        // to exchange tha data from the presentation's video stream to a stream for newly created video file.
        byte[] buffer = new byte[8 * 1024];

        // iterate through the videos
        for (var index = 0; index < pres.Videos.Count; index++)
        {
            IVideo video = pres.Videos[index];

            // open the presentation video stream. Please note that we intentionally avoid accessing properties
            // like video.BinaryData - this property returns a byte array containing full video, and that means
            // this bytes will be loaded into memory. We will use video.GetStream, which will return Stream and
            // that allows us to not load the whole video into memory.
            using (Stream presVideoStream = video.GetStream())
            {
                using (FileStream outputFileStream = File.OpenWrite($"video{index}.avi"))
                {
                    int bytesRead;
                    while ((bytesRead = presVideoStream.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        outputFileStream.Write(buffer, 0, bytesRead);
                    }
                 }
              }

            // memory consumption will stay low no matter what size the videos or presentation is.
        }

        // do the same for audios if needed.
     }
    
     }
        //ExEnd:ExportBlobFromPresentation
    }
  }

