using System;
using System.Net;
using Aspose.Slides;
using Aspose.Slides.Export;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    class AddVideoFrameFromWebSource
    {
        //ExStart:AddVideoFrameFromWebSource
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Shapes();

            using (Presentation pres = new Presentation())
            {
                AddVideoFromYouTube(pres, "Tj75Arhq5ho");
                pres.Save(dataDir + "AddVideoFrameFromWebSource_out.pptx", SaveFormat.Pptx);
            }
        }

        private static void AddVideoFromYouTube(Presentation pres, string videoId)
        {
            //add videoFrame
            IVideoFrame videoFrame = pres.Slides[0].Shapes.AddVideoFrame(10, 10, 427, 240, "https://www.youtube.com/embed/" + videoId);
            videoFrame.PlayMode = VideoPlayModePreset.Auto;

            //load thumbnail
            using (WebClient client = new WebClient())
            {
                string thumbnailUri = "http://img.youtube.com/vi/" + videoId + "/hqdefault.jpg";
                videoFrame.PictureFormat.Picture.Image = pres.Images.AddImage(client.DownloadData(thumbnailUri));
            }
        }
        //ExEnd:AddVideoFrameFromWebSource
    }
}



