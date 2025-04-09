﻿using System;
using System.IO;
using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Slides.Media
{
    class ExtractVideo
    {
        public static void Run()
        {
            
            //ExStart:ExtractVideo
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations_Media();
            string outMedia = Path.Combine(RunExamples.OutPath, "NewVideo_out.");

            // Instantiate a Presentation object that represents a presentation file 
            Presentation presentation = new Presentation(dataDir + "Video.pptx");

            foreach (ISlide slide in presentation.Slides)
            {
                foreach (IShape shape in presentation.Slides[0].Shapes)
                {
                    if (shape is VideoFrame)
                    {
                        IVideoFrame vf = shape as IVideoFrame;
                        String type = vf.EmbeddedVideo.ContentType;
                        int ss = type.LastIndexOf('/');
                        type = type.Remove(0, type.LastIndexOf('/') + 1);
                        Byte[] buffer = vf.EmbeddedVideo.BinaryData;
                        using (FileStream stream = new FileStream(outMedia + type, FileMode.Create, FileAccess.Write, FileShare.Read))
                        {                                                     
                            stream.Write(buffer, 0, buffer.Length);
                        }
                    }
                }
            }
            //ExEnd:ExtractVideo
        }
    }
}