using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;

/*
Aspose.Slides can play presentations and generate a set of frames for an entire animation with a specific frame per second (FPS). 
Those frames can then be used to create a video through tools like FFmpeg.

This C# code demonstrates a presentation to video export operation with frames set at 30FPS 
*/

namespace CSharp.Presentations.Conversion
{
    class ConvertToAnimation
    {
        public static void Run()
        {
            string presentationName = Path.Combine(RunExamples.GetDataDir_Conversion(), "SimpleAnimations.pptx");
            const int FPS = 30;

            using (Presentation presentation = new Presentation(presentationName))
            {
                using (var animationsGenerator = new PresentationAnimationsGenerator(presentation))
                using (var player = new PresentationPlayer(animationsGenerator, FPS))
                {
                    player.FrameTick += (sender, args) =>
                    {
                        args.GetFrame().Save(Path.Combine(RunExamples.OutPath, $"frame_{sender.FrameIndex}.png"));
                    };

                    animationsGenerator.Run(presentation.Slides);
                }
            }
        }
    }
}
