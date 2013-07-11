//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;
using Aspose.Slides.Pptx;
using Aspose.Slides.Pptx.Animation;
using System.Drawing;

namespace ApplyingAnimationsOnShapesInsideSlideEx
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            //Instantiate PrseetationEx class that represents the PPTX
            PresentationEx pres = new PresentationEx();
            SlideEx sld = pres.Slides[0];

            //Now create effect "PathFootball" for existing shape from scratch.
            int idx = sld.Shapes.AddAutoShape(ShapeTypeEx.Rectangle, 150, 150, 250, 25);
            AutoShapeEx ashp = (AutoShapeEx)sld.Shapes[idx];
            ashp.AddTextFrame("Animated TextBox");

            //Add PathFootBall animation effect
            ShapeEx shape = pres.Slides[0].Shapes[idx];
            pres.Slides[0].Timeline.MainSequence.AddEffect(shape, EffectTypeEx.PathFootball,
                            EffectSubtypeEx.None, EffectTriggerTypeEx.AfterPrevious);

            //Create some kind of "button".
            int index = pres.Slides[0].Shapes.AddAutoShape(ShapeTypeEx.Bevel, 10, 10, 20, 20);
            ShapeEx shapeTrigger = pres.Slides[0].Shapes[index];

            //Create sequence of effects for this button.
            SequenceEx seqInter = pres.Slides[0].Timeline.InteractiveSequences.Add(shapeTrigger);

            //Create custom user path. Our object will be moved only after "button" click.
            EffectEx fxUserPath = seqInter.AddEffect(shape, EffectTypeEx.PathUser, EffectSubtypeEx.None, EffectTriggerTypeEx.OnClick);

            //Created path is empty so we should add commands for moving.
            MotionEffectEx motionBhv = ((MotionEffectEx)fxUserPath.Behaviors[0]);
            PointF[] pts = new PointF[1];
            pts[0] = new PointF(0.076f, 0.59f);
            motionBhv.Path.Add(MotionCommandPathTypeEx.LineTo, pts, MotionPathPointsTypeEx.Auto, true);
            pts[0] = new PointF(-0.076f, -0.59f);
            motionBhv.Path.Add(MotionCommandPathTypeEx.LineTo, pts, MotionPathPointsTypeEx.Auto, false);
            motionBhv.Path.Add(MotionCommandPathTypeEx.End, null, MotionPathPointsTypeEx.Auto, false);

            //Write the presentation as PPTX to disk
            pres.Write(dataDir + "AnimExample.pptx");


        }
    }
}