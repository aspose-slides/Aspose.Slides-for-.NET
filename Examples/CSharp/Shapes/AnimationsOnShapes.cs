using System.IO;

using Aspose.Slides;
using Aspose.Slides.Export;
using Aspose.Slides.Animation;
using System.Drawing;

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    public class AnimationsOnShapes
    {
        public static void Run()
        {
            //ExStart:AnimationsOnShapes
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Shapes();

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            // Instantiate PrseetationEx class that represents the PPTX
            using (Presentation pres = new Presentation())
            {
                ISlide sld = pres.Slides[0];

                // Now create effect "PathFootball" for existing shape from scratch.
                IAutoShape ashp = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 150, 150, 250, 25);

                ashp.AddTextFrame("Animated TextBox");

                // Add PathFootBall animation effect
                pres.Slides[0].Timeline.MainSequence.AddEffect(ashp, EffectType.PathFootball,
                                       EffectSubtype.None, EffectTriggerType.AfterPrevious);

                // Create some kind of "button".
                IShape shapeTrigger = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Bevel, 10, 10, 20, 20);

                // Create sequence of effects for this button.
                ISequence seqInter = pres.Slides[0].Timeline.InteractiveSequences.Add(shapeTrigger);

                // Create custom user path. Our object will be moved only after "button" click.
                IEffect fxUserPath = seqInter.AddEffect(ashp, EffectType.PathUser, EffectSubtype.None, EffectTriggerType.OnClick);

                // Created path is empty so we should add commands for moving.
                IMotionEffect motionBhv = ((IMotionEffect)fxUserPath.Behaviors[0]);

                PointF[] pts = new PointF[1];
                pts[0] = new PointF(0.076f, 0.59f);
                motionBhv.Path.Add(MotionCommandPathType.LineTo, pts, MotionPathPointsType.Auto, true);
                pts[0] = new PointF(-0.076f, -0.59f);
                motionBhv.Path.Add(MotionCommandPathType.LineTo, pts, MotionPathPointsType.Auto, false);
                motionBhv.Path.Add(MotionCommandPathType.End, null, MotionPathPointsType.Auto, false);

                //Write the presentation as PPTX to disk
                pres.Save(dataDir + "AnimExample_out.pptx", SaveFormat.Pptx);
            }
            //ExEnd:AnimationsOnShapes
        }
    }
}