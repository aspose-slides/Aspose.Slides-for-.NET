using System.IO;

using Aspose.Slides;

namespace Aspose.Slides.Examples.CSharp.Text
{
    public class GetTextFromSmartArtNode
    {
        public static void Run()
        {
            // ExStart:GetTextFromSmartArtNode
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();

          using (Presentation pres = new Presentation("Presentation.pptx"))
{
            ISlide slide = presentation.Slides[0];
            ISmartArt smartArt = (ISmartArt)slide.Shapes[0];

            ISmartArtNodeCollection smartArtNodes = smartArt.AllNodes;
            foreach (ISmartArtNode smartArtNode in smartArtNodes)
            {
                foreach (ISmartArtShape nodeShape in smartArtNode.Shapes)
                {
                    if (nodeShape.TextFrame != null)
                        Console.WriteLine(nodeShape.TextFrame.Text);
                }
            }
            }
            }
        // ExEnd:GetTextFromSmartArtNode
        }
    }
