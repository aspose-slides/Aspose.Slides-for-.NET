using System.IO;

using Aspose.Slides;

namespace Aspose.Slides.Examples.CSharp.Text
{
    public class AddColumnInTexBoxes
    {
        public static void Run()
        {
            // ExStart:AddColumnInTexBoxes
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();
            using (Presentation presentation = new Presentation())
{
           // Get the first slide of presentation
            ISlide slide = presentation.Slides[0];

          // Add an AutoShape of Rectangle type
            IAutoShape aShape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 300, 300);

        // Add TextFrame to the Rectangle
           aShape.AddTextFrame("All these columns are limited to be within a single text container -- " +
        "you can add or delete text and the new or remaining text automatically adjusts " +
        "itself to flow within the container. You cannot have text flow from one container " +
        "to other though -- we told you PowerPoint's column options for text are limited!");

       // Get text format of TextFrame
        ITextFrameFormat format = aShape.TextFrame.TextFrameFormat;

      // Specify number of columns in TextFrame
        format.ColumnCount = 3;

     // Specify spacing between columns
        format.ColumnSpacing = 10;

    // Save created presentation
       presentation.Save("ColumnCount.pptx", SaveFormat.Pptx);

            }
         
            }
        // ExEnd:AddColumnInTexBoxes
    }
}