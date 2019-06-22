using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Text
{
    class AddColumnsinTextFrame
    {
        public static void Run()
        {

            //ExStart:AddColumnsinTextFrame
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();

            string outPptxFileName = dataDir + "ColumnsTest.pptx";
            using (Presentation pres = new Presentation())
            {
                IAutoShape shape1 = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 300, 300);
                TextFrameFormat format = (TextFrameFormat)shape1.TextFrame.TextFrameFormat;

                format.ColumnCount = 2;
                shape1.TextFrame.Text = "All these columns are limited to be within a single text container -- " +
                                          "you can add or delete text and the new or remaining text automatically adjusts " +
                                          "itself to flow within the container. You cannot have text flow from one container " +
                                          "to other though -- we told you PowerPoint's column options for text are limited!";
                pres.Save(outPptxFileName, SaveFormat.Pptx);

                using (Presentation test = new Presentation(outPptxFileName))
                {
                    Assert.AreEqual(2, ((AutoShape)test.Slides[0].Shapes[0]).TextFrame.TextFrameFormat.ColumnCount);
                    Assert.AreEqual(double.NaN, ((AutoShape)test.Slides[0].Shapes[0]).TextFrame.TextFrameFormat.ColumnSpacing);
                }

                format.ColumnSpacing = 20;
                pres.Save(outPptxFileName, SaveFormat.Pptx);

                using (Presentation test = new Presentation(outPptxFileName))
                {
                    Assert.AreEqual(2, ((AutoShape)test.Slides[0].Shapes[0]).TextFrame.TextFrameFormat.ColumnCount);
                    Assert.AreEqual(20, ((AutoShape)test.Slides[0].Shapes[0]).TextFrame.TextFrameFormat.ColumnSpacing);
                }

                format.ColumnCount = 3;
                format.ColumnSpacing = 15;
                pres.Save(outPptxFileName, SaveFormat.Pptx);

                using (Presentation test = new Presentation(outPptxFileName))
                {
                    Assert.AreEqual(3, ((AutoShape)test.Slides[0].Shapes[0]).TextFrame.TextFrameFormat.ColumnCount);
                    Assert.AreEqual(15, ((AutoShape)test.Slides[0].Shapes[0]).TextFrame.TextFrameFormat.ColumnSpacing);
                }

            }

            //ExEnd:AddColumnsinTextFrame
        }
    }
}
