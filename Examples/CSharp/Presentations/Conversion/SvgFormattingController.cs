using System.IO;
using Aspose.Slides.Export;

namespace Aspose.Slides.Examples.CSharp.Presentations.Conversion
{

    // Code below shows how to use ISvgShapeAndTextFormattingController interface for
    // tspan Id attribute manipulation.

    public class SvgFormattingController
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_Conversion();
            string pptxFileName = Path.Combine(dataDir, "Convert_Svg_Custom.pptx");
            string outSvgFileName = Path.Combine(RunExamples.OutPath, "Convert_Svg_Custom.svg");

            using (Presentation pres = new Presentation(pptxFileName))
            {
                using (FileStream stream = new FileStream(outSvgFileName, FileMode.Create))
                {
                    SVGOptions svgOptions = new SVGOptions
                    {
                        ShapeFormattingController = new MySvgShapeFormattingController()
                    };

                    pres.Slides[0].WriteAsSvg(stream, svgOptions);
                }
            }
        }
    }

    class MySvgShapeFormattingController : ISvgShapeAndTextFormattingController
    {
        private int m_shapeIndex, m_portionIndex, m_tspanIndex;

        public MySvgShapeFormattingController(int shapeStartIndex = 0)
        {
            m_shapeIndex = shapeStartIndex;
            m_portionIndex = 0;
        }

        public void FormatShape(Aspose.Slides.Export.ISvgShape svgShape, IShape shape)
        {
            svgShape.Id = string.Format("shape-{0}", m_shapeIndex++);
            m_portionIndex = m_tspanIndex = 0;
        }

        public void FormatText(Aspose.Slides.Export.ISvgTSpan svgTSpan, IPortion portion, ITextFrame textFrame)
        {
            int paragraphIndex = 0; int portionIndex = 0;
            for (int i = 0; i < textFrame.Paragraphs.Count; i = i + 1)
            {
                portionIndex = textFrame.Paragraphs[i].Portions.IndexOf(portion);
                if (portionIndex > -1) { paragraphIndex = i; break; }
            }
            if (m_portionIndex != portionIndex)
            {
                m_tspanIndex = 0;
                m_portionIndex = portionIndex;
            }
            svgTSpan.Id = string.Format("paragraph-{0}_portion-{1}_{2}", paragraphIndex, m_portionIndex, m_tspanIndex++);
        }

        public ISvgShapeFormattingController AsISvgShapeFormattingController
        {
            get { return this; }
        }
    }
}