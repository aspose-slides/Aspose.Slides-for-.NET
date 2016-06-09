using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Slides;

namespace CSharp.ProgrammersGuide.Presentations
{
    class GetFileFormat
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Presentations();

            IPresentationInfo info = PresentationFactory.Instance.GetPresentationInfo(dataDir + "HelloWorld.pptx");

            switch (info.LoadFormat)
            {
                case LoadFormat.Pptx:
                {
                 break;
                }

                case LoadFormat.Unknown:
                {
                 break;
                }
            }
        }
    }

}
