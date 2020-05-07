using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using Aspose.Slides.Util;

namespace CSharp.Presentations.Properties
{
    class CheckPresentationProtection
    {
        public static void Run()
        {
            //Path for source presentation
            string pptxFile = Path.Combine(RunExamples.GetDataDir_PresentationProperties(), "modify_pass2.pptx");
            string pptFile = Path.Combine(RunExamples.GetDataDir_PresentationProperties(), "open_pass1.ppt");

            // Check the Write Protection Password via IPresentationInfo Interface
            IPresentationInfo presentationInfo = PresentationFactory.Instance.GetPresentationInfo(pptxFile);
            bool isWriteProtectedByPassword = presentationInfo.IsWriteProtected == NullableBool.True && presentationInfo.CheckWriteProtection("pass2");
            Console.WriteLine("Is presentation write protected by password = " + isWriteProtectedByPassword);

            // Check the Write Protection Password via IProtectionManager Interface
            using (var presentation = new Presentation(pptxFile))
            {
                bool isWriteProtected = presentation.ProtectionManager.CheckWriteProtection("pass2");
                Console.WriteLine("Is presentation write protected = " + isWriteProtected);
            }

            // Check Presentation Open Protection via IPresentationInfo Interface
            presentationInfo = PresentationFactory.Instance.GetPresentationInfo(pptFile);
            if (presentationInfo.IsPasswordProtected)
            {
                Console.WriteLine("The presentation '" + pptxFile + "' is protected by password to open.");
            }
        }
    }
}
