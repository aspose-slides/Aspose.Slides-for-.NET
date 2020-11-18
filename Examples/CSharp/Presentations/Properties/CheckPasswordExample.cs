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
    public class CheckPasswordExample
    {
        // The example below demonstrates how to check a password to open a presentation

        public static void Run()
        {
            //Path for source presentation
            string pptFile = Path.Combine(RunExamples.GetDataDir_PresentationProperties(), "open_pass1.ppt");

            // Check the Password via IPresentationInfo Interface
            IPresentationInfo presentationInfo = PresentationFactory.Instance.GetPresentationInfo(pptFile);
            bool isPasswordCorrect = presentationInfo.CheckPassword("my_password");
            Console.WriteLine("The password \"my_password\" for the presentation is " + isPasswordCorrect);
            
            isPasswordCorrect = presentationInfo.CheckPassword("pass1");
            Console.WriteLine("The password \"pass1\" for the presentation is " + isPasswordCorrect);
        }
    }
}
