using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

/*
Thise example demonstrates you how the GetSubstitutions method is used to get all fonts 
that will be substituted when a presentation is rendered.
*/
namespace CSharp.Text
{
    class GetFontsSubstitution
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();

            using (Presentation pres = new Presentation(dataDir + "PresFontsSubst.pptx"))
            {
                foreach (var fontSubstitution in pres.FontsManager.GetSubstitutions())
                {
                    Console.WriteLine("{0} -> {1}", fontSubstitution.OriginalFontName, fontSubstitution.SubstitutedFontName);
                }
            }
        }
    }
}
