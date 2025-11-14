using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

/*
Thise example demonstrates how to obtain information about fonts 
that will be substituted during the rendering of the specified slides.
*/

namespace CSharp.Text
{
    class GetFontsSlideSubstitution
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();

            using (Presentation pres = new Presentation(dataDir + "PresFontsSubst.pptx"))
            {
                foreach (var fontSubstitution in pres.FontsManager.GetSubstitutions(new int[] {1, 2}))
                {
                    Console.WriteLine("{0} -> {1}", fontSubstitution.OriginalFontName, fontSubstitution.SubstitutedFontName);
                }
            }
        }
    }
}