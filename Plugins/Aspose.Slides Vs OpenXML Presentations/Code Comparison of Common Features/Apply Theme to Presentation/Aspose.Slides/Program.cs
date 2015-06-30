using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Slides
{
    class Program
    {
        static void Main(string[] args)
        {
            string presentationFile = @"E:\Aspose\Aspose Vs OpenXML\Aspose.Slides Vs OpenXML Presentation v1.1\Sample Files\My Sheet2.pptx";
            string themePresentation = @"E:\Aspose\Aspose Vs OpenXML\Aspose.Slides Vs OpenXML Presentation v1.1\Sample Files\My Theme.pptx";
            ApplyThemeToPresentation(themePresentation,presentationFile);
        }

        public static void ApplyThemeToPresentation(string presentationFile, string outputFile)
        {
            //Instantiate Presentation class to load the source presentation file
            Presentation srcPres = new Presentation(presentationFile);

            //Instantiate Presentation class for destination presentation (where slide is to be cloned)
            Presentation destPres = new Presentation(outputFile);

            //Instantiate ISlide from the collection of slides in source presentation along with
            //master slide
            ISlide SourceSlide = srcPres.Slides[0];

            //Clone the desired master slide from the source presentation to the collection of masters in the
            //destination presentation
            IMasterSlideCollection masters = destPres.Masters;
            IMasterSlide SourceMaster = SourceSlide.LayoutSlide.MasterSlide;

            //Clone the desired master slide from the source presentation to the collection of masters in the
            //destination presentation
            IMasterSlide iSlide = masters.AddClone(SourceMaster);

            //Clone the desired slide from the source presentation with the desired master to the end of the
            //collection of slides in the destination presentation
            ISlideCollection slds = destPres.Slides;
            slds.AddClone(SourceSlide, iSlide, true);

            //Clone the desired master slide from the source presentation to the collection of masters in the//destination presentation
            //Save the destination presentation to disk
            destPres.Save(outputFile, SaveFormat.Pptx);
        }
    }
}
