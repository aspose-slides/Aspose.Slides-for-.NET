using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides;
using Aspose.Slides.Export;
using Aspose.Slides.Examples.CSharp;

/*
This example demonstrates how to get and change the properties of a presentation document.
*/

namespace CSharp.Presentations.Properties
{
    class ExtendedPresentationProperties
    {
        public static void Run()
        {
            //Path for source presentation
            string pptxFile = Path.Combine(RunExamples.GetDataDir_PresentationProperties(), "ExtendDocumentProperies.pptx");
            //Out path
            string resultPath = Path.Combine(RunExamples.OutPath, "ExtendDocumentProperies-out1.pptx");

            using (var presentation = new Presentation(pptxFile))
            {
                IDocumentProperties documentProperties = presentation.DocumentProperties;

                // Print the read-only properties
                Console.WriteLine("Slides: " + documentProperties.Slides);
                Console.WriteLine("HiddenSlides: " + documentProperties.HiddenSlides);
                Console.WriteLine("Notes: " + documentProperties.Notes);
                Console.WriteLine("Paragraphs: " + documentProperties.Paragraphs);
                Console.WriteLine("MultimediaClips: " + documentProperties.MultimediaClips);
                Console.WriteLine("TitlesOfParts: " + string.Join("; ", documentProperties.TitlesOfParts));
                Console.WriteLine("HeadingPairs: ");
                IHeadingPair[] headingPairs = documentProperties.HeadingPairs;
                if (headingPairs.Length > 0)
                {
                    foreach (var headingPair in headingPairs)
                        Console.WriteLine(headingPair.Name + " " + headingPair.Count);
                }

                // Change several boolean properties
                documentProperties.ScaleCrop = true;
                documentProperties.LinksUpToDate = true;

                // Save the presentation with changed properties
                presentation.Save(resultPath, SaveFormat.Pptx);

                //Use the IPresentationInfo interface to read and change the document properties
                Console.WriteLine("\nProperties obtained by IPresentationInfo:\n");

                IPresentationInfo documentInfo = PresentationFactory.Instance.GetPresentationInfo(resultPath);
                documentProperties = documentInfo.ReadDocumentProperties();

                // Print the read-only properties
                Console.WriteLine("Slides: " + documentProperties.Slides);
                Console.WriteLine("HiddenSlides: " + documentProperties.HiddenSlides);
                Console.WriteLine("Notes: " + documentProperties.Notes);
                Console.WriteLine("Paragraphs: " + documentProperties.Paragraphs);
                Console.WriteLine("MultimediaClips: " + documentProperties.MultimediaClips);
                Console.WriteLine("TitlesOfParts: " + string.Join("; ", documentProperties.TitlesOfParts));
                Console.WriteLine("HeadingPairs: ");
                headingPairs = documentProperties.HeadingPairs;
                if (headingPairs.Length > 0)
                {
                    foreach (var headingPair in headingPairs)
                        Console.WriteLine(headingPair.Name + " " + headingPair.Count);
                }

                // Change several boolean properties
                documentProperties.HyperlinksChanged = true;

                // Save the presentation with changed properties
                documentInfo.UpdateDocumentProperties(documentProperties);
                documentInfo.WriteBindedPresentation(resultPath);
            }
        }
    }
}
