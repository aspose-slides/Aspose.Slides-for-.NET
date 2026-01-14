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

/*
This example demonstrates how to move the sensitivity labels information from the custom document properties 
to the modern SensitivityLabels collection, add the new sensitivity label to the presentation document and 
print the sensitivity labels applied to the presentation. 
*/

namespace CSharp.Presentations.Properties
{
    public class SensitivityLabelsExample
    {
        // The example below demonstrates how to check a password to open a presentation

        public static void Run()
        {
            //Path for source presentation
            string pptxFile = Path.Combine(RunExamples.GetDataDir_PresentationProperties(), "OldSensitivitiLabels.pptx");

            //Out path
            string outPath = Path.Combine(RunExamples.OutPath, "SensitivitiLabels_out.pptx");

            string labelId = "{0372a796-4aa3-4c41-9a98-8232cac474f6}";
            string labelId2 = "{c0c0bc41-48d8-4bf2-a038-8ec8c93813b5}";
            Guid siteId = new Guid("{c336d4c6-89ce-480c-beb0-3bfa5538f186}");

            using (Presentation pres = new Presentation(pptxFile))
            {
                // Get sensitivity labels from the custom document properties
                ISensitivityLabel[] mipSensitivityLabels = pres.DocumentProperties.GetSensitivityLabels();

                ISensitivityLabelCollection sensitivityLabels = pres.SensitivityLabels;
                foreach (var sensitivityLabel in mipSensitivityLabels)
                {
                    // Add label to the collection 
                    sensitivityLabels.Add(sensitivityLabel);
                }

                // Add sensitivity labels
                var label1 = sensitivityLabels.Add(labelId, siteId, true, SensitivityLabelAssignmentType.Standard);
                label1.ContentMarkTypes.Add(SensitivityLabelContentType.Header);
                label1.IsRemoved = true;

                var label2 = sensitivityLabels.Add(labelId2, siteId, true, SensitivityLabelAssignmentType.Privileged);
                label2.ContentMarkTypes.Add(SensitivityLabelContentType.Footer);
                label2.ContentMarkTypes.Add(SensitivityLabelContentType.Watermark);


                // Print sensitivity labels
                foreach (var sensitivityLabel in sensitivityLabels)
                    Console.WriteLine("Label Id " + sensitivityLabel.Id + " from site " + sensitivityLabel.SiteId);

                pres.Save(outPath, SaveFormat.Pptx);
            }
        }
    }
}