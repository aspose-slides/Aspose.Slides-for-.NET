using Aspose.Slides.Export;
using Aspose.Slides.Vba;
using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.VBA
{
    class AddVBAMacros
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_VBA();

            // ExStart:AddVBAMacros
            // Instantiate Presentation
            using (Presentation presentation = new Presentation())
            {
                // Create new VBA Project
                presentation.VbaProject = new VbaProject();

                // Add empty module to the VBA project
                IVbaModule module = presentation.VbaProject.Modules.AddEmptyModule("Module");
              
                // Set module source code
                module.SourceCode = @"Sub Test(oShape As Shape) MsgBox ""Test"" End Sub";

                // Create reference to <stdole>
                VbaReferenceOleTypeLib stdoleReference =
                    new VbaReferenceOleTypeLib("stdole", "*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation");

                // Create reference to Office
                VbaReferenceOleTypeLib officeReference =
                    new VbaReferenceOleTypeLib("Office", "*\\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.0#0#C:\\Program Files\\Common Files\\Microsoft Shared\\OFFICE14\\MSO.DLL#Microsoft Office 14.0 Object Library");

                // Add references to the VBA project
                presentation.VbaProject.References.Add(stdoleReference);
                presentation.VbaProject.References.Add(officeReference);

            
                // Save Presentation
                presentation.Save(dataDir + "AddVBAMacros_out.pptm", SaveFormat.Pptm);
            }
            // ExStart:AddVBAMacros
        }
    }
}