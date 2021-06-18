using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Slides.Export;
using Aspose.Slides.Export.Xaml;

/*
This example demonstrates how to export a Presentation to a set of XAML files.
*/

namespace Aspose.Slides.Examples.CSharp.Presentations.Conversion
{
    public class ConvertToXaml
    {
        public static void Run()
        {
            // Path to source presentation
            string presentationFileName = Path.Combine(RunExamples.GetDataDir_Conversion(), "XamlEtalon.pptx");

            using (Presentation pres = new Presentation(presentationFileName))
            {
                // Create convertion options
                XamlOptions xamlOptions = new XamlOptions();
                xamlOptions.ExportHiddenSlides = true;

                // Define your own output-saving service
                NewXamlSaver newXamlSaver = new NewXamlSaver();
                xamlOptions.OutputSaver = newXamlSaver;

                // Convert slides
                pres.Save(xamlOptions);

                // Save XAML files to an output directory
                foreach (var pair in newXamlSaver.Results)
                {
                    File.AppendAllText(Path.Combine(RunExamples.OutPath, pair.Key), pair.Value);
                }
            }
        }

        /// <summary>
        /// Represents an output saver implementation for transfer data to the external storage.
        /// </summary>
        class NewXamlSaver : IXamlOutputSaver
        {
            private Dictionary<string, string> m_result =  new Dictionary<string, string>();
            
            public Dictionary<string, string> Results
            {
                get { return m_result; }
            }

            public void Save(string path, byte[] data)
            {
                string name = Path.GetFileName(path);
                Results[name] = Encoding.UTF8.GetString(data);
            }
        }
    }
}