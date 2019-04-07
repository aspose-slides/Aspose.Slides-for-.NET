using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Presentations.Opening
{
    class LoadPresentation
    {
        public static void Run() {

            //ExStart:LoadPresentation
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_PresentationOpening();

            LoadOptions opts = new LoadOptions();
            opts.ResourceLoadingCallback = new ImageLoadingHandler();
            Presentation presentation = new Presentation(dataDir + "presentation.pptx", opts);
            //ExEnd:LoadPresentation
        }
    }

    //ExStart:IResourceLoadingCallback
    public class ImageLoadingHandler : IResourceLoadingCallback
    {
        // The path to the documents directory.
        string dataDir = RunExamples.GetDataDir_PresentationOpening();
        public ResourceLoadingAction ResourceLoading(IResourceLoadingArgs args)
        {
            if (args.OriginalUri.EndsWith(".jpg"))
            {
                try // load substitute image
                {
                    byte[] imageBytes = File.ReadAllBytes(Path.Combine(dataDir, "aspose-logo.jpg"));
                    args.SetData(imageBytes);
                    return ResourceLoadingAction.UserProvided;
                }
                catch (Exception)
                {
                    return ResourceLoadingAction.Skip;
                }
            }
            else if (args.OriginalUri.EndsWith(".png"))
            {
                // set substitute url
                args.Uri = "http://www.google.com/images/logos/ps_logo2.png";
                return ResourceLoadingAction.Default;
            }

            // skip all other images
            return ResourceLoadingAction.Skip;
        }
    }
    //ExEnd:IResourceLoadingCallback
}
