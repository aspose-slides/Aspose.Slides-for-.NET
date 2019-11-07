using Aspose.Slides;
using Aspose.Slides.Charts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Charts
{
    class SetExternalWorkbookFromNetwork
    {
        public static void Run() {
            //ExStart:SetExternalWorkbookFromNetwork

            string externalWbPath = @"http://606178d2.ngrok.io/webgrind/styles/2.xlsx";
            LoadOptions opts = new LoadOptions();
            opts.ResourceLoadingCallback = new WorkbookLoadingHandler();

            using (Presentation pres = new Presentation(opts))
            {
                IChart chart = pres.Slides[0].Shapes.AddChart(ChartType.Pie, 50, 50, 400, 600, false);
                IChartData chartData = chart.ChartData;

                (chartData as ChartData).SetExternalWorkbook(externalWbPath);
            }

            //ExEnd:SetExternalWorkbookFromNetwork

        }
    }


    class WorkbookLoadingHandler : IResourceLoadingCallback
    {
        public ResourceLoadingAction ResourceLoading(IResourceLoadingArgs args)
        {
            string workbookPath = args.OriginalUri;

            if (workbookPath.IndexOf(':') > 1 && !workbookPath.StartsWith("file:///")) // schemed path
            {
                try
                {
                    WebRequest request = WebRequest.Create(workbookPath);
                    request.Credentials = new System.Net.NetworkCredential("testuser", "testuser");
                    using (WebResponse response = request.GetResponse())
                    using (Stream responseStream = response.GetResponseStream())
                    {
                        //byte[] buffer = BlobDownloadManager.Download(responseStream);
                        // args.SetData(buffer);
                        return ResourceLoadingAction.UserProvided;
                    }
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException(ex.ToString());
                }
            }
            else
            {
                return ResourceLoadingAction.Default;
            }
        }
    }
}
