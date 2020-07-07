using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Http;
using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;
using Aspose.Slides.Live.Demos.UI.Config;


namespace Aspose.Slides.Live.Demos.UI
{
	public class Global : HttpApplication
	{
		
		protected void Application_Error(object sender, EventArgs e)
		{			
			
		}

		void Application_Start(object sender, EventArgs e)
		{
			AreaRegistration.RegisterAllAreas();
			GlobalConfiguration.Configure(WebApiConfig.Register);
			RouteConfig.RegisterRoutes(RouteTable.Routes);			
			BundleConfig.RegisterBundles(BundleTable.Bundles);
			RegisterCustomRoutes(RouteTable.Routes);
		}
		void Session_Start(object sender, EventArgs e)
		{
			//Check URL to set language resource file
			string _language = "EN";
			
			SetResourceFile(_language);
		}

		private void SetResourceFile(string strLanguage)
		{
			if (Session["AsposeSlidesResources"] == null)
				Session["AsposeSlidesResources"] = new GlobalAppHelper(HttpContext.Current, Application, Configuration.ResourceFileSessionName, strLanguage);
		}
		
			void RegisterCustomRoutes(RouteCollection routes)
		{
			routes.RouteExistingFiles = true;
			routes.Ignore("{resource}.axd/{*pathInfo}");
					

			routes.MapRoute(
				name: "Default",
				url: "Default",
				defaults: new { controller = "Home", action = "Default" }
			);
			
			routes.MapRoute(
				"AsposeSlidesConversionRoute",
				"{product}/Conversion",
				 new { controller = "Conversion", action = "Conversion" }
			);
			routes.MapRoute(
				"AsposeSlidesRemoveAnnotationRoute",
				"annotation/remove",
				 new { controller = "Annotation", action = "Remove" }
			);
			routes.MapRoute(
				"AsposeSlidesUnlockRoute",
				"{product}/unlock",
				 new { controller = "Unlock", action = "Unlock" }
			);
			routes.MapRoute(
				"AsposeSlidesSearchRoute",
				"{product}/search",
				 new { controller = "Search", action = "Search" }
			);
			routes.MapRoute(
				"AsposeSlidesSplitterRoute",
				"{product}/splitter",
				 new { controller = "Splitter", action = "Splitter" }
			);
			routes.MapRoute(
				"AsposeSlidesRedactionRoute",
				"{product}/redaction",
				 new { controller = "Redaction", action = "Redaction" }
			);
			routes.MapPageRoute(
				"AsposeSlidesWatermarkRoute",
				"slides/watermark",
				"~/Watermark/WatermarkSlides.aspx"
			);
			
			routes.MapRoute(
				"AsposeSlidesParserRoute",
				"{product}/parser",
				 new { controller = "Parser", action = "Parser" }
			);
			routes.MapRoute(
				"AsposeSlidesAnnotationRoute",
				"{product}/annotation",
				 new { controller = "Annotation", action = "Annotation" }
			);
			routes.MapRoute(
				"AsposeSlidesMetadataRoute",
				"{product}/metadata",
				 new { controller = "Metadata", action = "Metadata" }
			);
			routes.MapRoute(
				"AsposeSlidesMergerRoute",
				"{product}/merger",
				 new { controller = "Merger", action = "Merger" }
			);
			routes.MapRoute(
				"AsposeSlidesProtectRoute",
				"{product}/lock",
				 new { controller = "Lock", action = "Lock" }
			);
			routes.MapRoute(
				"AsposeSlidesViewerRoute",
				"{product}/viewer",
				 new { controller = "Viewer", action = "Viewer" }
			);
			routes.MapPageRoute(
			  "AsposeSlidesDefaultViewerRoute",
			  "slides/view",
			  "~/ViewerApp/Default.aspx"
			);
			
			routes.MapRoute(
				"DownloadFileRoute",
				"common/download",
				new { controller = "Common", action = "DownloadFile" }				
				
			);
			routes.MapRoute(
				"UploadFileRoute",
				"common/uploadfile",
				new { controller = "Common", action = "UploadFile" }

			);
		}

		private void MapProductToolPageRoute(RouteCollection routes, string routeName, string routeUrl, string physicalFile, string productRegex)
		{
			routes.MapPageRoute(routeName, routeUrl, physicalFile, false, null, new RouteValueDictionary { { "Product", productRegex } });
		}
	}
}
