
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using Aspose.Slides.Live.Demos.UI.Models;
using Aspose.Slides.Live.Demos.UI.Services;
using System.Collections.ObjectModel;
using Aspose.Slides;

namespace Aspose.Slides.Live.Demos.UI.Controllers
{
	/// <summary>
	/// AsposeSlidesMetadataController class
	/// </summary>
	public class AsposeSlidesMetadataController : SlidesBase
	{
		///<Summary>
		/// Properties method to get metadata
		///</Summary>
		///
		[HttpPost]
		public HttpResponseMessage Properties(string folderName, string fileName)
		{
			try
			{				
				return Request.CreateResponse(HttpStatusCode.OK, new PropertiesResponse(Path.Combine(Config.Configuration.WorkingDirectory, folderName, fileName)));
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
				return Request.CreateResponse(HttpStatusCode.InternalServerError, ex.Message);
			}
		}

		private Aspose.Slides.Export.SaveFormat GetFormatFromSource(string sourceFile)
		{
			switch (Path.GetExtension(sourceFile))
			{
				case ".ppt":
					return Aspose.Slides.Export.SaveFormat.Ppt;
				case ".odp":
					return Aspose.Slides.Export.SaveFormat.Odp;
				case ".pptx":
				default:
					return Aspose.Slides.Export.SaveFormat.Pptx;
			}
		}
		/// <summary>
		/// PropertiesResponse
		/// </summary>
		private class PropertiesResponse
		{
			public List<DocProperty> BuiltIn { get; set; }
			public List<DocProperty> Custom { get; set; }	

			public PropertiesResponse(string  path)
			{
				Aspose.Slides.Presentation presentation = new Aspose.Slides.Presentation(path);
				IDocumentProperties properties = presentation.DocumentProperties;

				BuiltIn = new List<DocProperty>();
				Custom = new List<DocProperty>();

				BuiltIn.Add(new DocProperty() { Name = "ApplicationTemplate", Value = properties.ApplicationTemplate, Type = PropertyType.String });
				BuiltIn.Add(new DocProperty() { Name = "AppVersion", Value = properties.AppVersion, Type = PropertyType.String });
				BuiltIn.Add(new DocProperty() { Name = "Author", Value = properties.Author, Type = PropertyType.String });
				BuiltIn.Add(new DocProperty() { Name = "Category", Value = properties.Category, Type = PropertyType.String });
				BuiltIn.Add(new DocProperty() { Name = "Comments", Value = properties.Comments, Type = PropertyType.String });
				BuiltIn.Add(new DocProperty() { Name = "Company", Value = properties.Company, Type = PropertyType.String });
				BuiltIn.Add(new DocProperty() { Name = "ContentStatus", Value = properties.ContentStatus, Type = PropertyType.String });
				BuiltIn.Add(new DocProperty() { Name = "ContentType", Value = properties.ContentType, Type = PropertyType.String });
				BuiltIn.Add(new DocProperty() { Name = "HyperlinkBase", Value = properties.HyperlinkBase, Type = PropertyType.String });

				BuiltIn.Add(new DocProperty() { Name = "Keywords", Value = properties.Keywords, Type = PropertyType.String });
				BuiltIn.Add(new DocProperty() { Name = "LastSavedBy", Value = properties.LastSavedBy, Type = PropertyType.String });
				BuiltIn.Add(new DocProperty() { Name = "Manager", Value = properties.Manager, Type = PropertyType.String });
				BuiltIn.Add(new DocProperty() { Name = "NameOfApplication", Value = properties.NameOfApplication, Type = PropertyType.String });
				BuiltIn.Add(new DocProperty() { Name = "PresentationFormat", Value = properties.PresentationFormat, Type = PropertyType.String });
				BuiltIn.Add(new DocProperty() { Name = "Subject", Value = properties.Subject, Type = PropertyType.String });

				BuiltIn.Add(new DocProperty() { Name = "Title", Value = properties.Title, Type = PropertyType.String });
				BuiltIn.Add(new DocProperty() { Name = "CreatedTime", Value = properties.CreatedTime, Type = PropertyType.DateTime });
				BuiltIn.Add(new DocProperty() { Name = "LastPrinted", Value = properties.LastPrinted, Type = PropertyType.DateTime });
				BuiltIn.Add(new DocProperty() { Name = "LastSavedTime", Value = properties.LastSavedTime, Type = PropertyType.DateTime });
				BuiltIn.Add(new DocProperty() { Name = "RevisionNumber", Value = properties.RevisionNumber, Type = PropertyType.Number });
				BuiltIn.Add(new DocProperty() { Name = "SharedDoc", Value = properties.SharedDoc, Type = PropertyType.Boolean });
				BuiltIn.Add(new DocProperty() { Name = "TotalEditingTime", Value = properties.TotalEditingTime, Type = PropertyType.Time });

				for (int i = 0; i < properties.CountOfCustomProperties; ++i)
				{
					string name = properties.GetCustomPropertyName(i);
					string _propertyValue = "";
					properties.GetCustomPropertyValue(name, out _propertyValue);
					Custom.Add(new DocProperty() { Name = name, Value = _propertyValue, Type = PropertyType.String });
				}				
			}
		}
		/// <summary>
		/// The same fields as in DocumentProperty
		/// </summary>
		private class DocProperty
		{
			public string Name { get; set; }
			public object Value { get; set; }
			public PropertyType Type { get; set; }
		}
		public enum PropertyType
		{
			/// <summary>The property is a boolean value.</summary>
			Boolean,
			/// <summary>The property is a date time value.</summary>
			DateTime,
			/// <summary>The property is a floating number.</summary>
			Double,
			/// <summary>The property is an integer number.</summary>
			Number,
			/// <summary>The property is a string value.</summary>
			String,
			/// <summary>The property is an array of strings.</summary>
			StringArray,
			/// <summary>The property is an array of objects.</summary>
			ObjectArray,
			/// <summary>The property is an array of bytes.</summary>
			ByteArray,
			/// <summary>The property is an time.</summary>
			Time,
			/// <summary>The property is some other type.</summary>
			Other,
		}

		///<Summary>
		/// Properties method. Should include 'FileName', 'id', 'properties' as params
		///</Summary>
		[HttpPost]
		[AcceptVerbs("GET", "POST")]
		public Response Download()
		{
			Opts.AppName = "MetadataApp";
			Opts.MethodName = "Download";
			try
			{
				var request = Request.Content.ReadAsAsync<JObject>().Result;
				Opts.FileName = Convert.ToString(request["FileName"]);
				Opts.ResultFileName = Opts.FileName;
				Opts.FolderName = Convert.ToString(request["id"]);
				Presentation _presentation = new Presentation(Opts.WorkingFileName);

				var pars = request["properties"]["BuiltIn"].ToObject<List<DocProperty>>();
				SetBuiltInProperties(_presentation, pars);
				pars = request["properties"]["Custom"].ToObject<List<DocProperty>>();
				SetCustomProperties(_presentation, pars);

				return Process((inFilePath, outPath, zipOutFolder) => { _presentation.Save(outPath, GetFormatFromSource(Opts.WorkingFileName)); });
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
				return new Response
				{
					Status = "500 " + ex.Message,
					StatusCode = 500
				};
			}
		}

		///<Summary>
		/// Properties method. Should include 'FileName', 'id' as params
		///</Summary>
		[HttpPost]
		[AcceptVerbs("GET", "POST")]
		public Response Clear()
		{
			Opts.AppName = "MetadataApp";
			Opts.MethodName = "Clear";
			try
			{
				var request = Request.Content.ReadAsAsync<JObject>().Result;
				Opts.FileName = Convert.ToString(request["FileName"]);
				Opts.ResultFileName = Opts.FileName;
				Opts.FolderName = Convert.ToString(request["id"]);

				Presentation _presentation = new Presentation(Opts.WorkingFileName);
				_presentation.DocumentProperties.ClearBuiltInProperties();
				_presentation.DocumentProperties.ClearCustomProperties();

				return Process((inFilePath, outPath, zipOutFolder) => { _presentation.Save(outPath, GetFormatFromSource(Opts.WorkingFileName)); });
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
				return new Response
				{
					Status = "500 " + ex.Message,
					StatusCode = 500
				};
			}
		}

		/// <summary>
		/// SetBuiltInProperties
		/// </summary>
		/// <param name="Presentation"></param>
		/// <param name="pars"></param>
		private void SetBuiltInProperties(Presentation presentation, List<DocProperty> pars)
		{
			//presentation.DocumentProperties.ClearBuiltInProperties();
			
			foreach (var par in pars)
			{
				
					switch (par.Name)
					{
						case "ApplicationTemplate":
						presentation.DocumentProperties.ApplicationTemplate = par.Value.ToString();
						break;
					case "Author":
						presentation.DocumentProperties.Author = par.Value.ToString();
						break;
					case "Category":
						presentation.DocumentProperties.Category = par.Value.ToString();
						break;
					case "Comments":
						presentation.DocumentProperties.Comments = par.Value.ToString();
						break;
					case "Company":
						presentation.DocumentProperties.Company = par.Value.ToString();
						break;
					case "ContentStatus":
						presentation.DocumentProperties.ContentStatus = par.Value.ToString();
						break;
					case "ContentType":
						presentation.DocumentProperties.ContentType = par.Value.ToString();
						break;
					case "HyperlinkBase":
						presentation.DocumentProperties.HyperlinkBase = par.Value.ToString();
						break;
					case "Keywords":
						presentation.DocumentProperties.Keywords = par.Value.ToString();
						break;
					case "LastSavedBy":
						presentation.DocumentProperties.LastSavedBy = par.Value.ToString();
						break;
					case "Manager":
						presentation.DocumentProperties.Manager = par.Value.ToString();
						break;
					case "NameOfApplication":
						presentation.DocumentProperties.NameOfApplication = par.Value.ToString();
						break;
					case "PresentationFormat":
						presentation.DocumentProperties.PresentationFormat = par.Value.ToString();
						break;
					case "Subject":
						presentation.DocumentProperties.Subject = par.Value.ToString();
						break;
					case "Title":
						presentation.DocumentProperties.Title = par.Value.ToString();
						break;
					case "CreatedTime":
						presentation.DocumentProperties.CreatedTime =  DateTime.Parse( par.Value.ToString());
						break;
					case "LastPrinted":
						presentation.DocumentProperties.LastPrinted = DateTime.Parse( par.Value.ToString());
						break;
					case "RevisionNumber":
						presentation.DocumentProperties.RevisionNumber = int.Parse( par.Value.ToString());
						break;
					case "SharedDoc":
						presentation.DocumentProperties.SharedDoc = bool.Parse(par.Value.ToString());
						break;
					case "TotalEditingTime":
						presentation.DocumentProperties.TotalEditingTime = TimeSpan.Parse(par.Value.ToString());
						break;

				}

			}
		}


		/// <summary>
		/// SetCustomProperties
		/// </summary>
		/// <param name="Presentation"></param>
		/// <param name="pars"></param>
		private void SetCustomProperties(Presentation presentation, List<DocProperty> pars)
		{
			presentation.DocumentProperties.ClearCustomProperties();

			foreach (var par in pars)
			{
				if (par.Name != null)
				{
					presentation.DocumentProperties[par.Name] = par.Value.ToString();
				}

			}
			
		}

	}

}
