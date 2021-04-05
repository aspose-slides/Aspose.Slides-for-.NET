using System.Collections.Generic;
using Aspose.Slides.Web.UI.Models.Interfaces;

namespace Aspose.Slides.Web.UI.Models
{
    internal class ShowroomModel : BaseViewModel, IShowroomModel
    {
        public List<Showcase> Showcases { get; set; } = new()
        {
            new()
            {
                Name = "Conversion",
                Url = "conversion",
                Description =
                    "Convert your PowerPoint presentations to PDF, Open Office, HTML, XPS, Image and other formats"
            },
            new()
            {
                Name = "Metadata", Url = "metadata",
                Description = "View and edit your PowerPoint and OpenOffice presentation's metadata properties"
            },
            new() {Name = "Viewer", Url = "viewer", Description = "View your PowerPoint file online from anywhere"},
            new()
            {
                Name = "Editor", Url = "editor",
                Description = "Edit PowerPoint presentations files online from anywhere"
            },
            new()
            {
                Name = "Presentation to Video", Url = "video", Description = "Convert your presentation to video online"
            },
            new()
            {
                Name = "Annotation",
                Url = "annotation",
                Description = "Remove annotations from your PowerPoint presentations online from anywhere"
            },
            new() {Name = "Search", Url = "search", Description = "Search text in your PowerPoint presentations"},
            new() {Name = "Redaction", Url = "redaction", Description = "Search and replace text in your documents"},
            new()
            {
                Name = "Watermark", Url = "watermark",
                Description = "Add or Remove watermark in PowerPoint presentations files online from anywhere"
            },
            new() {Name = "Parser", Url = "parser", Description = "Extract text and images from your document"},
            new()
            {
                Name = "Unlock", Url = "unlock", Description = "Unlock your password protected PowerPoint and ODP file"
            },
            new()
            {
                Name = "Lock", Url = "lock",
                Description = "Lock your PowerPoint and ODP file by making it password protected"
            },
            new() {Name = "Merger", Url = "merger", Description = "Merge Microsoft PowerPoint documents"},
            new()
            {
                Name = "Splitter", Url = "splitter",
                Description = "Split PowerPoint presentations files online from anywhere"
            },
            new()
            {
                Name = "Signature", Url = "signature",
                Description = "Sign presentations with a handwritten, graphical or text signature"
            },
            new()
            {
                Name = "Charts", Url = "chart",
                Description = "Create charts from table data in Microsoft Excel and OpenOffice documents."
            },
            new()
            {
                Name = "Comparison",
                Url = "comparison",
                Description = "Compare the text contents of two Presentations documents online"
            },
            new()
            {
                Name = "Convert to PowerPoint", Url = "import",
                Description = "Convert PDFs or images to Presentation online"
            },
            new()
            {
                Name = "Remove Macros", Url = "remove-macros",
                Description = "Remove all active content(Macros) from Presentation online"
            }
        };

        public string IndexPageTitle { get; } = "Free PowerPoint Presentation Apps";
        public string IndexPageSubTitle { get; } = "Convert or View your PowerPoint presentations online from anywhere";
        public string ProductFamilyInclude { get; } = "{0} Product Family Includes";
        public string AsposeSlides { get; } = "Aspose.Slides";
    }
}