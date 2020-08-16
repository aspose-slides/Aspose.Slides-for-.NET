# PowerPoint File Manipulation API

[Aspose.Slides for .NET](https://products.aspose.com/slides/net) is a cross-platform API that helps in developing applications with the ability to create, manipulate, inspect or convert Microsoft PowerPoint and OpenOffice presentation files without any dependency.

<p align="center">

  <a title="Download complete Aspose.Slides for .NET source code" href="https://github.com/aspose-slides/Aspose.Slides-for-.NET/archive/master.zip">
	<img src="https://raw.github.com/AsposeExamples/java-examples-dashboard/master/images/downloadZip-Button-Large.png" />
  </a>
</p>

Directory | Description
--------- | -----------
[Demos](Demos)  | Source code for live demos hosted at https://products.aspose.app/slides/family.
[Examples](Examples)  | A collection of .NET examples that help you learn the product features.
[Plugins](Plugins)  | Visual Studio Plugins related to Aspose.Slides for .NET.

## Presentation Processing Features

- Intuitive Document Object Model: Aspose.Slides' object model gives complete control over presentation elements such as slides, shapes, frames, charts, multimedia, embedded objects, controls, tables, text, transitions and formatting. Developers can use this object model to create complex PowerPoint File Processing applications that can dynamically generate presentations, create or manipulate presentation slides, apply transitions or add animation effects.
- [File Format Conversion](https://docs.aspose.com/slides/net/supported-file-formats/s): API allows to load & convert PowerPoint presentation, template & slideshow file formats to other supported formats without needing to understand the underlying structure of source or destination formats. The conversion process is simple yet reliable with results identical to the original file in its native application.
- Rendering & Printing: Developers can render the whole presentation or selective slides to fixed layout formats such as PDF & XPS as well as raster & vector image formats including PNG, JPEG, SVG and so on. It is also possible to print presentations via physical or virtual printers.
- Presentation Security: Load protected presentations or control access to presentations, slides or objects via advanced security features.
- Availability of 24 pre-defined textures & 48 patterns for quick styling.

## Read & Write Presentations

**Microsoft PowerPoint:** PPT, PPTX, PPS, POT, PPSX, PPTM, PPSM, POTX, POTM\
**OpenOffice:** ODP, OTP

## Save Presentations As

**Fixed Layout:** PDF, PDF/A, XPS\
**Image:** JPEG, PNG, BMP, TIFF, GIF, SVG\
**Web:** HTML

## Platform Independence

Aspose.Slides for .NET can be used to build any type of a 32-bit or 64-bit .NET application including ASP.NET, WCF & WinForms as well as via COM Interop while using diverse programming languages including C++, VBScript & classic ASP. The package provides assemblies to be used with Mono on various flavors of Linux, .NET Core, Xamarin.Android & Xamarin.Mac.

## Get Started with Aspose.Slides for .NET

Let's give Aspose.Slides for .NET a try! Simply execute `Install-Package Aspose.Slides.NET` from Package Manager Console in Visual Studio to fetch the NuGet package. If you already have Aspose.Slides for .NET and want to upgrade the version, please execute `Update-Package Aspose.Slides.NET` to get the latest version.

## Create a PPTX Presentation from Scratch

```csharp
// instantiate a Presentation object that represents a presentation file
using (Presentation presentation = new Presentation())
{
    // get the first slide
    ISlide slide = presentation.Slides[0];

    // add an autoshape of type line
    slide.Shapes.AddAutoShape(ShapeType.Line, 50, 150, 300, 0);
    presentation.Save(dir + "output.pptx", SaveFormat.Pptx);
}
```

## Convert Specific Slides to PDF Format

```csharp
// instantiate a Presentation object that represents a presentation file
using (Presentation presentation = new Presentation(dir + "template.pptx"))
{
    // setting array of slides positions
    int[] slides = { 1, 3 };
    // save the presentation to PDF
    presentation.Save(dir + "output.pdf", slides, SaveFormat.Pdf);
}
```

[Home](https://www.aspose.com/) | [Product Page](https://products.aspose.com/slides/net) | [Docs](https://docs.aspose.com/slides/net/) | [Demos](https://products.aspose.app/slides/family) | [API Reference](https://apireference.aspose.com/slides/net) | [Examples](https://github.com/aspose-slides/Aspose.Slides-for-.NET) | [Blog](https://blog.aspose.com/category/slides/) | [Free Support](https://forum.aspose.com/c/slides) | [Temporary License](https://purchase.aspose.com/temporary-license)
