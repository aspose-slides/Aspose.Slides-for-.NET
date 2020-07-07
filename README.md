# Aspose.Slides for .NET

[Aspose.Slides for .NET](https://products.aspose.com/slides/net) is a unique PowerPoint® management API that enables .NET applications to read, write and manipulate PowerPoint documents without using Microsoft PowerPoint.

<p align="center">

  <a title="Download complete Aspose.Slides for .NET source code" href="https://github.com/aspose-slides/Aspose.Slides-for-.NET/archive/master.zip">
	<img src="https://raw.github.com/AsposeExamples/java-examples-dashboard/master/images/downloadZip-Button-Large.png" />
  </a>
</p>

This repository contains [Demos](Demos), [Examples](Examples), [Plugins](Plugins) and Showcase projects for [Aspose.Slides for .NET](https://products.aspose.com/slides/net) to help you learn and write your own applications.

Directory | Description
--------- | -----------
[Demos](Demos)  | Aspose.Slides for .NET Live Demos Source Code
[Examples](Examples)  | A collection of .NET examples that help you learn the product features
[Plugins](Plugins)  | Plugins that will demonstrate one or more features of Aspose.Slides for .NET


# Presentation Manipulation .NET API

[Aspose.Slides for .NET](https://products.aspose.com/slides/net) is a cross-platform API that helps in developing applications with the ability to create, manipulate, inspect or convert Microsoft PowerPoint and OpenOffice presentation files without any dependency.

## Presentation Processing Features

- Intuitive Document Object Model: Aspose.Slides' object model gives complete control over presentation elements such as [slides](https://docs.aspose.com/display/slidesnet/Presentation+Slide), [shapes](https://docs.aspose.com/display/slidesnet/Powerpoint+Shapes), frames, [charts](https://docs.aspose.com/display/slidesnet/Powerpoint+Charts), multimedia, embedded [objects](https://docs.aspose.com/display/slidesnet/OLE), [controls](https://docs.aspose.com/display/slidesnet/ActiveX), [tables](https://docs.aspose.com/display/slidesnet/Powerpoint+Table), text, [transitions](https://docs.aspose.com/display/slidesnet/PowerPoint+Animation) and formatting. Developers can use this object model to create complex PowerPoint File Processing applications that can dynamically generate presentations, create or manipulate presentation slides, apply transitions or add animation effects.
- [File Format Conversion](https://docs.aspose.com/display/slidesnet/Supported+File+Formats#SupportedFileFormats-SupportedFileFormats): API allows to load & convert PowerPoint presentation, template & slideshow file formats to other supported formats without needing to understand the underlying structure of source or destination formats. The conversion process is simple yet reliable with results identical to the original file in its native application.
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

## Getting Started with Aspose.Slides for .NET

Let's give Aspose.Slides for .NET a try! Simply execute `Install-Package Aspose.Slides.NET` from Package Manager Console in Visual Studio to fetch the NuGet package. If you already have Aspose.Slides for .NET and want to upgrade the version, please execute `Update-Package Aspose.Slides.NET` to get the latest version.

## Create a PPTX Presentation from Scratch with C# Code

You can execute below code snippet to see how Aspose.Slides API performs in your environment or check the [GitHub Repository](https://github.com/aspose-slides/Aspose.Slides-for-.NET) for other common usage scenarios. 

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

## Convert Specific Slides to PDF Format using C# Code

Aspose.Slides for .NET works as an independent rendering engine for presentations and slides with flexibly overriding certain aspects such as converting specific PowerPoint slides to PDF format.

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

[Product Page](https://products.aspose.com/slides/net) | [Docs](https://docs.aspose.com/display/slidesnet/Home) | [Demos](https://products.aspose.app/slides/family) | [API Reference](https://apireference.aspose.com/slides/net) | [Examples](https://github.com/aspose-slides/Aspose.Slides-for-.NET) | [Blog](https://blog.aspose.com/category/slides/) | [Free Support](https://forum.aspose.com/c/slides) | [Temporary License](https://purchase.aspose.com/temporary-license)
