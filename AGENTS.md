# AGENTS.md

Guidance for AI coding agents (Claude Code, GitHub Copilot, Cursor, etc.) working in this repository. It reflects the repository's current state so that automated contributions stay consistent.

## Repository overview

Aspose.Slides-for-.NET is a code-examples repository for **Aspose.Slides for .NET**, a commercial, cross-platform library for creating, manipulating, inspecting, and converting Microsoft PowerPoint (PPT, PPTX, PPS, POT, …) and OpenOffice (ODP, OTP) presentation files. The product code itself ships in the `Aspose.Slides.NET` NuGet package; this repository contains only examples, demos, and comparison-style plugin projects.

The repository is updated roughly monthly in sync with Aspose.Slides for .NET releases (see git history: commits titled "Updated w.r.t .NN", e.g. "Updated w.r.t 26.10").

## Directory layout

| Path | Contents |
| ---- | -------- |
| `Demos/` | Source of the free online apps at https://products.aspose.app/slides/family — an ASP.NET Core solution (`Demos/Demo.sln`, projects under `Demos/Aspose.Slides.Web.*`). |
| `Examples/` | The main example collection (see below). |
| `Examples/CSharp/` | C# console project, .NET Framework 4.6.2, default platform x86. One folder per feature area: `Charts`, `Presentations`, `Shapes`, `Slides`, `SmartArts`, `Tables`, `Text`, `VBA`, etc. |
| `Examples/CSharp/RunExamples.cs` | Entry point. `Main()` contains commented-out `Xxx.Run();` calls — uncomment one to execute an example. Also holds the `GetDataDir_*()` helpers mapping to the `Examples/Data/` folder. |
| `Examples/Data/` | Input files and per-feature subfolders used by the examples. |
| `Examples/Aspose.Slides.Examples.CSharp.sln` | Solution file for the C# examples. |
| `Plugins/` | Side-by-side comparisons: "Aspose.Slides vs OpenXML" and "Aspose.Slides vs VSTO". |
| `README.md` | Public landing page. Keep its links and platform claims in sync with the current release. |
| `AGENTS.md` | This file. |

## Build & run

### Examples (C#)

The examples target .NET Framework 4.6.2 and use `packages.config`-style NuGet restore.

- Open `Examples/Aspose.Slides.Examples.CSharp.sln` in Visual Studio, restore the NuGet packages, and run the `CSharp` project. Uncomment the desired `Xxx.Run();` line in `RunExamples.Main()` first.
- Every example class must also be registered as a `<Compile Include="...">` item in `Examples/CSharp/CSharp.csproj`, otherwise it will not be part of the build.

### Demos

```
dotnet run --project Demos\Aspose.Slides.Web\Aspose.Slides.Web.csproj
```

(see `Demos/README.md`).

## Conventions for example code

Each example is a standalone public class exposing a runnable static method:

```csharp
namespace Aspose.Slides.Examples.CSharp.Shapes
{
    public class AddPlainLineToSlide
    {
        public static void Run()
        {
            //ExStart:AddPlainLineToSlide
            string dataDir = RunExamples.GetDataDir_Shapes();
            // ... example body ...
            //ExEnd:AddPlainLineToSlide
        }
    }
}
```

- The namespace mirrors the folder path: `Aspose.Slides.Examples.CSharp.<SubFolder>`.
- The runnable method is `public static void Run()`, invoked from `RunExamples.Main()` as a commented call and referenced by the docs.
- Keep the `//ExStart:<Name>` / `//ExEnd:<Name>` markers around the example body; documentation at docs.aspose.com extracts code samples using them.
- Obtain input/output paths via the matching `RunExamples.GetDataDir_*()` helper (maps to `Examples/Data/<Folder>/`). Ensure the target directory exists before saving (the `Directory.Exists` / `CreateDirectory` pattern).
- Use `using (Presentation pres = new Presentation()) { ... }` and `pres.Save(path, SaveFormat.Xxx);`.
- Add short descriptive comments, no XML doc comments, to match the existing style.

## Adding a new example

1. Create `Examples/CSharp/<CategoryFolder>/<FeatureName>.cs` modeled on an existing file.
2. Add the class with `public static void Run()` and `ExStart`/`ExEnd` markers, using a `GetDataDir_*` helper (add a new helper in `RunExamples.cs` if none fits).
3. Register the file as a `<Compile Include="CategoryFolder\FeatureName.cs" />` entry in `Examples/CSharp/CSharp.csproj`.
4. Add a commented `FeatureName.Run();` line under the matching section in `RunExamples.Main()`.
5. If the example needs an input file, add it to the matching `Examples/Data/<Folder>/` subfolder.

## Packages & versions

- `Examples/CSharp/packages.config` pins NuGet versions; `Examples/CSharp/CSharp.csproj` references them via `HintPath` (`..\packages\...`).
- Current dependencies: `Aspose.Slides.NET` 26.10.0, `Aspose.Cells` 26.9.0 (used by chart examples), plus System.* support packages required on .NET Framework 4.6.2.
- On a monthly release bump, update `packages.config` **and** the matching csproj references/HintPaths together.
- `packages/`, `bin/`, `obj/` and `*.lic` are git-ignored. Never commit restored binaries or license files.

## Rules & pitfalls

- This is an examples repository for a commercial library. Write examples against the public API only (reference: https://reference.aspose.com/slides/net/) and do not rely on undocumented internal behavior.
- Keep commits small and focused — the project history uses single-purpose commits ("Updated w.r.t 26.10", README fixes, etc.). Do not squash unrelated changes together.
- The `Xamarin/` demo and `Aspose.Slides.Examples.Xamarin.sln` were removed (Xamarin is no longer supported) — do not reintroduce them.
- When touching `README.md`, verify version/platform claims and URLs against the current release. Use `https://reference.aspose.com/slides/net/` (not `apireference.aspose.com`) for the API reference.