---
title: Getting Started
nav_order: 1
---

![Logo]({{ site.baseurl }}\images\logo_transparent.png){: width="200" }

- [Introduction](#introduction)
- [Quickstart](#quickstart)
- [Distribution](#distribution)
- [Compatibility](#compatibility)
  - [.NET](#net)
  - [Excel](#excel)
- [License](#license)
- [Useful Links](#useful-links)

## Introduction

Excel-DNA is an open source and free independent project to integrate .NET languages (C#, Visual Basic.NET and F#) into Excel.

The library is useful for:
- VBA users looking for a more robust development experience - language features, deployment, version control, etc.
- Users of other languages which are in some order compatible with .NET (C/C++, Python, JavaScript, R, Rust, etc.) who want a easier/faster/better integration with Excel.

## Quickstart

1. Install **Visual Studio** - not Visual Studio Code, which has a limited support for .NET. [Visual Studio Community](https://visualstudio.microsoft.com/vs/community/) is free but has restriction for Enterprise users. See [Visual Studio installation](/basic-features/visual-studio-installation) for more help.

2. Create a new **Class Library (.NET Framework)** project in C#, F#, or Visual Basic.
   - If you're using a Visual Studio version compatible with .NET 6, you can create a .NET (not .NET framework) Class Library. This has experimental support on Excel-DNA version 1.6.0-preview.

3. Use the **Manage NuGet Packages** dialog or the Package Manager Console to install the **[ExcelDna.AddIn](https://www.nuget.org/packages/ExcelDna.AddIn/)** package:
    ```
    PM> Install-Package ExcelDna.AddIn
    ```
    The NuGet Package installs the required files and configures your project to build an Excel-DNA add-in.

4. Add one of the following snippets to your code to make your first Excel-DNA function.

   - C#

    ```csharp
    using ExcelDna.Integration;

    public static class MyFunctions
    {
        [ExcelFunction(Description = "My first .NET function")]
        public static string HelloDna(string name)
        {
            return "Hello " + name;
        }
    }
    ```

   - Visual Basic .NET

    ```vb
    Imports ExcelDna.Integration

    Public Module MyFunctions

    <ExcelFunction(Description:="My first .NET function")> _
    Public Function HelloDna(name As String) As String
        Return "Hello " & name
    End Function

    End Module
    ```

   - F#

    ```fsharp
    module MyFunctions

    open ExcelDna.Integration

    [<ExcelFunction(Description="My first .NET function")>]
    let HelloDna name =
        "Hello " + name
    ```

5. Then press F5 (Start Debugging) to compile, run Excel and load the add-in, and enter your function into a cell: =HelloDna("your name")
6. For more details about the installation/build process, check [Excel-DNA installation](/basic-features/excel-dna-installation). 

## Distribution

Excel-DNA add-in users have to install the freely available .NET runtime (.NET Framework or .NET 6 depending on the version of your add-in). The integration is done by an Excel Add-In (.xll file) that exposes .NET code to Excel.

The user code usually resides in compiled .NET libraries (.dll), but text-based script files (C#, Visual Basic or F#) are also supported inside the library configuration file (.dna).

## Compatibility

### .NET

- .NET Framework from 2.0 to 4.8
- .NET 6 experimentally on version 1.6.0-preview.

### Excel

Excel versions 1997 through 2021/365 can be targeted with a single add-in.

Advanced Excel features are supported, including:
- Multi-threaded recalculation
- Asynchronous worksheet functions
- RTD (Real Time Data) servers
- Customized Ribbon interfaces
- Custom Task Panes
- Offloading UDF computations to a Windows HPC cluster
- Etc.

## License

The Excel-DNA Runtime is free for all use, and distributed under a permissive open-source license that also allows commercial use.

## Useful Links
The home page for Excel-DNA is at [http://www.excel-dna.net](http://www.excel-dna.net).

You are also welcome to contact us with questions, comments or suggestions.

- The core library project can be found on [GitHub](https://github.com/Excel-DNA/ExcelDna), where the latest source versions are hosted.
- For general questions and discussion about Excel-DNA, use the [Google group](https://groups.google.com/group/exceldna) or [Stack Overflow](http://stackoverflow.com/questions/tagged/excel-dna). The Google group has an extensible an searchable question base since 2007.
- Specific issues, bug reports and feature requests can be added to the [GitHub Issues](https://github.com/Excel-DNA/ExcelDna/issues) list.


