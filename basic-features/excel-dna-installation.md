---
title: Excel-DNA Installation
nav_order: 2
parent: Basic Features
---
### Configuration file (.dna)

Excel-DNA is configured through a .dna file, which is basically a XML file that contains settings such as:
- Packing: bundling all your .dll libraries into a single .xll for easier usage/distribution.
- Excel COM server: used for exposing your add-in functions in VBA. COM stands for Component Object Model, which is a standard used for communication within Microsoft applications.
- Ribbon layout.
- Etc.

Minimal example of a .dna file:

```xml
<?xml version="1.0" encoding="utf-8"?>
<DnaLibrary Name="ProjectName Add-In" RuntimeVersion="v4.0" xmlns="http://schemas.excel-dna.net/addin/2020/07/dnalibrary">
  <ExternalLibrary Path="ProjectName.dll" ExplicitExports="false" LoadFromBytes="true" Pack="true" IncludePdb="false" />
</DnaLibrary>
```

The add-in file (.xll) generated on build will have the same name of the .dna file.

### Installation

After installation of the Excel-DNA NuGet, the following will happen:

1. If using .NET Framework:
  -  The add-in will be referenced in packages.config file in your project (the 'old' style package import).
  - A file called **ProjectName-AddIn.dna** will be added automatically to your project, and set to be copied to the output directory.
  - Under your Properties item group, a build properties file called ExcelDna.Build.props will be added. This file allows build customization, like configuring whether the packing tool will be run.

2. If using .NET 6:
   - There will be a &lt;PackageReference&gt; tag in your .csproj / .vbproj / .fsproj project file.
   - In this case the .dna file is not added to your project when the package in installed, but is created automatically in the output directory when the project builds.
   - If you want to customize the settings, please copy this .dna file into the project folder or create a new .dna file like the example in the previous section.
  
A reference to **ExcelDna.Integration.dll** will be also added. This contains helper classes like ExcelDnaUtil and ExcelFunctionAttribute that you may use in your add-in.

Upon compilation, the project debugging settings will be configured to start Excel and load the appropriate (32-bit or 64-bit) unpacked version of your add-in, **ProjectName-AddIn.xll** or **ProjectName-AddIn64.xll**.

After building your project you will find (at least) the following files in your output directory (typically bin\Debug or bin\Release):

* **ProjectName.dll** & **ProjectName.pdb** - the normal outputs from compiling your library.
* **ProjectName-AddIn.xll** - the 32-bit native add-in loader for the unpacked add-in.
* **ProjectName-AddIn.dna** - a copy of the add-in master file for the 32-bit unpacked add-in.
* **ProjectName-AddIn64.xll** - the 64-bit native add-in loader for the unpacked add-in.
* **ProjectName-AddIn64.dna** - a copy of the add-in master file for the 64-bit unpacked add-in.
* **ProjectName-AddIn-packed.xll** - the 32-bit packed (all-in-one) redistributable add-in.
* **ProjectName-AddIn64-packed.xll** - the 64-bit packed (all-in-one) redistributable add-in.

For redistribution (if everything is set up correctly) you only need the two (32-bit and 64-bit) -packed.xll files. These files can be renamed as you like.

### Troubleshooting

* If Excel does not open after compiling, check that the path under the 'Debug Properties' is correct. If not, make sure that you have rebuilt the project successfully - this should automatically configure the debug options.
* If Excel starts but no add-in is loaded, check the Excel security settings under File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings. Any option is fine _except_ "Disable all macros without notification."
* If Excel starts but you get a message saying "The file you are trying to open, [...], is in a different format than specified by the file extension.", then there is a mismatch between the bitness (32 or 64-bit) of Excel and the add-in being loaded.

### More about .dna files

Additional referenced assemblies can be specified by adding 'Reference' tags. 
If you specify `Pack="true"` these libraries will be packed into the -packed.xll file and loaded at runtime as needed.
For example:

```xml
<DnaLibrary Name="ProjectName Add-In" RuntimeVersion="v4.0" xmlns="http://schemas.excel-dna.net/addin/2020/07/dnalibrary">
  <ExternalLibrary Path="ProjectName.dll" ExplicitExports="false" LoadFromBytes="true" Pack="true" IncludePdb="false" />
  <Reference Path="Another.Library.dll" Pack="true" />
</DnaLibrary>
```

It is possible to include functions directly in the .dna file. These functions can be altered without the need for compilation. In this case, specify the language with the `Language="CS"` tag.

```csharp
<DnaLibrary Language="CS" Name="ProjectName Add-In" RuntimeVersion="v4.0" xmlns="http://schemas.excel-dna.net/addin/2020/07/dnalibrary">
  <ExternalLibrary Path="ProjectName.dll" ExplicitExports="false" LoadFromBytes="true" Pack="true" IncludePdb="false" />
  <![CDATA[
    using ExcelDna.Integration;

      public static class MyFunctions
      {
          [ExcelFunction(Description = "My first .NET function")]
          public static string HelloDna(string name)
          {
              return "Hello " + name;
          }
      }
  ]]>
</DnaLibrary>
```

Excel-DNA also allows the XML for ribbon UI extensions to be specified in the .dna file.