
- [Excel-DNA Packing Tool](exceldna-packing-tool.md) The packing utility allow you to pack your add-in into a single `.xll` file for easy distribution.
- [Installing your add-in](installing-your-add-in.md) and running generally.
- Accepting [Range Parameters](range-parameters.md) in UDFs.
- [Integrating with VBA](integrating-with-vba.md)
- [Performing Asynchronous Work](performing-asynchronous-work.md)
- [Optional Parameters and Default Values](optional-parameters-and-default-values.md)
- [Keyboard Shortcut](keyboard-shortcut.md)
- [Excel Programming Interfaces](excel-programming-interfaces.md)
    * [Using the Excel COM Automation Interfaces](using-the-excel-com-automation-interfaces.md)
    * [Excel C API](excel-c-api.md)
- [Ribbon Customization](ribbon-customization.md) and various ribbon links.
- A note on [AutoClose and Detecting Excel Shutdown](autoclose-and-detecting-excel-shutdown.md).
- [Debugging Notes](debugging-notes.md)
- [COM Server Support](com-server-support.md)
- Some notes on [FSharp Type Inference](fsharp-type-inference.md), and [FSharp Standalone Assemblies](fsharp-standalone-assemblies.md).
- [Asynchronous Functions](asynchronous-functions.md)
- [Asynchronous Functions with Tasks](asynchronous-functions-with-tasks.md) example in VB.NET.
- [Reactive Extensions for Excel](reactive-extensions-for-excel.md)
- [Dynamic delegate registration](dynamic-delegate-registration.md) - an advances feature to implement runtime registration and function wrappers.
- [User settings and the .xll.config file](user-settings-and-the-xllconfig-file.md)
- A step-by-step guide to build a new add-in using the NuGet package, and then [Configure NLog logging](configure-nlog-logging.md) for your add-in.
- [Creating a help file](creating-a-help-file.md)
- [Returning 1-D Arrays](returning-1-d-arrays.md)
- [Async macro example - formatting the calling cell from a UDF](async-macro-example-formatting-the-calling-cell-from-a-udf.md)
- [Enumerating Excel COM Automation collections in VB.NET](enumerating-excel-com-automation-collections-in-vbnet.md)
- [Modal dialog on new thread](modal-dialog-on-new-thread.md)
- [Utilizing custom XML parts in Excel Workbooks](customXMLparts.md)
- [Checking for updates and download of Excel-DNA Addins (or other .NET based programs)](update-check-download.md)

## Samples

Various sample projects and snippets related to Excel-DNA are available in the [Samples](https://github.com/Excel-DNA/Samples) repository.

Additional samples are available in the [Distribution/Samples](https://github.com/Excel-DNA/ExcelDna/tree/master/Distribution/Samples) folder and contains various .dna files, each of which is a self-contained add-in that exhibit some Excel-DNA functionality.

The .dna files are .xml files that can be edited with a regular text editor.

To run any of the sample .dna files, make a copy of the `Distribution\ExcelDna.xll` file, place it next to the .dna file, and rename to have the same prefix. E.g. to run `Optional.dna`, make a copy of `ExcelDna.xll` called `Optional.xll`, and double-click, or `File->Open` to load in Excel.

### Power Tools
- [ExcelDnaDoc](https://github.com/Excel-DNA/ExcelDnaDoc) provides tools to make help generation easier.
- [Registration](https://github.com/Excel-DNA/Registration) allows the automatic generation of parameter and function conversions, removing boiler-plate code for optional parameters, async functions etc.
- [IntelliSense](https://github.com/Excel-DNA/IntelliSense) add in-sheet IntelliSense for Excel UDFs.

## Community Projects

- [ExcelDna-Unpack](https://github.com/augustoproiete/exceldna-unpack) is a command-line utility to extract the contents of Excel-DNA add-ins packed with [ExcelDnaPack](exceldna-packing-tool.md)
- [ExcelDna.Abstractions](https://github.com/augustoproiete/exceldna-abstractions) facilitates mocking & unit testing of Excel-DNA Add-Ins
- [ExcelDna.WiXInstaller](https://github.com/Excel-DNA/WiXInstaller) is a user-contributed template (thank you very much to Lee Zeitz!) for making a WiX-based installer for an Excel-DNA add-in.
- [ExcelDna.StrongName](https://github.com/Excel-DNA/ExcelDna.StrongName) provides strong name key pair used to sign Excel-DNA assemblies.
- [ExcelDna.Diagnostics.Serilog](https://github.com/augustoproiete/exceldna-diagnostics-serilog) integrates Excel-DNA Diagnostic Logging with your Serilog logging pipeling within your Excel-DNA Add-In
- [Serilog.Sinks.ExcelDnaLogDisplay](https://github.com/augustoproiete/serilog-sinks-exceldnalogdisplay) is a Serilog sink that writes events to Excel-DNA LogDisplay
- [Serilog.Enrichers.ExcelDna](https://github.com/augustoproiete/serilog-enrichers-exceldna) is a Serilog Enricher with properties from Excel-DNA Add-Ins

## External Links

If you are a VBA developer interested in moving to .NET, you should start with [Patrick O'Beirne's detailed VBA to Excel-DNA migration guide](http://sysmod.wordpress.com/2012/11/06/migrating-an-excel-vba-add-in-to-a-vb-net-xll-with-excel-dna-update/).

### Various Samples and Tutorials
- [Hugo Diz - TextUtilsDNA](https://github.com/hugodiz/TextUtilsDNA) is a useful text processing add-in, and also a nice example add-in created with VB.NET.
- As a comprehensive example using many of the Excel-DNA features, be inspired by the [Financial Analytics Suite (FinAnSu)](http://brymck.github.com/finansu/), an open-source C# add-in built by Bryan McKelvey.
- [Ross McLean](https://web.archive.org/web/20140902002824/http://www.blog.methodsinexcel.co.uk/2010/08/16/why-excel-dna/) has a series of posts on getting started with Excel-DNA.
- [Mikael Katajamäki shows how to use Microsoft Solver Foundation to build a curve fitting function for Excel](http://mikejuniperhill.blogspot.com/2013/06/using-ms-solver-foundation-and-c-in.html)
- [Mikael Katajamäki shows how to use C++/CLI code as a wrapper class for native (Quantlib based) C++ code and interfaced the C# client code to Excel by using Excel-DNA](http://mikejuniperhill.blogspot.com/2018/10/wilmott-software-interoperability-in.html)
- [Simon Murphy - xlls with Excel-DNA](http://smurfonspreadsheets.wordpress.com/2010/02/18/xlls-with-exceldna/)
- [Ed Parcell - Numerical analysis in Excel using C# with Excel-DNA and AlgLib](https://web.archive.org/web/20100511213800/http://edparcell.posterous.com/tag/excel)
- [Mathias Brandewinder - Mutant Excel with .NET and Excel-DNA](http://www.clear-lines.com/blog/post/Mutant-Excel-and-Net-with-ExcelDNA.aspx)
- [Mathias Brandewinder - Supercharge Excel functions with Excel-DNA and .NET parallelism](http://www.clear-lines.com/blog/post/Supercharge-Excel-functions-with-ExcelDNA-and-Net-parallelism.aspx)
- [Mike Woodhouse - A third way: DNA?](http://grumpyop.wordpress.com/2009/11/25/a-third-way-dna/)
- [Patrick O'Beirne - From VBA to VB.NET using Excel-DNA](http://sysmod.wordpress.com/2012/02/06/from-vba-to-vb-net-using-exceldna/)
- [Doctor Torsten - Bring Excel 2010 to Speed: Remote UDFs with Excel 2010 and HPC Server 2008 R2](http://web.archive.org/web/20140831133544/http://doctortorsten.wordpress.com/2011/01/10/remoteudfs/)
- [Luca Bolognese - A trading/portfolio management Excel Add-in based on the books by Ralph Vince](https://www.lucabol.com/posts/2007-01-04-a-tradingportfolio-management-excel-add-in-based-on-the-books-by-ralph-vince/)
- [Supermab's series of blog posts introducing Excel-DNA to Japan (in Japanese)](http://supermab.com/wp/category/excel/)
- [Joao Morais - WCF client sample](http://blog.ilab8.com/2011/01/28/excel-dna/)
- [teramonagi - Using R from Excel using Excel-DNA](http://mockquant.blogspot.com/2011/07/another-way-to-use-r-in-excel-for-net.html) (Also check out the [F# R type provider.](https://bluemountaincapital.github.io/FSharpRProvider/))
- [Gert-Jan van der Kamp - Streaming real-time data to Excel](http://www.codeproject.com/Articles/662009/Streaming-realtime-data-to-Excel)
- [Eddie Gahan - TransposeBy - Extending Excel With C# And Excel-DNA](https://www.c-sharpcorner.com/article/transposeby-extending-excel-with-c-sharp-and-excel-dna/)

### Various F# Examples
- a [wrapper for the ExcelReference type and C API](https://gist.github.com/mndrake/5963107) with some cell access features,
- an [object handle wrapper](https://github.com/mndrake/ExcelObjectHandler), and
- a [WPF-based Custom Task Pane](https://github.com/mndrake/ExcelCustomTaskPane) for Excel.
- [Three samples, including a function using the R type provider](https://web.archive.org/web/20171228052128/http://luajalla.azurewebsites.net/excel-dna-three-stories/) by Natallie Baikevich.
- [Bram Jochems](https://web.archive.org/web/20140403050217/http://bramjochems.com/blog/2013/10/example-f-exceldna/) has published a wonderful [collection of finance-related functions on GitHub](https://github.com/bramjochems/MyExcelLib), as well as some details on [creating a ribbon menu with F#](https://web.archive.org/web/20160714194609/http://bramjochems.com/blog/2014/05/creating-ribbon-menu-exceldna-f/).
- Useful Range wrappers by Kit Eason: [Higher-Order Functions for Excel](http://www.fssnip.net/aV).

### External projects using Excel-DNA
- [Dodoni.net is a free/open-source library for quantitative finance and numerical computing.](https://dodoni.github.io/).
- [Cubicle Tools](https://cubicle.codeplex.com) is a collection of tools that extends Excel for analytical and rapid development purposes. It includes an object handler and an add-in distribution system.
- [EQ Finance - Analytics library for derivatives pricing and risk management](http://www.eqfltd.com/software)
- [Technoscience UK](http://excelxll.com/) has some interesting add-ins to mirror Excel data between PCs.
- [Niels Bosma -  SEOTools add-in (free, but not open source)](http://nielsbosma.se/projects/seotools/)
- [compute!box!](http://web.archive.org/web/20130616043202/http://computebox.wordpress.com/) allows real-time interchange of data between spreadsheets (via Azure Service Bus).
- This [Office icon gallery](https://archive.codeplex.com/?p=imagemso) has an Excel-based viewer.
- [Jon Nyman's FxToExcel add-in](https://github.com/jon49/FxToExcel) brings financial program data into Excel.
- [Stock Quote Add-In for Excel](https://github.com/jbaurle/PMStockQuote) provides access to the Yahoo financial data through an Excel-DNA add-in.
- [DB-Addin for Excel](https://rkapl123.github.io/DBAddin/) is an MS Excel Addin for retrieving Database data via userdefined functions into Excel and writing Data (DBMapper), executing generic DML (DBAction) and doing all this in Sequences (DBSequence).
- [Datepicker](https://github.com/rkapl123/DatePicker) is a replacement for the MSCOMCT2 based Datepicker that Microsoft abandoned in 64bit versions of Excel. It passes the .NET MonthCalendar widget to VBA Userforms.
- Alex Chirokov's **ACQ** add-in provides a library of interpolation routines for Excel. The add-in includes 1D and 2D interpolators, scatter plot smoothing and a Mersenne Twister random number generator. To have a closer look:
    - Find the current release on GitHub: [https://github.com/ratesquant/ACQ/releases][acq-releases]
    - With the main repository on GitHub at [https://github.com/ratesquant/ACQ][acq-repo]
    - A very clear introduction to the library, including some of it's advanced features, is posted on Code Project: [http://www.codeproject.com/Articles/1097174/Interpolation-in-Excel-using-Excel-DNA][interpolation-article]
Features I like about the add-in (apart from it using Excel-DNA) include:
    - A liberal open-source license
    - A clear and authoritative implementation of a particular domain
    - Very nice example of using object handles - an interpolator is build from the data, and then used to interpolate many values. ACQ has a clean implementation and great example of this technique.
    - All the functions have a common prefix ("`=acq`..."), making them easy to find in the function list, and use with the Excel-DNA IntelliSense extension.
    - PS: ACQ has a bonus feature that implements a Sudoku solver (and generator)! See the write-up here: [Sudoku Solver in Excel using C# and Excel-DNA][sudoku-solver-article].
- I noticed a very nice add-in developed by [Bryan McKelvey][bryan-mckelvey] called [FinAnSu][finansu]. The whole add-in is generously available under the MIT open source license, and is a fantastic example of what can be built with Excel-DNA.
    - [FinAnSu][finansu] uses a ribbon interface to make the various functions and macros easy to find. The RTD server support is used to implement asynchronous data update functions, providing a live quote feed from Bloomberg, Google or Yahoo! And then there is a bunch of useful-looking financial functions. Here's a little preview:

        ![FinAnSu Quote Animated][finansu-quote-img]

    * Find the project on GitHub: [https://github.com/brymck/finansu][finansu], with detailed [documentation][finansu-docs].
    * You can browse through the [source code][finansu] online, and you can also download a copy of the whole project.

### Support

There is a searchable record of more than 5000 messages on the [Excel-DNA Google Group](https://groups.google.com/group/exceldna).

There are many questions answered on Stack Overflow under the tag [`excel-dna`](http://stackoverflow.com/questions/tagged/excel-dna).

**Please don't hesitate to ask.** If you are stuck or need some help using Excel-DNA your questions really are very welcome - whether you are just getting started, or an Excel-DNA expert.

And if you could help put together some proper documentation, please contact me. We'd be happy to add you as an editor in this repository.

## Related Projects
- [NetOffice](https://github.com/netoffice) is a set of version-independent assemblies to allow Office integration through the COM automation interface. The NetOffice libraries can be used from an Excel-DNA add-in to ease version-independent Excel add-in development, and ease compatibility with VBA.
- Visual Studio Tools for Office (VSTO) is Microsoft's preferred plan for integrating .NET with Office. It is mainly aimed at making it easy for Visual Studio developers to create solutions integrated with the Office applications. In contrast, Excel-DNA is (eventually) aimed at Excel end-users, as a compelling replacement for VBA, completely independent of Visual Studio.
- [Add-in Express](http://www.add-in-express.com) is a commercial alternative to VSTO for users with Visual Studio. It support making add-ins for the various Office products, not just Excel, and has helpful wizards and graphics designers.
- Jens Thiel's ManagedXll was an established, commercial product to easily create .xll libraries in .NET. If ManagedXll were free, Excel-DNA would not exist.
- [Statfactory's NeXL](https://statfactory.wordpress.com) are F# based connectors to get data from various platforms (Bloomberg, Quandl, Worldbank, IMF and the R language) into Excel.
- For making Excel Add-Ins in Python, have a look at [PyXLL](http://www.pyxll.com).
- There are a number of C/C++ libraries and tools that make creating .xlls easier than using the [Excel SDK](https://docs.microsoft.com/en-us/office/client-developer/excel/welcome-to-the-excel-software-development-kit) directly:
    * I initially used the [XLW](http://xlw.sourceforge.net/) open-source library.
    * The [XLL+ toolkit](https://www.planatechsolutions.com/xllplus/) is a highly regarded commercial option.
    * Keith Lewis has some modern C++ libraries for making .xlls, available on CodePlex at [https://archive.codeplex.com/?p=xll](https://archive.codeplex.com/?p=xll).

## Performance
Information about the performance of Excel-DNA user-defined functions can be found on the [Excel-DNA Performance](exceldna-performance.md) page.

## Formal Support Agreements
Corporate users of Excel-DNA, using the library as part of their mission critical infrastructure, are encouraged to enter into a formal support arrangement. We offer an annual subscription-based technical support agreement, providing direct support, priority bug-fixes and feature development and ensuring that Excel-DNA will continue to be updated and developed.

## Donations
Financial support for the Excel-DNA project encourages future development and is greatly appreciated. Transactions are processed by PayPal.
[![Donate via PayPal][paypal-image]][paypal-link]

## More Details
Excel-DNA is developed using .NET, and users have to install the freely available .NET Framework runtime. The integration is by an Excel Add-In (.xll) that exposes .NET code to Excel. The user code can be in text-based (.dna) script files (C#, Visual Basic or F#), or compiled .NET libraries (.dll). Excel-DNA supports both the .NET runtime version 2.0 (which is used by .NET versions 2.0, 3.0 and 3.5) and version 4. Add-ins can target either version of the runtime, and concurrent loading of both runtime versions into an Excel instance is supported.

Excel versions '97 through 2016 can be targeted with a single add-in. Advanced Excel features are supported, including multi-threaded recalculation (Excel 2007 and later), registration-free RTD servers (Excel 2002 and later) and customized Ribbon interfaces (Excel 2007 and 2010). There is support for integrated Custom Task Panes (Excel 2007 and later), offloading UDF computations to a Windows HPC cluster (Excel 2010 and later), and for the 64-bit versions of Excel 2010 and 2013.

Most managed UDF assemblies developed for Excel Services can be exposed to the Excel client with no modification. (Please contact me if you are interested in this feature.)

Since Excel-DNA uses the Excel C API, porting C/C++ add-in code based on the Excel XLL SDK is very easy. (No more `XLOPER`s!)

The Excel-DNA Runtime is free for all use, and distributed under a permissive open-source license that also allows commercial use.

Originally, the project was hosted on [https://exceldna.codeplex.com](https://exceldna.codeplex.com), where you can still download the site in it's historic state as a package. After CodePlex' shutdown the archive site is however mostly unusable by now.

## Latest Releases
The current version is [Excel-DNA v1.1], released in June 2020 and includes numerous improvements and bug-fixes.

[Excel-DNA v1.1]: https://excel-dna.net/2020/06/29/excel-dna-version-1-1/
[paypal-link]: https://www.paypal.com/cgi-bin/webscr?cmd=_donations&amp;business=92N99RV5NQ29C&amp;lc=US&amp;item_name=Govert%20van%20Drimmelen&amp;item_number=ExcelDna&amp;currency_code=USD&amp;bn=PP%2dDonationsBF%3abtn_donate_LG%2egif%3aNonHosted
[paypal-image]: https://www.paypal.com/en_GB/i/btn/btn_donateCC_LG.gif "Donate via PayPal"
[acq-releases]: https://github.com/ratesquant/ACQ/releases
[acq-repo]: https://github.com/ratesquant/ACQ
[interpolation-article]: http://www.codeproject.com/Articles/1097174/Interpolation-in-Excel-using-Excel-DNA
[sudoku-solver-article]: http://www.codeproject.com/Articles/1098156/Sudoku-Solver-in-Excel-using-Csharp-and-Excel-DNA
[bryan-mckelvey]: https://github.com/brymck
[finansu]: https://github.com/brymck/finansu
[finansu-quote-img]: /assets/finansu-quote-animated.gif "FinAnSu Quote Animated"
[finansu-docs]: http://brymck.github.com/finansu/


# Appendix

## Why Excel-DNA?

Microsoft publishes a Software Development Kit (SDK) for Excel, that describes how to make an 'Excel add-in'. These are similar to .xla add-ins, but the code would typically be written in C or C++, and is compiled to binary file with a .xll extension. Such an Excel add-in is typically called an ".xll". Internally, it is just a normal Windows .dll file with a few pre-arranged exports so that Excel and the add-in can hook up.

Xll add-in have some advantages over .xla add-ins developed in VBA. They can define user-defined worksheet functions (UDFs) which run very fast and which can integrate with newer features like multi-threaded calculation in Excel 2007+, and asynchronous calculation in Excel 2010+. A disadvantage of xll add-ins is that they are hard to develop. Typically C or C++ is used, and there are some toolkits and a nice book to help you, but it's still hard.

The .NET Framework (normally just called .NET) is a (twelve-year old) software framework from Microsoft, including the (Java-like) C# language and an updated version of Visual Basic, extensive standard libraries, a runtime environment (libraries that manage execution of your code) and an intermediate 'assembly' language with corresponding just-in-time compiler. The .NET languages and runtime environment is often called 'managed'. So a .NET library would be called a 'managed' library, as opposed to a library compiled from C / C++, which would be a 'native' library. Wikipedia can tell you a lot more about .NET. It has become the standard development environment for corporate software development on the Windows platform.

The Microsoft development tool (giving you the compilers and Integrated Development Environment (IDE)) associated with .NET is called Visual Studio. What you might think of as the 'real' Visual Basic (the last version was VB6) was upgraded to become a language as part of .NET, often called VB.NET, and Visual Studio is the standard IDE for developing VB.NET applications. There are free editions of Visual Studio, and then a range of paid for editions with more and more features.

There is some support in Visual Studio for making Office add-ins using .NET, with a library called Visual Studio Tools for Office (VSTO). However, initially (ten years ago) VSTO had many complications with deployment, and particular for Excel has serious limitations - that UDFs could not be created.

So for Excel there was a problem - how to use .NET to create full-featured and high-performance add-ins (meaning .xll add-ins). There was a commercial solution called ManagedXLL when I looked around in 2004, but it was too expensive to be useful to me. So I started an open-source project called Excel-DNA (the 'DNA' stands for DotNet for Applications, as opposed to Visual Basic for Applications).

The main Excel-DNA sites are [http://excel-dna.net](http://excel-dna.net) and [https://docs.excel-dna.net](https://docs.excel-dna.net).

Now (after ten years) Excel-DNA is mature and widely used as the standard .NET to Excel integration tool. But in practice, it is most useful for developers already using .NET and Visual Studio who want to make high-performance Excel add-ins.

To be sure, there are other ways of making .xll add-ins, including a tool called PyXll that is similar to Excel-DNA, but for the Python language rather than the .NET Framework.

I hope in the next few years to make it easier for Excel VBA users to upgrade to using VB.NET and Excel-DNA. It has proven harder than I expected, as there is a steep learning curve, and the advantages are uneven and very dependent on the programmer's background.

One guide, written from a VBA user's perspective, that you might look at to see whether you might be interested is a porting guide written by Patrick O'Beirne: [http://sysmod.wordpress.com/2012/11/06/migrating-an-excel-vba-add-in-to-a-vb-net-xll-with-excel-dna-update/](http://sysmod.wordpress.com/2012/11/06/migrating-an-excel-vba-add-in-to-a-vb-net-xll-with-excel-dna-update/)

So if you have existing VBA code (mainly worksheet functions) that you want to run fast in newer versions of Excel on modern hardware, Excel-DNA is a good tool. You'd port your VBA code to VB.NET (using the free version of Visual Studio) and use Excel-DNA to glue that back to Excel. I'm more than happy to help users learning how to do this on the Google group (or sometimes directly), but there is a steep and long learning curve once you put your foot outside the familiarity of VBA.

We believe it's worth the effort.