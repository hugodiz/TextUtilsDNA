# TextUtilsDNA
High-performance text wrangling and fuzzy lookup functions for Excel, powered by .NET via ExcelDNA

## Introduction
These hopefully useful, general-purpose Excel worksheet functions for text processing and fuzzy matching are written in VB.NET and plugged-in to Excel as an **xll** (C API).

These functions rely entirely on ExcelDNA (by Govert van Drimmelen) in order for them to be exposed as Excel functions.

The UNPACK utility function also uses the extremely popular .NET library NewtonSoft.Json (by James Newton-King):  
https://www.newtonsoft.com/json  
Here I use it for deserialization of Excel arrays encoded as JSON strings.

I have attempted to optimize these functions, including using parallelization in a way analogous to what Excel does where it can, when using built-in functions. For more details on the approach, see "ExcelFunctions - NOTES.md" doc  **=>**  "NOTE about parallelization of loops in these functions".

The ability to create .NET-powered functions such as these and then exposing those functions to Excel worksheets is exactly the type of thing that is made dramatically easier, more tracktable and more seamless using the excellent ExcelDNA open-source project.

This code is itself open-source (MIT license) and these functions, whilst (hopefully) useful on their own, are also meant as a small contribution to showcase the ExcelDNA toolset to experienced Excel users and programmers who may at times either feel limited by VBA, or tend to build extremely complex programs in VBA which would be better suited for .NET.

Hence, depending on your project's size, performance and interoperability needs, VB.NET might be a much better choice than VBA.  
As of 2021, ExcelDNA is one of the best ways to bring the power of .NET (C#/VB.NET/F#) to Excel. If this is new to you please visit:  
https://docs.excel-dna.net/what-and-why-an-introduction-to-net-and-excel-dna/  
as a starting point.

These functions are ideally meant to be used with Excel 365, because they levarage the power of dynamic arrays.  
However, calling LSDLOOKUP with a scalar lookup_value and K = 1 should work without problems in most Excel versions. Basically, you should be fine whenever one of these functions would return a scalar anyway.

Without dynamic arrays in your Excel version, I believe you will need to pre-select a range of the right size, then use a TextUtilsDNA function normally, but trigger it with ctrl + shift + Enter instead of just Enter. Otherwise, Excel might just show you the upper-left corner of the result instead of the whole (array) result.  

At the end of the day, if you can, you really should be using Excel 365, it's worth it.

ExcelDNA will automatically produce 32 and 64-bit versions of the **xll** if you build the project in Visual Studio - you'll then use the appropriate one for your system (meaning, check the *bitness* of your Excel version). The above link to "ExcelDna - What and why" does a very good job of explaining what an **xll** is and how it relates to the other types of Excel add-ins available.

## Getting Started
Documentation work in progress - but it's going to be pretty much the same process one would go through with any other **xll** C API Excel add-in, in general, and any other ExcelDNA add-in, in particular.

As a helpful reference by analogy, a very similar structure can be seen in the following github link, because those too are Excel functions made in .NET and exposed to Excel using ExcelDNA:
https://github.com/Excel-DNA/XFunctions

Binary releases of TextUtilsDNA are hosted on GitHub: (to do: set up releases and link)  
In principle, downloading a copy of either the 32 or 64-bit TextUtilsDNA **xll** binary and having Excel ready go on your end, then adding the **xll** as an "Excel addin" in the Developer tab, is all one should need to do in order to get the functions up and running

However, I will also include a consolidated step-by-step approach, should you wish to build this project from scratch using Visual Studio but not at all be already familiar with that environment. This section will be written primarily for those who are quite at home with Excel and VBA and would like to expand that familiarity into the .NET/VisualStudio/ExcelDNA toolset. However, my instructions here won't preclude the need (or at least the very strong recommendation) that you have a look at the series of excellent YouTube tutorials by Govert on getting started with coding .NET functions for Excel via ExcelDNA:   
https://www.youtube.com/user/govertvd

## Examples
Documentation work in progress - in the meantime, I've tried to make the Excel IntelliSense auto-complete help as comprehensive as possible. I'll complement that with usage examples here.

## Support and participation
Any help or feedback is greatly appreciated, including ideas and coding efforts to fix, improve or expand this suite of functions, as well as any efforts of testing and probing, to make sure the functions are indeed 100% bug-free.

Please log bugs and feature suggestions on the GitHub 'Issues' page.   

Note in particular that, since this is all in an early stage, I expect we may find a few bugs in these functions. I expect that if one does find a bug, it will likely manifest in one of three likely ways:

1. The function returns #VALUE! when, by all accounts, it should be returning an actual, non-error value   
2. The function returns an incorrect value, a value not in line with what the function spec says it should return, given the inputs   
3. Calculating the function makes Excel crash (just closes without warning or message) - see how a bug might cause this to happen in the "ExcelFunctions - NOTES.md" doc  **=>**  "NOTE about ExceptionSafe and ThreadSafe functions".

## License
The TextUtilsDNA VB.NET functions are published under the standard MIT license, with the associated Excel integration relying on ExcelDNA (Zlib License):   
https://excel-dna.net/   
https://github.com/Excel-DNA/ExcelDna

Hugo Diz

hugodiz@gmail.com

11 April 2021

