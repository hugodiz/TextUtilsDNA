# TextUtilsDNA
High-performance text wrangling and fuzzy lookup functions for Excel, powered by .NET via ExcelDNA

## Introduction
These hopefully useful, general-purpose Excel worksheet functions for text processing and fuzzy matching are written in VB.NET and plugged-in to Excel as an XLL (C API).

These functions rely entirely on ExcelDNA (by Govert van Drimmelen) in order for them to be exposed as Excel functions.

The UNPACK utility function also uses the extremely popular .NET library NewtonSoft.Json (by James Newton-King)
https://www.newtonsoft.com/json
Here I use it for JSON deserialization of arrays

I have attempted to optimize these functions, including using parallelization in a way akin to what Excel would do at times.
The ability to create .NET-powered functions such as these and then exposing those functions to Excel worksheets
    is exactly the type of thing that is made dramatically easier, more tracktable and more seamless 
    using the excellent ExcelDNA open-source project

This code is itself open-source (MIT license) and these functions, whilst (hopefully) useful on their own, are also meant as 
    a small contribution to showcase the ExcelDNA toolset to experienced Excel users and programmers who may at times 
    either feel limited by VBA, or tend to build extremely complex programs in VBA which would be better suited for .NET.
So, depending on your project's size, performance and interoperability needs, VB.NET might be a much better choice than VBA
As of 2021, ExcelDNA is one of the best ways to bring the power of .NET (C#/VB.NET/F#) to Excel. If this is new to you
    please visit ' https://docs.excel-dna.net/what-and-why-an-introduction-to-net-and-excel-dna/ as a starting point

These functions are ideally meant to be used with Excel 365, because they they levarage the power of dynamic arrays.
However, calling LSDLOOKUP with a scalar lookup_value and K = 1 should work without problems in most Excel versions.
    Basically, you should be fine whenever a function here wold return a scalar.
    Without dynamic arrays in your Excel version, I believe you will need to pre-select a range of the right size,
        then use a TextUtilsDNA function normally, but trigger it with ctrl + shift + Enter instead of just Enter
        Otherwise, Excel might just show you the upper-left corner of the result instead of the whole (array) result
    At the end of the day, if you can, you really should be using Excel 365, it's worth it.

ExcelDNA will automatically produce 32 and 64-bit versions of the XLL - you'll then use the appropriate one for your system

## Getting Started

Documentation work in progress

## Examples

Documentation work in progress

