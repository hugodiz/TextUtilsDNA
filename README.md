# TextUtilsDNA
High-performance text wrangling and fuzzy lookup functions for Excel, powered by .NET via ExcelDNA

**LSDLOOKUP:** takes a column of lookup_values and retrieves the K closest matches to each lookup_value, as found in a lookup_array, where "closest" means "least typos", and and the "number of typos" is basically the Levenshtein distance between 2 text strings (check out the wikipedia page for Levenshtein Distance if unfamiliar). Searches may be narrowed down in 3 possible ways simultaneously: 
- by defining a maximum allowed Levenshtein Distance in matches
- by requiring matches to exhibit a regular expression pattern P (a P positive filter, if you will)
- by requiring matches to NOT exhibit a regular expression pattern Q (a negative Q filter, if you wil)

**UNPACK:** because LSDLOOKUP can optionally give back a match's coordinates in the lookup_array instead of the text itself, and because lookup_array is allowed to be 2D, we need a way to represent arrays in single cells (in this case containing a tuple [row index, column index]) - the convention we'll use is JSON. This function takes a JSON string representation of any 1D or 2D array(ie. [A(1), A(2), ..] or [[A(1,1), A(1,2), ..], [A(2,1), A(2,2), ..],..]) and produices the actual array. Note that 1D arrays are single rows (not columns) by convention.

**TEXTSPLIT:** the inverse of the built-in Excel function TEXTJOIN. TEXTSPLIT takes a single (scalar input) string and returns a row containing each piece of the string, resulting from splitting the string according to a delimiter.

**RESUB:** takes an input array (2D allowed) and returns a similarly-sized array, where each entry has undergone a regular expression transformation (similar to .NET Regex.Replace). A regular expression pattern P is specified, as well as a replacement string R. All occurrences of P in each input are replaced by R. Although R must be a literal string, it may include the usual $**G** methodology (for re-using pieces of the pattern), where **G** is an integer number representing the **G**th captured group, if specified in P.

**REGET:** Similar idea as RESUB, except here we *extract* the **i**th occurrence of the pattern P in each input (or optionally only the **j**th captured group of the **i**th occurence) and simply return the extracted bit from each input.

**RECOUNTIF:** Recycles the P and Q filter ideas from LSDLOOKUP, and here we conditionally count the entries from an input range which simultaneously exhibit the specified regular expression pattern P (if specified) and do NOT exhibit the regular expresion pattern Q (if specified). So it is like the built-in COUNTIF but for regular expressions.

**GETCOUNTS:** Counts the number of occurrences of each distinct text string in an input range (i.e. it builds the hit counts of each word into a Dictionary(string : integer), which is ultimately output as a range with 2 columns for (Keys, Values), and as many rows as unique words. The output is not sorted in any particular order, because GETCOUNTS can be easily composed with the Excel built-in SORT in order to conform to any desired sorting. Also, if the *case-sensitive* flag is set to false, the output deliberately shows all words converted to uppercase to remind us of the fact that it did not care about Case when counting.

## Introduction
These hopefully useful, general-purpose Excel worksheet functions for text processing and fuzzy matching are written in VB.NET and plugged-in to Excel as an **xll** add-in.

These functions rely entirely on ExcelDNA (by Govert van Drimmelen) in order for them to be exposed as Excel functions.

The UNPACK utility function also uses the extremely popular .NET library NewtonSoft.Json (by James Newton-King):  
https://www.newtonsoft.com/json  
Here I use it for deserialization of Excel arrays encoded as JSON strings.

I have attempted to optimize these functions, including using parallelization in a way analogous to what Excel does where it can, when using built-in functions. For more details on the approach, see [ExcelFunctions - NOTES.md](https://github.com/hugodiz/TextUtilsDNA/blob/main/Source/TextUtilsDNA/ExcelFunctions/ExcelFunctions%20-%20NOTES.md)  **=>**  "Note about parallelization of loops in these functions".

Also, for those who enjoy thinking about algorithms in detail, I've commented the LSDLOOKUP code extensively. Although the fundamental techniques are well-known and rather standard (ie. Levenshtein distance calculation and a binary max-heap-based priority queue), I have tried to adapt these techniques and tailor them to purpose so that each piece would be minimal, economic and attach seamlessly to the overall LSDLOOKUP function. I have enjoyed the process thoroughly and tried to document it well. In addition to the code itself, I'd invite you to have a look at [MaxHeapMechanics - NOTES.md](https://github.com/hugodiz/TextUtilsDNA/blob/main/Source/TextUtilsDNA/MaxHeapMechanics%20-%20NOTES.md) and also the first section of [ExcelFunctions - NOTES.md](https://github.com/hugodiz/TextUtilsDNA/blob/main/Source/TextUtilsDNA/ExcelFunctions/ExcelFunctions%20-%20NOTES.md)

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

ExcelDNA will automatically produce 32 and 64-bit versions of the **xll** if you build the project in Visual Studio - you'll then use the appropriate one for your system (meaning, check the *bitness* of your Excel version). The above link to "ExcelDna - What and why" does a very good job of explaining what an **xll** is and how it relates to the other types of Excel add-ins available. From the end-user's point of view, an **xll** add-in is just something to be *added* to Excel, in a similar fashion to how you'd *add* a **xla** or **xlam** add-in.

## Getting Started
Documentation work in progress - but it's going to be pretty much the same process one would go through with any other **xll** add-in, in general, and any other ExcelDNA add-in, in particular.

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
3. Calculating the function makes Excel crash (just closes without warning or message) - see how a bug might cause this to happen in [ExcelFunctions - NOTES.md](https://github.com/hugodiz/TextUtilsDNA/blob/main/Source/TextUtilsDNA/ExcelFunctions/ExcelFunctions%20-%20NOTES.md)  **=>**  "Note about ExceptionSafe and ThreadSafe functions".

## License
The TextUtilsDNA VB.NET functions are published under the standard MIT license, with the associated Excel integration relying on ExcelDNA (Zlib License):   
https://excel-dna.net/   
https://github.com/Excel-DNA/ExcelDna

Hugo Diz

hugodiz@gmail.com

11 April 2021

