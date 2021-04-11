# ExcelIntelliSense notes: 
(using ExcelDna.IntelliSense by Govert van Drimmelen)

This class contains code for embedding Excel auto-complete intellisense into the TextUtilsDNA xll add-in, using ExcelDNA. The usage of this feature is STRONGLY recommended because I have included a lot of function usage help through this medium. The contents of the ExcelIntelliSense help are those present in the function signatures in the actual code. 

This style of IntelliSense mimicks the behaviour of inbuilt Excel formulae and is very useful.

This implementation is what the ExcelDNA documentation refers to as 'Integrated mode' of IntelliSense (as opposed to 'Standalone mode'). In integrated mode, the TextUtilsDNA-AddIn.dna file contains a reference that "packs" ExcelDna.IntelliSense.dll into the TextUtilsDNA.xll final add-in file, as per the code below. Hence, all is contained is TextUtilsDNA.xll

However, if you already have or prefer to have ExcelIntellisense running as an standalone add-In In your Excel, then instructions on doing this are available at    https://github.com/Excel-DNA/IntelliSense/releases   
If using the standalone mode, then you won't need this code file (ExcelIntelliSense.vb) at all, because the standalone IntelliSense add-In will activated independently.

However, it may be simplest to simply use this embedded lightweight boiler-plate as shown here, which keeps the Excel setup hassle To a minimum. In fact, ExcelDna.IntelliSense is not the only thing being "packed" - NewtonSoft.Json also is, so that NewtonSoft.Json.dll does not then need to seperately follow the XLL file everywhere.
