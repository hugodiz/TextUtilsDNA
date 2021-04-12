# TextUtilDNA.ExcelFunctions notes: 
by Hugo Diz, 2021

The algorithms used in these functions (most notably LSDLOOKUP) are slight variations of well-known techniques and structures deliberately implemented from scratch and tailored-to-purpose, for pedagogical purposes and to maximize performance.

The LSDLOOKUP search uses a minimalistic implementation of the Levenshtein distance metric with a "capped" algorithm, ie.:   
- it early-discards the dynammic programming calculations where the Lev dist is already too high, given the context   
- it only computes a minimal matrix diagonal strip required to either compute the distance or conclude it's too high

The best K results "thus far" are kept in a binary max-heap serving as a priority queue which only holds K items at a time, such that:   
- K is specified by the user
- Higher priority means "worse match", and these nodes remain closer to the root (the root is the Kth worst match)
- Upon enqueing new items, the "worst" matches beyond K are automatically dequeued:
    - In other words, since the heap only holds K items, it automatically dequeues ("serves") the highest priority item to make room for a new item. Each such dequeued item is in this case discarded (ie. didn't make the best K)
- At the end of the function, the items are dequeued in turn, "worst match first", until the heap is empty, except now the dequeued items are placed in the Output, as they represent the best K matches.

## Note about ExceptionSafe and ThreadSafe functions:
The functions marked with ExceptionSafe and ThreadSafe are such because they should not to be able to encounter unhandled exceptions whilst in use. As per the ExcelDNA documentation, by specifying the function as ExceptionSafe and ThreadSafe it speeds it up considerably.

The trade-off is that an automatic, back-end layer of function "baby-sitting" is removed and so if this function were to actually encounter an unhandled exception whilst in usage, Excel would crash instead of just catching the exception and popping a notifying dialog box explaining the error.

So, if you can get Excel to systematically crash when doing something with these functions, then you've probably found a bug, please let me know :)

## Note about parallelization of loops in these functions:
Excel parallelizes worksheet calculations by default where it can, say for instance when dragging down a VLOOKUP of single lookup_value inputs, over an very large column of lookup_value input cells.

Some TextUtilsDNA functions (most notably LSDLOOKUP) implement some Parallel For loops instead of sequential ones. These parallel For's, when used here, are always trivial in that they're ideally paralellizable, namely:   
- They're used when the function allows the main function input to be a range of independent inputs
- In LSDLOOKUP, that main input is a column A(Hx1) of lookup_value(s), each element being A(h)
- Each parallel thread will only do work on a disjunct slice of A(start..end) and writes to a correspondingly disjunct slice of FnOutput(start..end, 1..K)
- There are no data races because no thread ever attempts to write anywhere another thread might write to
- Also, Read/Write operations to the Excel Session Caches are safe because we're using .NET Concurrent Dicts, and because even if bad timing were to cause a thread reading from the cache to "miss" a value which technically had just been written onto the cache by another thread, such a "miss", although regrettable and in theory not mandatory, is just business as usual when using caches: nothing terrible would come out of it because cache reads are always *tentative* by definition
- The function will split the workloads uniformly over all available cores using 1 Thread per Core (like Excel does by default).

The main point is that our deliberate .NET parallelization kicks-in precisely when Excel normal parallelization wouldn't: when calling these functions just once but with multi-cell A(Hx1) - It is essential to explicitly parallelize here.

But why then not just restrict A to be a scalar (eg. LSDLOOKUP a single lookup_value and let Excel be parallel when dragging?

Because we always pay the cost of reading any other inputs from Excel to .NET (eg. lookup_array) once per function call. Hence, dragging down, whilst convenient, can potentially add a lot of overhead redundant work when there are big secondary input ranges, such as the lookup_array in LSDLOOKUP, which is just repeatedly being read into .NET over and over again, once per each "dragged-down" instance of the formula.

As a result, calling LSDLOOKUP once with a large A(Hx1) lookup_values range will be much, much faster than calling LSDLOOKUP with a single lookup_value and dragging the formula down (assuming lookup_array B is also very large). However, this "being much faster" is only true so long as we do indeed parallelize (otherwise, dragging would probably end up being better).
