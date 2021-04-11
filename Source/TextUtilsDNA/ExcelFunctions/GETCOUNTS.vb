Option Explicit On
Option Strict On

Imports ExcelDna.Integration

Namespace TextUtilsDna
    Public Module Excel_GETCOUNTS

        ' *********************************************************************************************************************
        ' GETCOUNTS Spec : Counts the number of occurrences of each distinct word in an input range
        ' ---------------------------------------------------------------------------------------------------------------------
        ' This function essentially computes a {String : Integer} dictionary of {Word : Counts for this Word}
        '     for each distinct word in an input range
        ' If Not is_case_sensitive, then all strings in the input range are converted to upper case and the output shows
        '     all words as full upper-case, to make it obvious that Case was not taken into account during the counting.
        ' The output is a 2-column range where each row is Word X : WordCounts(X)
        ' The output is is NO particular order, because the TEXTCOUNTS output can then be easily composed
        '     with the Excel SORT function
        ' EG.: =SORT(TEXTCOUNTS(A),1) will show the output sorted by ascending alphabetical order (ie. top -> bottom = a -> z )
        ' EG.: =SORT(TEXTCOUNTS(A),2,-1) will show the output sorted descending frequency order (ie. "greatest hits" on top) 
        '
        ' This function is ExceptionSafe and ThreadSafe (read note at the top of this file)
        ' *********************************************************************************************************************
        ' GETCOUNTS Function Signature
        ' ---------------------------------------------------------------------------------------------------------------------
        <ExcelFunction(
            IsExceptionSafe:=True,
            IsThreadSafe:=True,
            Description:="Gets the number of occurrences of each text string in an input range")>
        Function GETCOUNTS(
            <ExcelArgument(Name:="text(s)",
                           Description:="<[Range]> Range of Text string(s) A to analyze:
A(HxW) : each A(h,w) = value as seen in Excel under ""General"" formatting
A Omitted : Output = #N/A
A(h,w) = """" Or Empty Or Error : A(h,w) excluded
Hint: =SORT(TEXTCOUNTS(A),[1 or 2], [1 or -1])")>
            texts As Object(,),
 _
            <ExcelArgument(Name:="[is_case_sensitive]",
                           Description:="<[SCALAR]> Defaults to FALSE if Omitted or Empty")>
            is_case_sensitive As Boolean) As Object
            ' *****************************************************************************************************************
            ' GETCOUNTS Function Implementation 
            ' -----------------------------------------------------------------------------------------------------------------
            ' Notes: "Fn" control variables are defined and scoped to the main function call, and then used from within loops
            '        snake_case variables are Excel function "raw" inputs - typically these will be processed and be used to
            '            define internal control variables (such as the "Fn" variables)

            If TypeOf texts(0, 0) Is ExcelMissing Then Return ExcelError.ExcelErrorNA

            Dim FnFreqDict = New Dictionary(Of String, Integer)

            Dim h As Integer
            Dim ThisText As String

            ' I chose not to parallelize here for two main reasons:
            '     The unitary workload (adding 1 to the total count of each found word in one's thread input surface area)
            '         is very minimal, which challenges the benefit/cost gains against the overhead cost of threading
            '     More importantly, ultimately we need a single (concurrent) dictionary to hold all counts, which means:
            '         - Either we keep updating the same singleton dictionary accross threads
            '             (which even if Concurrent Dict, inevitably requires some sort of thread locking in order to write
            '             (regardless of whether such locking is explicit or would happen backstage)
            '         - Or we produce a partial dictionary per thread, but then the cost of merging in the end isn't that cheap
            ' Maybe there's a clever way to exploit this better, but since benchmarking shows this is never exactly slow, 
            '     even with extremely large text arrays, and considering the time complexity is no worse than linear, 
            '         in practice I thought it fine to keep it simple (possibly naive)
            For h = 0 To texts.GetLength(0) - 1
                For w = 0 To texts.GetLength(1) - 1

                    If TypeOf texts(h, w) Is ExcelError Or TypeOf texts(h, w) Is ExcelEmpty Then Continue For
                    ThisText = If(is_case_sensitive, CStr(texts(h, w)), CStr(texts(h, w)).ToUpper())
                    If ThisText = "" Then Continue For

                    If FnFreqDict.ContainsKey(ThisText) Then
                        FnFreqDict(ThisText) += 1
                    Else
                        FnFreqDict(ThisText) = 1
                    End If

                Next w
            Next h

            If FnFreqDict.Count = 0 Then Return ExcelError.ExcelErrorNA

            Dim FnOutput(0 To FnFreqDict.Count - 1, 0 To 1) As Object

            h = 0
            For Each k As String In FnFreqDict.Keys
                FnOutput(h, 0) = k
                FnOutput(h, 1) = FnFreqDict(k)
                h += 1
            Next k

            Return FnOutput

        End Function

    End Module
End Namespace
