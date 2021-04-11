Option Explicit On
Option Strict On

Imports System.Collections.Concurrent

Namespace TextUtilsDna
    Public Module Excel_Caches

        ' *********************************************************************************************************************
        ' Excel Session Caches
        ' ---------------------------------------------------------------------------------------------------------------------
        ' These concurrent dictionaries remember things whilst the Excel application is open, memory clears when closing Excel

        ' Number of available CPU cores for multi-threading parallel calculations
        Public ReadOnly NCores As Integer = Environment.ProcessorCount

        ' Persistent cache for memoizing the results of transforming lookup_array(m,n) strings into upper case when applicable
        Public LazyStringToUpperCache As ConcurrentDictionary(Of String, String) =
            New ConcurrentDictionary(Of String, String)(NCores, 32768)

        ' Persistent cache for memoizing the results of subjecting lookup_array(m,n) strings to the P and Q filters (pass/fail)
        '     P and Q are Regular Expression patterns used as filters, and documented in the LSDLOOKUP function signature
        Public LazyFilteredCandidatesCache As ConcurrentDictionary(Of (String, String, String), String) =
            New ConcurrentDictionary(Of (String, String, String), String)(NCores, 32768)

        ' MaxInt is a general stand-in for positive infinity in the context of the LSDLOOKUP algorithm
        Public Const MaxInt As Integer = Integer.MaxValue

    End Module
End Namespace
