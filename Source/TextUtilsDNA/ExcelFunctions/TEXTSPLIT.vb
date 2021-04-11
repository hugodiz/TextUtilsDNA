Option Explicit On
Option Strict On

Imports ExcelDna.Integration

Namespace TextUtilsDna
    Public Module Excel_TEXTSPLIT

        ' *********************************************************************************************************************
        ' TEXTSPLIT Spec : Splits a text string into a row dynamic array with delimiter-separated words
        ' ---------------------------------------------------------------------------------------------------------------------
        ' This function is the reverse to the TEXTJOIN Excel built-in function;
        '
        ' This function is ExceptionSafe and ThreadSafe (read note at the top of this file)
        ' *********************************************************************************************************************
        ' TEXTSPLIT Function Signature
        ' ---------------------------------------------------------------------------------------------------------------------
        <ExcelFunction(
            IsExceptionSafe:=True,
            IsThreadSafe:=True,
            Description:="Splits a text string into a row range using a delimiter")>
        Function TEXTSPLIT(
            <ExcelArgument(Name:="delimiter")>
            delimiter As String,
 _
            <ExcelArgument(Name:="ignore_empty",
                           Description:="Ignore empty entries when splitting
TRUE - Ignore empty entries                        
FALSE - Include empty entries")>
            ignore_empty As Boolean,
 _
            <ExcelArgument(Name:="text",
                           Description:="<[SCALAR]> The text to split")>
            text As Object) As Object
            ' *****************************************************************************************************************
            ' TEXTSPLIT Function Implementation 
            ' -----------------------------------------------------------------------------------------------------------------
            ' Notes: "Fn" control variables are defined and scoped to the main function call, and then used from within loops
            '        snake_case variables are Excel function "raw" inputs - typically these will be processed and be used to
            '            define internal control variables (such as the "Fn" variables)

            ' text must be a scalar, not an array, because we want to in general be able to output a variable number of items,
            '     in a row, per each input text
            ' Also, the simplicity of the function and the fact that no large arrays are being read to .NET per function call
            ' means the inherent parallelization of Excel when dragging down TEXTSPLIT is optimization enough
            If TypeOf text Is Object(,) OrElse TypeOf text Is ExcelMissing Then Return ExcelError.ExcelErrorValue

            ' Edge cases
            If TypeOf text Is ExcelEmpty Then Return ""
            If TypeOf text Is ExcelError Then Return text
            Dim FnText = CStr(text)
            If FnText.Length = 0 Then Return ""

            Dim FnSplitStrArr = FnText.Split({delimiter}, If(ignore_empty, StringSplitOptions.RemoveEmptyEntries, StringSplitOptions.None))
            Dim FnSplitObjArr(0 To FnSplitStrArr.Length - 1) As Object

            For x = 0 To FnSplitObjArr.Length - 1
                FnSplitObjArr(x) = FnSplitStrArr(x)
            Next x

            Return FnSplitObjArr

        End Function

    End Module
End Namespace
