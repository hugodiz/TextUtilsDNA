Option Explicit On
Option Strict On

Imports ExcelDna.Integration
Imports Newtonsoft.Json

Namespace TextUtilsDna
    Public Module Excel_UNPACK

        ' *********************************************************************************************************************
        ' UNPACK Spec : Unpack JSON notation array into an Excel dynamic array
        ' ---------------------------------------------------------------------------------------------------------------------
        ' The LSDLOOKUP function has a mode of working called get_index(es), where
        '     it returns the coordinates (int Excel-style 1-based row and column position) of each matched value in the range B
        '         instead of the matched values themselves
        '     If B is a 1D range (row or column), then a single index is returned
        '     If B is 2D, then a (row, col) pair is returned
        ' In order to allow a (row,col) pair to be displayed in a single cell, the pair is "packed" into a JSON string = [x,y]
        ' This function can take the JSON representation any so "packed" array and "unpacks" it into the actual Excel range
        '     as a dynamic array
        '
        ' This function is ExceptionSafe and ThreadSafe (read note at the top of this file)
        ' *********************************************************************************************************************
        ' UNPACK Function Signature
        ' ---------------------------------------------------------------------------------------------------------------------
        <ExcelFunction(
            IsExceptionSafe:=True,
            IsThreadSafe:=True,
            Description:=
            "Unpacks a serialized Excel range back into the corresponding array (1D ranges are rows by convention)",
            HelpTopic:="https://www.newtonsoft.com/json/help/html/SerializingJSON.htm")>
        Function UNPACK(
            <ExcelArgument(Name:="packed_range",
                           Description:=
            "<[SCALAR]> The JSON text representation of a 1D or 2D array A i.e. [A(1),A(2),..] or [[A(1,1),A(1,2),..],[A(2,1),A(2,2),..],..]")>
            packed As Object) As Object
            ' *****************************************************************************************************************
            ' UNPACK Function Implementation 

            ' packed_range must be a scalar, not an array, because we want to in general be able to output a variable range
            '     per each packed_range
            ' Also, the simplicity of the function and the fact that no large arrays are being read to .NET per function call
            ' means the inherent parallelization of Excel when dragging UNPACK (if geometrically sound) is optimization enough
            If TypeOf packed Is Object(,) OrElse TypeOf packed Is ExcelMissing Then Return ExcelError.ExcelErrorValue

            ' Edge cases
            If TypeOf packed Is ExcelEmpty Then Return ""
            If TypeOf packed Is ExcelError Then Return packed
            Dim FnPacked = CStr(packed)
            If FnPacked.Length = 0 Then Return ""

            Try
                If FnPacked(0) <> "[" Then Return JsonConvert.DeserializeObject(Of Object)(FnPacked)
                If FnPacked(1) <> "[" Then Return JsonConvert.DeserializeObject(Of Object())(FnPacked)
                Return JsonConvert.DeserializeObject(Of Object(,))(FnPacked)
            Catch ex As Exception
                Return "#JSON! [Invalid array]"
            End Try

        End Function

    End Module
End Namespace
