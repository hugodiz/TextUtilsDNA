Option Explicit On
Option Strict On

Imports System.Text.RegularExpressions
Imports ExcelDna.Integration

Namespace TextUtilsDna
    Public Module Excel_RESUB

        ' *********************************************************************************************************************
        ' RESUB Spec : Finds all occurrences of a regex pattern in each input text and replaces it with a specified string
        ' ---------------------------------------------------------------------------------------------------------------------
        ' This function is similar to the SUBSTITUTE Excel built-in function, except it allows defining 
        '     the text portions by using a symbolic regex pattern, instead of having to specify a literal string
        ' The replacement string, however, needs to be an (almost) literal string - it allows including $1, $2, $3 ...
        '     and these bits will be replaced by the Gth captured group (if defined in the regex pattern)
        ' In other words, in the replacement string, the only allowed symbolic bits are these references to re-use
        '     captured groups
        '
        ' This function is ExceptionSafe and ThreadSafe (read note at the top of this file)
        ' *********************************************************************************************************************
        ' RESUB Function Signature
        ' ---------------------------------------------------------------------------------------------------------------------
        <ExcelFunction(
            IsExceptionSafe:=True,
            IsThreadSafe:=True,
            Description:="Replaces all occurrences of regex pattern P by a specified pattern, for each text of an input range",
            HelpTopic:="https://docs.microsoft.com/en-us/dotnet/standard/base-types/regular-expressions")>
        Function RESUB(
            <ExcelArgument(Name:="text(s)",
                           Description:="<[RANGE]> The input texts to apply the transform to:
Input is range A : each input A(h,w) = value as seen in Excel under ""General"" formatting
A Omitted : Output = #VALUE!
A(h,w) = """" Or Empty: Output(h,w) =""""")>
            texts As Object(,),
 _
            <ExcelArgument(Name:="find_pattern",
                           Description:="<[SCALAR]> Regex pattern P to find in each A(h,w):
P = """" Or Omitted Or Empty : not applied, no transform
P = valid .NET regex : replace all occurrences of P by the replacement pattern
Hint: go to http://regexstorm.net/tester to test regexes")>
            find_pattern As String,
 _
            <ExcelArgument(Name:="replace_by",
                           Description:="<[SCALAR]> Replacement pattern R to apply to each A(h,w):
R = """" Or Omitted Or Empty : replace by """"
R must be a string but may include $G to re-use the Gth captured sub-group 
Hint: go to http://regexstorm.net/tester to test regexes")>
            replace_by As String,
 _
            <ExcelArgument(Name:="[case_sensitive]",
                           Description:="<[SCALAR]> Case sensivity:
TRUE : P pattern search is Case sensitive
FALSE Or Ommitted Or Empty : Ignore Case")>
            case_sensitive As Boolean) As Object
            ' *****************************************************************************************************************
            ' RESUB Function Implementation 
            ' -----------------------------------------------------------------------------------------------------------------
            ' Notes: "Fn" control variables are defined and scoped to the main function call, and then used from within loops
            '        "My" control variables are thread-local variables within parallel loops
            '        snake_case variables are Excel function "raw" inputs - typically these will be processed and be used to
            '            define internal control variables (such as the "Fn" variables)

            If find_pattern = "" Then Return texts

            Dim FnPatternRegex As Regex

            Try
                FnPatternRegex = New Regex(find_pattern, If(case_sensitive, RegexOptions.Compiled, RegexOptions.Compiled Or RegexOptions.IgnoreCase))
            Catch ex As ArgumentException
                Return String.Format("#REGEX_P! [{0}]", ex.Message)
            End Try

            Dim FnOutput(0 To texts.GetLength(0) - 1, 0 To texts.GetLength(1) - 1) As Object

            ' The loop below is trivially parallel :
            '     Each parallel thread Th only does work on A(FromH..ToH, ..) and assigns its result to Output(FromH..ToH, ..)
            '     Note that the parallelization targets inputs rows, but not columns (because rows will tend to be much longer)
            '         so if the input range A is 1 row and an absurd number of columns, it would still be just 1 thread
            '         This could be improved, but I saw no value in practice, and also does not add anything conceptually
            '     The function will split workloads uniformly over all available cores using 1 Thread per Core
            '     For a more in-depth explanation, please read note at the top of this file
            ' -----------------------------------------------------------------------------------------------------------------
            Dim Hsize = FnOutput.GetLength(0)
            Dim WorkerThreadsH = If(Hsize < NCores, Hsize, NCores)
            Dim UniformIterationsPerThreadH = Hsize \ WorkerThreadsH
            Dim AddOneUntilH = Hsize Mod WorkerThreadsH
            Parallel.For(
            0,
            WorkerThreadsH,
            Sub(Th)

                Dim MyFromH = Th * UniformIterationsPerThreadH + If(Th < AddOneUntilH, Th, AddOneUntilH)
                Dim MyToH = (Th + 1) * UniformIterationsPerThreadH + If(Th + 1 < AddOneUntilH, Th + 1, AddOneUntilH) - 1
                Dim MyText As String

                For h = MyFromH To MyToH
                    For w = 0 To FnOutput.GetLength(1) - 1
                        If TypeOf texts(h, w) Is ExcelError Or TypeOf texts(h, w) Is ExcelEmpty Then
                            FnOutput(h, w) = texts(h, w)
                            Continue For
                        Else

                            MyText = CStr(texts(h, w))
                            If MyText = "" Then
                                FnOutput(h, w) = texts(h, w)
                                Continue For
                            End If

                            FnOutput(h, w) = FnPatternRegex.Replace(MyText, replace_by)
                        End If
                    Next w
                Next h

            End Sub) ' Next slice A(FromH..ToH, ..))

            Return FnOutput

        End Function

    End Module
End Namespace
