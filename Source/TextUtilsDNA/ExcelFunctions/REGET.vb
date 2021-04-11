Option Explicit On
Option Strict On

Imports System.Text.RegularExpressions
Imports ExcelDna.Integration

Namespace TextUtilsDna
    Public Module Excel_REGET

        ' *********************************************************************************************************************
        ' REGET Spec : Gets the specified occurrence of the captured group(s) of a regex pattern in each input text
        ' ---------------------------------------------------------------------------------------------------------------------
        ' This function exposes a variation of the .NET Regex.Matches to Excel
        ' For each input text A(h,w), it runs a specified regex and outputs a string containing the ith occurrence of
        '     a match, or optionally just the jth captured group of said match
        '
        ' This function is ExceptionSafe and ThreadSafe (read note at the top of this file)
        ' *********************************************************************************************************************
        ' REGET Function Signature
        ' ---------------------------------------------------------------------------------------------------------------------
        <ExcelFunction(
            IsExceptionSafe:=True,
            IsThreadSafe:=True,
            Description:="Gets the Nth occurrence of regex pattern P in (each) input text",
            HelpTopic:="https://docs.microsoft.com/en-us/dotnet/standard/base-types/regular-expressions")>
        Function REGET(
            <ExcelArgument(Name:="text(s)",
                           Description:="<[RANGE]> The input texts to extract patterns from:
Input is range A : each input A(h,w) = value as seen in Excel under ""General"" formatting
A Omitted : Output = #VALUE!
A(h,w) = """" Or Empty: Output(h,w) =""""")>
            texts As Object(,),
 _
            <ExcelArgument(Name:="find_pattern",
                           Description:="<[SCALAR]> Regex pattern P to find in each A(h,w):
P = """" Or Omitted Or Empty : not applied, Output = #N/A 
P = valid .NET regex : retrieve Nth matched instance of P
Hint: go to http://regexstorm.net/tester to test regexes")>
            find_pattern As String,
 _
            <ExcelArgument(Name:="[instance_num]",
                           Description:="<[SCALAR]> Match occurrence i to extract from each A(h,w):
i < 2 Or Omitted Or Empty : get the 1st occurrence of P in each A(h,w)
i > 1 : get the ith occurrence of P in each A(h,w) [if i > number of occurrences, Output(h,w) = #N/A]")>
            instance_num As Integer,
 _
            <ExcelArgument(Name:="[group_num]",
                           Description:="<[SCALAR]> Captured group(s) j to extract from the ith occurrence of P in A(h,w)]:
j < 1 Or Omitted Or Empty : get the entire ith match occurrence, NOT just a particular captured group
j > 0 : get just the jth captured group [if j > number of groups, Output(h,w) = #N/A]")>
            group_num As Integer,
 _
            <ExcelArgument(Name:="[case_sensitive]",
                           Description:="<[SCALAR]> Case sensivity:
TRUE : P pattern search is Case sensitive
FALSE Or Ommitted Or Empty : Ignore Case")>
            is_case_sensitive As Boolean) As Object
            ' *****************************************************************************************************************
            ' REGET Function Implementation 
            ' -----------------------------------------------------------------------------------------------------------------
            ' Notes: "Fn" control variables are defined and scoped to the main function call, and then used from within loops
            '        "My" control variables are thread-local variables within parallel loops
            '        snake_case variables are Excel function "raw" inputs - typically these will be processed and be used to
            '            define internal control variables (such as the "Fn" variables)

            If find_pattern = "" Then Return texts

            Dim FnPatternRegex As Regex

            Try
                FnPatternRegex = New Regex(find_pattern, If(is_case_sensitive, RegexOptions.Compiled, RegexOptions.Compiled Or RegexOptions.IgnoreCase))
            Catch ex As ArgumentException
                Return String.Format("#REGEX_P! [{0}]", ex.Message)
            End Try

            Dim FnMatchIndex = If(instance_num < 1, 1, instance_num)
            Dim FnGroupIndex = If(group_num < 0, 0, group_num)

            Dim FnOutput(0 To texts.GetLength(0) - 1, 0 To texts.GetLength(1) - 1) As Object

            ' The loop below is trivially parallel :
            '     Each parallel thread Th only does work on A(FromH..ToH, ..) and assigns its result to Output(FromH..ToH, ..)
            '     The function will split workloads uniformly over all available cores using 1 Thread per Core
            '     Note that the parallelization targets inputs rows, but not columns (because rows will tend to be much longer)
            '         so if the input range A is 1 row and an absurd number of columns, it would still be just 1 thread
            '         This could be improved, but I saw no value in practice, and also does not add anything conceptually
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

                Dim FromH = Th * UniformIterationsPerThreadH + If(Th < AddOneUntilH, Th, AddOneUntilH)
                Dim ToH = (Th + 1) * UniformIterationsPerThreadH + If(Th + 1 < AddOneUntilH, Th + 1, AddOneUntilH) - 1

                Dim MyText As String
                Dim MyMatch As Match
                Dim MyMatches As MatchCollection

                For h = FromH To ToH
                    For w = 0 To FnOutput.GetLength(1) - 1

                        If TypeOf texts(h, w) Is ExcelError Or TypeOf texts(h, w) Is ExcelEmpty Then
                            FnOutput(h, w) = ExcelError.ExcelErrorNA
                            Continue For
                        End If

                        MyText = CStr(texts(h, w))
                        If MyText = "" Then
                            FnOutput(h, w) = ExcelError.ExcelErrorNA
                            Continue For
                        End If


                        If FnMatchIndex > 1 Then
                            MyMatches = FnPatternRegex.Matches(MyText)

                            If FnMatchIndex > MyMatches.Count Then
                                FnOutput(h, w) = ExcelError.ExcelErrorNA
                                Continue For
                            End If

                            MyMatch = MyMatches(FnMatchIndex - 1)
                        Else
                            MyMatch = FnPatternRegex.Match(MyText)
                        End If

                        If FnGroupIndex > MyMatch.Groups.Count - 1 Then
                            FnOutput(h, w) = ExcelError.ExcelErrorNA
                            Continue For
                        End If

                        FnOutput(h, w) = MyMatch.Groups(FnGroupIndex).Value

                    Next w
                Next h

            End Sub) ' Next slice A(FromH..ToH, ..)

            Return FnOutput

        End Function

    End Module
End Namespace
