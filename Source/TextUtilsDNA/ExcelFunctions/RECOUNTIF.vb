Option Explicit On
Option Strict On

Imports System.Text.RegularExpressions
Imports ExcelDna.Integration

Namespace TextUtilsDna
    Public Module Excel_RECOUNTIF

        ' *********************************************************************************************************************
        ' RECOUNTIF Spec : Counts how many entries of range of texts exhibit and/or don't exhibit a regex pattern
        ' ---------------------------------------------------------------------------------------------------------------------
        ' This function is an adaptation of Excel's COUNTIF formula, except here the criteria is given by regex pattern(s)
        '     A "positive" and a "negative" filter can be specified (one or both at the same time)
        '     A cell will be counted if and only if it is NOT excluded by either filter (ie. must pass both tests)
        '     A cell is excluded by positive filter P if it does not exhibit regex pattern P
        '     A cell is excluded by negative filter Q if it exhibits regex pattern Q
        '     If either filter, P or Q, is empty or "", then that filter is inactive and excludes nothing    
        '
        ' This function is ExceptionSafe and ThreadSafe (read note at the top of this file)
        ' *********************************************************************************************************************
        ' RECOUNTIF Function Signature
        ' ---------------------------------------------------------------------------------------------------------------------
        <ExcelFunction(
            IsExceptionSafe:=True,
            IsThreadSafe:=True,
            Description:="Counts the number of cells within a range that meet the given regex criteria",
            HelpTopic:="https://docs.microsoft.com/en-us/dotnet/standard/base-types/regular-expressions")>
        Function RECOUNTIF(
            <ExcelArgument(Name:="text(s)",
                           Description:="<[Range]> Range of Text string(s) A to conditionally count:
Input is range A : each input A(h,w) = value as seen in Excel under ""General"" formatting
A Omitted : Output = #VALUE!
A(h,w) = """" Or Error Or Empty: A(h,w) excluded")>
            texts As Object(,),
 _
            <ExcelArgument(Name:="[case_sensitive]",
                           Description:="<[SCALAR]> Case sensivity:
TRUE  : criteria check is Case sensitive
FALSE Or Ommitted Or Empty : Ignore Case")>
            case_sensitive As Boolean,
 _
            <ExcelArgument(Name:="must_have",
                           Description:="<[SCALAR]> Regex pattern P that each A(h,w) cell must exhibit:
P = """" Or Omitted Or Empty : no criterion, no exclusions here
P = valid .NET regex : exclude all A(h,w)'s NOT exhibiting P
Hint: go to http://regexstorm.net/tester to test regexes")>
            must_have As String,
 _
            <ExcelArgument(Name:="must_not_have",
                           Description:="<[SCALAR]> Regex pattern Q that each A(h,w) cell must NOT exhibit:
P = """" Or Omitted Or Empty : no criterion, no exclusions here
P = valid .NET regex : exclude all A(h,w)'s exhibiting Q
Hint: go to http://regexstorm.net/tester to test regexes")>
            must_not_have As String) As Object
            ' *****************************************************************************************************************
            ' RECOUNTIF Function Implementation 
            ' -----------------------------------------------------------------------------------------------------------------
            ' Notes: "Fn" control variables are defined and scoped to the main function call, and then used from within loops
            '        "My" control variables are thread-local variables within parallel loops
            '        snake_case variables are Excel function "raw" inputs - typically these will be processed and be used to
            '            define internal control variables (such as the "Fn" variables)

            ' Optional Regex positive filter P, which the A(h,w)'s from the input range must exhibit or else be excluded
            '     May be empty, in which case no exclusions occur
            ' Pre-compilation of P ensures later performance and guards against invalid regex pattern inputs by the user -
            '     if P is invalid, RECOUNTIF aborts with !REGEX_P and an error message detailing the regex syntactic issue
            ' Case insensitivity (if case_sensitive = False) is handled via uniform upper-case pre-conversion
            Dim FnMustHave As String = If(case_sensitive, must_have, must_have.ToUpper())
            Dim FnHasRegexPositiveFilter As Boolean = FnMustHave <> ""
            Dim FnRegexPositiveFilter As Regex
            If FnHasRegexPositiveFilter Then
                Try
                    FnRegexPositiveFilter = New Regex(FnMustHave, RegexOptions.Compiled)
                Catch ex As ArgumentException
                    Return String.Format("#REGEX_P! [{0}]", ex.Message)
                End Try
            End If

            ' Optional Regex negative filter Q, which the A(h,w)'s from the iput range must NOT exhibit or else be excluded
            '     May be empty, in which case no exclusions occur
            ' Pre-compilation of Q ensures later performance and guards against invalid regex pattern inputs by the user -
            '     if Q is invalid, LSDLOOKUP aborts with !REGEX_Q and an error message detailing the regex syntactic issue
            ' Case insensitivity (if is_case_sensitive = False) is handled via uniform upper-case pre-conversion
            Dim FnMustNotHave As String = If(case_sensitive, must_not_have, must_not_have.ToUpper())
            Dim FnHasRegexNegativeFilter As Boolean = FnMustNotHave <> ""
            Dim FnRegexNegativeFilter As Regex
            If FnHasRegexNegativeFilter Then
                Try
                    FnRegexNegativeFilter = New Regex(FnMustNotHave, RegexOptions.Compiled)
                Catch ex As ArgumentException
                    Return String.Format("#REGEX_Q! [{0}]", ex.Message)
                End Try
            End If

            ' The loop below is trivially parallel :
            '     Each parallel thread Th only does work on A(FromH..ToH, ..) and assigns its result to Output(FromH..ToH, ..)
            '     The function will split workloads uniformly over all available cores using 1 Thread per Core
            '     Note that the parallelization targets inputs rows, but not columns (because rows will tend to be much longer)
            '         so if the input range A is 1 row and an absurd number of columns, it would still be just 1 thread
            '         This could be improved, but I saw no value in practice, and also does not add anything conceptually
            '     For a more in-depth explanation, please read note at the top of this file
            ' -----------------------------------------------------------------------------------------------------------------
            Dim Hsize = texts.GetLength(0)
            Dim WorkerThreadsH = If(Hsize < NCores, Hsize, NCores)
            Dim UniformIterationsPerThreadH = Hsize \ WorkerThreadsH
            Dim AddOneUntilH = Hsize Mod WorkerThreadsH

            ' Intermediate output variable array which collects the counts from each thread, because
            '     using a single scalar Function-scoped variable to count all threads would be more inneficient, because
            '         each thread would need to lock the shared mutable state before incrementing
            ' This way, each thread works completely independently, and since the number of threads will not ever be more than
            '     a few dozens at most, summing the partial totals after threading is very cheap
            Dim FnCounts(0 To WorkerThreadsH - 1) As Integer

            Parallel.For(
            0,
            WorkerThreadsH,
            Sub(Th)

                Dim FromH = Th * UniformIterationsPerThreadH + If(Th < AddOneUntilH, Th, AddOneUntilH)
                Dim ToH = (Th + 1) * UniformIterationsPerThreadH + If(Th + 1 < AddOneUntilH, Th + 1, AddOneUntilH) - 1

                Dim MyText As String

                For h = FromH To ToH
                    For w = 0 To texts.GetLength(1) - 1

                        If TypeOf texts(h, w) Is ExcelError Or TypeOf texts(h, w) Is ExcelEmpty Then Continue For

                        MyText = If(case_sensitive, CStr(texts(h, w)), CStr(texts(h, w)).ToUpper())
                        If MyText = "" Then Continue For

                        If (FnHasRegexPositiveFilter AndAlso Not FnRegexPositiveFilter.IsMatch(MyText)) _
                            OrElse
                            (FnHasRegexNegativeFilter AndAlso FnRegexNegativeFilter.IsMatch(MyText)) Then
                            Continue For
                        End If

                        FnCounts(Th) += 1

                    Next w
                Next h

            End Sub) ' Next slice A(FromH..ToH, ..)

            Dim FnOutput As Integer
            For x = 0 To WorkerThreadsH - 1
                FnOutput += FnCounts(x)
            Next x

            Return FnOutput

        End Function

    End Module
End Namespace
