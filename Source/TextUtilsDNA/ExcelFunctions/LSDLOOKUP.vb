Option Explicit On
Option Strict On

Imports System.Text.RegularExpressions
Imports ExcelDna.Integration

Namespace TextUtilsDna
    Public Module Excel_LSDLOOKUP

        ' *********************************************************************************************************************
        ' LSDLOOKUP Spec : Levenshtein Distance Lookup
        ' ---------------------------------------------------------------------------------------------------------------------
        ' Excel worksheet function for fuzzy-lookup: LSDLOOKUP, implementing a range search based on 
        '     an "economic" version of the Levenshtein distance dynamic programming computation between 2 text strings
        ' The asymptotic time complexity of the usual dynamic programming solution for the Lev distance between 2 strings is 
        '     O(mn) on the lengths m and n of the intervening strings
        ' This "economic" version is instead quadratic on O(m(2d+1)) for each word-pair distance calculation, where 
        '     m is the length of the shortest string of the two, and 
        '     d is the max allowed Lev distance (which if surpassed, cancels the computation and moves to the next search item)

        ' This function is ExceptionSafe and ThreadSafe (read note at the top of this file)
        ' *********************************************************************************************************************
        ' LSDLOOKUP Function Signature
        ' --------------------------------------------------------------------------------------------------------------------- 
        <ExcelFunction(
            IsExceptionSafe:=True,
            IsThreadSafe:=True,
            Description:="Looks up K exact or approximate match(es) " &
                "[first K occurrences of least Levenshtein distances found] => Output is a range with K columns",
            HelpTopic:="https://en.wikipedia.org/wiki/Levenshtein_distance")>
        Function LSDLOOKUP(
            <ExcelArgument(Name:="lookup_value(s)",
                           Description:="<[COLUMN RANGE]> The value(s) to look for:
Input is column range A : Each lookup_value A(h) is the value as seen in Excel under ""General"" formatting
A Omitted Or Not a single column : Output = #VALUE!
A(h) = """" Or Empty : Output(h,1..K) = #N/A")>
            lookup_values As Object(,),
 _
            <ExcelArgument(Name:="lookup_array",
                           Description:="<[RANGE]> The candidates to search for matches in (scanned row by row):
Input is range B : each candidate B(m,n) = value as seen in Excel under ""General"" formatting
B Omitted : Output = #VALUE!
B(m,n) = """" Or Empty Or Error : B(m,n) excluded")>
            lookup_array As Object(,),
 _
            <ExcelArgument(Name:="[typo_tolerance]",
                           Description:="<[SCALAR]> Max Levenshtein distance L allowed in match(es):
L < 0 : no imposed limit, no exclusions
L = 0 or Omitted or Empty : get (the first K) exact match(es) or nothing
L > 0 : exclude all B(m,n)'s with Levenshtein distance > L")>
            typo_tolerance As Integer,
 _
            <ExcelArgument(Name:="[case_sensitive]",
                           Description:="<[SCALAR]> Case sensivity:
TRUE  : Everything is Case sensitive (the typo counting and the optional P and Q filters)
FALSE Or Ommitted Or Empty : Ignore Case everywhere (on both typo counting and the P and Q filters)")>
            case_sensitive As Boolean,
            <ExcelArgument(Name:="[K]",
                           Description:="<[SCALAR]> Number of matches K to return (by least typos, then by order found). For each A(h):
K < 2 Or Omitted Or Empty : get only the (first) best match found
K > 1 : get the K best matches as a row in the Output(h,1..K) [Max K = 1024]")>
            k_results As Integer,
 _
            <ExcelArgument(Name:="[must_have]",
                           Description:="<[SCALAR]> Regex pattern P that each B(m,n) candidate must exhibit:
P = """" Or Omitted Or Empty : not applied, no exclusions
P = valid .NET regex : exclude all B(m,n)'s NOT exhibiting P
Hint: go to http://regexstorm.net/tester to test regexes")>
            must_have As String,
 _
            <ExcelArgument(Name:="[must_not_have]",
                           Description:="<[SCALAR]> Regex pattern Q that each B(m,n) candidate must NOT exhibit:
Q = """" Or Omitted Or Empty : not applied, no exclusions
Q = valid .NET regex : exclude all B(m,n)'s exhibiting Q
Hint: go to http://regexstorm.net/tester to test regexes")>
            must_not_have As String,
 _
            <ExcelArgument(Name:="[get_index(es)]",
                           Description:="<[SCALAR]> Choice of what to return:
FALSE or Omitted or Empty : get the actual matched value(s) from B
TRUE : get the B index(es) instead (may be 2D ""packed"" coords ""[m,n]"")
Hint: If B is 2D then =INDEX(UNPACK(""[m,n]"", [1 Or 2]) gets each coord")>
            get_indexes As Boolean) As Object
            ' *****************************************************************************************************************
            ' LSDLOOKUP Function Implementation 
            ' -----------------------------------------------------------------------------------------------------------------
            ' Notes: "Fn" control variables are defined and scoped to the main function call, and then used from within loops
            '        "My" control variables are thread-local variables within parallel loops
            '        snake_case variables are Excel function "raw" inputs - typically these will be processed and be used to
            '            define internal control variables (such as the "Fn" variables)

            ' -----------------------------------------------------------------------------------------------------------------
            ' Function Input validation and pre-processing
            ' -----------------------------------------------------------------------------------------------------------------

            ' Trivial erroneous edge case - Omitted input lookup_values
            If TypeOf lookup_values(0, 0) Is ExcelMissing Then Return ExcelError.ExcelErrorValue

            ' Trivial erronoeus edge case - Omitted input lookup_array
            If TypeOf lookup_array(0, 0) Is ExcelMissing Then Return ExcelError.ExcelErrorValue

            ' Base threshold to beat (based on user specification) : any valid match must have Lev dist lower than this
            '     If set to negative, that means "infinity" (in practice, MaxInt)
            Dim FnThres = If(typo_tolerance < 0, MaxInt, typo_tolerance + 1)

            ' Number of results to return for each A(h) 
            '     ie.best K matches by greatest closeness, the order found ; also, number of columns in the Output range
            Dim TotalBs = lookup_array.GetLength(0) * lookup_array.GetLength(1)
            Dim FnK As Integer =
                If(k_results < 2,
                    1,
                    If(k_results > 1024,
                        1024,
                        If(k_results > TotalBs,
                            TotalBs,
                            k_results
                        )
                    )
                )

            ' lookup_values must be a column, not a 2D array, because we want to in general be able to output 
            '     a whole row of K matches per each lookup_value
            If lookup_values.GetLength(1) > 1 Then Return ExcelError.ExcelErrorValue

            ' Optional Regex positive filter P, which the B(m,n)'s from the lookup_array B must exhibit or else be excluded
            '     May be empty, in which case no exclusions occur
            ' Pre-compilation of P ensures later performance and guards against invalid regex pattern inputs by the user -
            '     if P is invalid, LSDLOOKUP aborts with !REGEX_P and an error message detailing the regex syntactic issue
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

            ' Optional Regex negative filter Q, which the B(m,n)'s from the lookup_array B must NOT exhibit or else be excluded
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

            ' Output will be, by design, a range of H by K, where H is the length of the input column range A
            '     each row of the Output, FnOutput(h, 1..K) contains the results of each lookup_value A(h) 
            Dim FnOutput(0 To lookup_values.GetLength(0) - 1, 0 To FnK - 1) As Object

            ' -----------------------------------------------------------------------------------------------------------------
            ' A loop -> look through lookup_values range A : A = A(H rows by 1 column) : Each element is A(h)
            ' -----------------------------------------------------------------------------------------------------------------
            ' The loop below is trivially parallel :
            '     Each parallel thread Th only does work on A(FromH..ToH) and assigns its results to FnOutput(FromH..ToH, 1..K)
            '     Each thread reads from (but does not write to) the entire lookup_array B
            '     Read/Write operations to the Excel Session Cache Dictionaries are safe because 
            '         - we 're using Concurrent Dicts
            '         - in a Cache the worst that would happen is "missing out" on trying to get a value that just got cached
            '             due to timing; in a cache, we expect to miss a lot anyway, so no harm done. That is, the thread
            '                 thought the key didn't exist and carried on with its life and then replaced the value redundantly
            '     The function will split workloads uniformly over all available cores using 1 Thread per Core
            '     For a more in-depth explanation, please read note at the top of this file
            ' -----------------------------------------------------------------------------------------------------------------

            ' Part of the calculation of the A(start..end) bounds for this worker Thread (working on a slice of A)
            Dim Hsize = FnOutput.GetLength(0)
            Dim WorkerThreadsH = If(Hsize < NCores, Hsize, NCores)
            Dim UniformIterationsPerThreadH = Hsize \ WorkerThreadsH
            Dim AddOneUntilH = Hsize Mod WorkerThreadsH
            Parallel.For(
            0,
            WorkerThreadsH,
            Sub(Th)

                ' Thread-local variables which control the inner (B) loop for each A(h) and keep the progress status of A(h)
                Dim MyAOut As (Integer, Integer, Integer)
                Dim MyAThres As Integer
                Dim MyBLoopSwitch As Boolean
                Dim MyBestK(0 To FnK) As (Integer, Integer, Integer)

                ' Thread-local worker rows of the "economic" Levenshtein computation matrix between any 2 strings;
                '     given the sequential nature of the calculation, only 2 alternating rows are needed at each time
                Dim D0(255) As Integer ' this is the 1st, 3rd, 5th ... row
                Dim D1(255) As Integer ' this is the 2nd, 4th, 6th ... row

                ' Rest of the calculation of the A(start..end) bounds for this worker Thread (working on a slice of A)
                Dim MyFromH = Th * UniformIterationsPerThreadH + If(Th < AddOneUntilH, Th, AddOneUntilH)
                Dim MyToH = (Th + 1) * UniformIterationsPerThreadH + If(Th + 1 < AddOneUntilH, Th + 1, AddOneUntilH) - 1
                For h = MyFromH To MyToH
                    ' Do work, sequentially, on each A(h) belonging to the slice of this worker Thread

                    ' ---------------------------------------------------------------------------------------------------------
                    ' A(h) validation and pre-processing
                    ' ---------------------------------------------------------------------------------------------------------

                    ' A(h) validation : if it contains an Excel error, echo the error to FnOutput(h,1..K)
                    '     Since FnOutput must contain K columns, echo as many times as necessary
                    If TypeOf lookup_values(h, 0) Is ExcelError Then
                        For k = 0 To FnK - 1
                            FnOutput(h, k) = lookup_values(h, 0)
                        Next k
                        Continue For
                    End If

                    ' A(h) validation : if it is an empty cell, assign #N/A to Output(h,1..K)
                    '     Since FnOutput must contain K columns, echo as many times as necessary
                    If TypeOf lookup_values(h, 0) Is ExcelEmpty Then
                        For k = 0 To FnK - 1
                            FnOutput(h, k) = ExcelError.ExcelErrorNA
                        Next k
                        Continue For
                    End If

                    ' A(h) pre-processing : if it is an empty cell, assign #N/A to Output(h)
                    '     Since FnOutput must contain K columns, echo as many times as necessary
                    Dim S As String = If(case_sensitive, CStr(lookup_values(h, 0)), CStr(lookup_values(h, 0)).ToUpper())
                    Dim SL As Integer = S.Length
                    If SL = 0 Then
                        For k = 0 To FnK - 1
                            FnOutput(h, k) = ExcelError.ExcelErrorNA
                        Next k
                        Continue For
                    End If

                    ' For each A(h), do the nested loop below
                    ' ---------------------------------------------------------------------------------------------------------
                    ' B loop -> look through lookup_array B : B = B(M rows by N columns) : Each element is B(m,n)
                    ' ---------------------------------------------------------------------------------------------------------

                    ' Initialization of this thread's MyBestK array of Integer triplets (Integer, Integer, Integer)
                    '     MyBestK is a binary max-heap that keeps track of the best K matches at every point in the B loop
                    '     Each match is described by a triplet (Lev distance, row position in B, column position in B)
                    ' A match is better if Lev dist is smaller, then 
                    '     tie break with row position if needed (upper row wins eg. row 3 beats row 5), then
                    '     tie break with column position if needed (upper column wins eg. col 0 beats col 1)
                    ' Row and column positions are zero-based in the triplets, minimum Lev dist is zero (ie. perfect match)

                    ' Initialize all triplets in the heap to a 'idle non-match' such that any match is better than this
                    '     so long as its Lev Dist is lower than FnThres (which any valid match must be, by design)
                    For x = 1 To FnK
                        MyBestK(x) = (FnThres, -1, -1)
                    Next x

                    ' B loop kill switch, in case we find K perfect matches (Lev dist = 0) before the end of B
                    '     In that case, there's no point continuing because the FnOutput(h,1..K) for A(h) is already determined
                    MyBLoopSwitch = True

                    ' Evolved threshold to beat: starts as the base threshold, and then gets evolved as the B loop progresses
                    '     and better and better matches are found. The evolved threshold to beat is
                    '     the Lev Dist of the worst match "so far", once there are already K matches in the heap
                    ' When there are already K matches in the heap, a new candidate B(m,n) must at least 
                    '     beat the worst match in there to secure a place in the heap (and the worst match getting kicked out)
                    ' MyAThres pertains to each A(h) in turn
                    MyAThres = FnThres

                    ' Loop though B columns
                    For m = 0 To lookup_array.GetLength(0) - 1
                        If Not MyBLoopSwitch Then Exit For

                        ' Loop through B rows
                        For n = 0 To lookup_array.GetLength(1) - 1
                            If Not MyBLoopSwitch Then Exit For

                            ' ---------------------------------------------------------------------------------------------
                            ' B(m,n) validation and pre-processing
                            ' ---------------------------------------------------------------------------------------------

                            ' Edge case where B(m,n) can match nothing by design, skip to next B(m,n)
                            If TypeOf lookup_array(m, n) Is ExcelError OrElse TypeOf lookup_array(m, n) Is ExcelEmpty Then
                                Continue For
                            End If

                            ' Handle B(m,n) case insensitivity (if applicable) by uniformizing to upper case
                            Dim T As String = If(case_sensitive,
                                            CStr(lookup_array(m, n)),
                                            LazyStringToUpperCache.GetOrAdd(
                                                CStr(lookup_array(m, n)),
                                                Function(x As String) x.ToUpper()))

                            ' Decide whether B(m,n) passes the P and Q filters (if applicable) and be lazy about it
                            '     by trying to see if already did this before for this set of inputs
                            T = If(FnMustHave = "" AndAlso FnMustNotHave = "",
                                    T,
                                    LazyFilteredCandidatesCache.GetOrAdd(
                                    (T, FnMustHave, FnMustNotHave),
                                        Function(x As (String, String, String))
                                            If (FnHasRegexPositiveFilter AndAlso Not FnRegexPositiveFilter.IsMatch(T)) _
                                            OrElse
                                            (FnHasRegexNegativeFilter AndAlso FnRegexNegativeFilter.IsMatch(T)) Then
                                                Return ""
                                            Else
                                                Return x.Item1
                                            End If
                                        End Function)
                                )

                            ' Edge case where B(m,n) can match nothing, by design : skip to next B(m,n)
                            Dim TL As Integer = T.Length
                            If TL = 0 Then Continue For

                            ' "Wiggle-room" on the Left/Right (WL/WR): 
                            '     A prescribed pair of numbers of allowed horizontal steps (in the computation matrix)
                            '         to the left/right of the main (i.e. top-left) diagonal of the matrix
                            '     It is generally not necessary to calculate the computation matrix rows in all their width
                            '         because we are not interested in any Lev distance scores below MyThres
                            '     Any matrix pathway which would involve stepping beyond the diagonal strip
                            '         defined by WL and WR, would necessarily mean a Lev score >= MyThres
                            '     Note the integer division \ instead of the usual division /
                            '     Note that if NoLimit, then WL and WR simply span the entire matrix
                            Dim Delta As Integer = TL - SL
                            Dim LevUpperBound As Integer = If(TL > SL, TL, SL)
                            Dim WL As Integer = If(MyAThres > LevUpperBound, TL, -Delta + (MyAThres - 1 + Delta) \ 2)
                            Dim WR As Integer = If(MyAThres > LevUpperBound, TL, Delta + (MyAThres - 1 - Delta) \ 2)

                            ' Calc cancellation condition : Check on the size compatibility of A(h) and B(m,n)
                            ' One of WL or WR will be negative if (and only if) 
                            '     abs(Len(A(h)) - Len(B(m, n))) > typo_tolerance
                            '         which means the string sizes alone preclude the possibility of a viable match
                            ' Intuitively, 0 is the least wiggle-room one can have, less than that just means "impossible"
                            If WL < 0 Or WR < 0 Then Continue For

                            ' Unit of storage for the score of a match : (LevDist, Row position in B, Column position in B)
                            '     A match x is better than a match y if and only if:
                            '       [ x(0) < y(0) ] 
                            '       Or 
                            '       [ x(0) = y(0) And x(1) < Y(1) ]
                            '       Or
                            '       [ x(0) = y(0) And x(1) = Y(1) And x(2) < y(2) ]
                            '     In other words, if Lev Dist is tied, then
                            '         tie break based on first occurrence found (assuming a row by row scan)
                            '     This way, there can never be an actual tie
                            MyAOut = (MyAThres, m, n)

                            ' ReDim the Levenshtein Distance computation rows if needed because of B(m,n) size
                            If TL > D0.Length - 1 Then
                                ReDim D0(TL)
                                ReDim D1(TL)
                            End If

                            ' Set D0 as the "next" row to be used (the first in this case, since we're initializing)
                            Dim IsD0Turn = True

                            ' Rb0 is the "right-most position of D0 worth calculating" 
                            '     Depends on WR and also the index of the current row that D0 represents, bounded by TL
                            '     Note that here it's the first row, hence index i = 0
                            Dim Rb0 = If(WR > TL, TL, WR)

                            ' Rb1 is to D1 what Rb0 is to D0 
                            Dim Rb1 As Integer

                            ' Lb is the "left-most position of D0/D1 worth calculating" 
                            ' Lb is to D0 and D1 what Rb0 is to D0 and Rb1 is to D1
                            '     Depends on WL and also the index of the current row, bounded by 0
                            '     Note that here it's the first row, hence index i = 0
                            Dim Lb = 0

                            ' Initialization of first row of Levenshtein computation matrix 
                            '     this is always 0,1,2,3 ... but here we only need to compute until Rb0
                            For j = 0 To Rb0
                                D0(j) = j
                            Next

                            ' Re-usable container for the current Lev matrix row's smallest value
                            '     this represents the smallest possible Lev distance score at that point
                            '        and will trigger calc cancellation is already >= MyThres
                            Dim MinOfRow As Integer

                            ' -------------------------------------------------------------------------------------------------
                            ' Calculation of Levenshtein matrix D((SL+1)x(TL+1)) : each element D(i,j)
                            ' -------------------------------------------------------------------------------------------------
                            ' This will either tell us A(h) and B(m,n) are not a viable match, or give us the match score
                            '     The score is the Lev distance (the lower the better, 0 = perfect match)
                            '     The score also includes the position of the candidate in the range B, for tie breaking
                            ' -------------------------------------------------------------------------------------------------

                            For i = 1 To SL ' First row was i = 0 (initialized above), here we start at i = 1 (second row)

                                ' Initializing MinOfRow at positive infinity by default
                                MinOfRow = MaxInt

                                ' Toggle the worker row:
                                '     If IsD0Turn = True , the "current" row is D0 and the "previous" row is D1
                                '     If IsD0Turn = False, the "current" row is D1 and the "previous" row is D0
                                IsD0Turn = Not IsD0Turn

                                ' update Lb for the current row index i
                                Lb = i - WL : If Lb < 0 Then Lb = 0

                                ' update Rb for the current row index i 
                                '      that means either update Rb0 (if it's D0 turn), or update Rb1 (if it's D1 turn)
                                If IsD0Turn Then
                                    Rb0 = i + WR : If Rb0 > TL Then Rb0 = TL
                                Else
                                    Rb1 = i + WR : If Rb1 > TL Then Rb1 = TL
                                End If

                                Select Case IsD0Turn

                                    ' Case where it is D0 turn : current row is D0, previous row is D1
                                    Case True

                                        ' -------------------------------------------------------------------------------------
                                        ' Toggler logic for D0
                                        '     the exact same logic Is applied below for D1, with the D0/D1 roles switched 

                                        ' Cycle each Levenshtein matrix cell D(i,j=[Lb..Rb]) for the current row index i
                                        '     Since it is D0 turn, D(i,j=[Lb..Rb]) == D0(j=[Lb..Rb0])
                                        For j = Lb To Rb0

                                            ' Case j = Lb 
                                            '     if j = 0, then D(i,0) = D0(0) = i, 
                                            '     otherwise just "max-out" the cell To the left Of (i, Lb)
                                            '     "maxing-out" == (exclusively) marking the bounds of the diagonal strip
                                            If j = Lb Then
                                                If j > 0 Then
                                                    D0(j - 1) = MaxInt
                                                Else
                                                    D0(j) = i
                                                    If D0(j) < MinOfRow Then MinOfRow = D0(j) ' Update MinOfRow
                                                    Continue For ' Job done for this cell, go to next j
                                                End If
                                            End If

                                            ' Case j = Rb = Rb0 (because D0 turn)
                                            '     when Rb0 is NOT more to the right than Rb1, we've hit the right-side wall
                                            '     Until then, need to max-out the cell above (i,Rb0), because Rb0 = Rb1 + 1
                                            If j = Rb0 AndAlso Rb0 > Rb1 Then D1(j) = MaxInt

                                            ' Non-edge, general case: D(i,j) = minimum of
                                            '     "cell-to-the-left + 1", 
                                            '     "cell-above + 1",
                                            '     "cell-left-and-above + cost-of-substitution",
                                            '         where cost-of-substitution = 0 if chars S(i-1) == T(j-1), else = 1
                                            ' The fact that the cell-update rule is based on propagating the minimum value 
                                            '     Is why "maxing-out" certain cells works as a path-way blocking strategy
                                            If D1(j) < D0(j - 1) Then
                                                If D1(j) < D1(j - 1) Then
                                                    D0(j) = D1(j) + 1
                                                Else
                                                    D0(j) = D1(j - 1) + If(S(i - 1) = T(j - 1), 0, 1)
                                                End If
                                            Else
                                                If D0(j - 1) < D1(j - 1) Then
                                                    D0(j) = D0(j - 1) + 1
                                                Else
                                                    D0(j) = D1(j - 1) + If(S(i - 1) = T(j - 1), 0, 1)
                                                End If
                                            End If

                                            ' Update MinOfRow
                                            If D0(j) < MinOfRow Then MinOfRow = D0(j)
                                        Next j

                                        ' -------------------------------------------------------------------------------------
                                        ' END of Toggler logic for D0
                                        '     the exact same logic Is applied below for D1, with the D0/D1 roles switched 

                                    ' Case where it is D1 turn : current row is D1, previous row is D0
                                    Case False

                                        ' -------------------------------------------------------------------------------------
                                        ' Toggler logic for D1
                                        '     the exact same logic Is applied above for D0, with the D1/D0 roles switched 

                                        ' Cycle each Levenshtein matrix cell D(i,j=[Lb..Rb]) for the current row index i
                                        '     Since it is D1 turn, D(i,j=[Lb..Rb]) == D1(j=[Lb..Rb1])
                                        For j = Lb To Rb1

                                            ' Case j = Lb 
                                            '     if j = 0, then D(i,0) = D1(0) = i, 
                                            '     otherwise just "max-out" the cell To the left Of (i, Lb)
                                            '     "maxing-out" == (exclusively) marking the bounds of the diagonal strip
                                            If j = Lb Then
                                                If j > 0 Then
                                                    D1(j - 1) = MaxInt
                                                Else
                                                    D1(j) = i
                                                    If D1(j) < MinOfRow Then MinOfRow = D1(j) ' Update MinOfRow
                                                    Continue For ' Job done for this cell, go to next j
                                                End If
                                            End If

                                            ' Case j = Rb = Rb1 (because D1 turn)
                                            '     when Rb1 is NOT more to the right than Rb0, we've hit the right-side wall
                                            '     Until then, need to max-out the cell above (i,Rb1), because Rb1 = Rb0 + 1
                                            If j = Rb1 AndAlso Rb1 > Rb0 Then D0(j) = MaxInt

                                            ' Non-edge, general case: D(i,j) = minimum of
                                            '     "cell-to-the-left + 1", 
                                            '     "cell-above + 1",
                                            '     "cell-left-and-above + cost-of-substitution",
                                            '         where cost-of-substitution = 0 if chars S(i-1) == T(j-1), else = 1
                                            ' The fact that the cell-update rule is based on propagating the minimum value 
                                            '     Is why "maxing-out" certain cells works as a path-way blocking strategy
                                            If D0(j) < D1(j - 1) Then
                                                If D0(j) < D0(j - 1) Then
                                                    D1(j) = D0(j) + 1
                                                Else
                                                    D1(j) = D0(j - 1) + If(S(i - 1) = T(j - 1), 0, 1)
                                                End If
                                            Else
                                                If D1(j - 1) < D0(j - 1) Then
                                                    D1(j) = D1(j - 1) + 1
                                                Else
                                                    D1(j) = D0(j - 1) + If(S(i - 1) = T(j - 1), 0, 1)
                                                End If
                                            End If

                                            ' Update MinOfRow
                                            If D1(j) < MinOfRow Then MinOfRow = D1(j)
                                        Next j

                                        ' -------------------------------------------------------------------------------------
                                        ' END of Toggler logic for D1
                                        '     the exact same logic Is applied above for D0, with the D1/D0 roles switched 

                                End Select

                                ' Here we reached the end of the Lev computation matrix row i (in practice either D0 or D1)
                                '     If MinOfRow is too high, cancel the rest of the calculation and signal it via D0(0)
                                '         because we already know it's too high and wouldn't enter the K heap anyway
                                If Not MinOfRow < MyAThres Then
                                    D0(0) = -1
                                    Exit For
                                End If

                            Next i ' next row i

                            ' if early-exited i cycle, we know A(h) and B(m,n) are not a viable match, skip to next B(m,n)
                            If D0(0) = -1 Then Continue For

                            ' if i was NOT early-exited, the final Lev distance is stored in either D0(TL) or D1(TL)
                            '     depending on whose was the final turn (still reflected in IsD0Turn now)
                            ' We must now form the score triplet for the B(m,n) and pit it against the K heap
                            ' It may or may not get enqueued, depending on whether it's at least better than
                            '     the worst match already there
                            MyAOut.Item1 = If(IsD0Turn, D0(TL), D1(TL))
                            MyBestK.TryEnqueueNode(MyAOut)

                            ' MyBestK(1) contains the worst match thus far IF there are already K matches in the heap
                            ' This means we can update the evolved threshold and early-exit subsequent Lev calculations
                            '     for the subsequent B(m,n)'s if they will surely be MyAthres or higher in Lev dist
                            MyAThres = MyBestK(1).Item1

                            ' If the evolved thresold MyAThres ever becomes zero (for this A(h)), that means that
                            '     there exist already K perfect matches in the heap 
                            '         (because even the worst one, kept in MyBestK(1), has Lev dist zero)
                            ' No future B(m,n) can then steal a place in the heap, so we're done with this A(h)
                            '     if so we set MyBLoopSwitch to False to kill the B loop
                            MyBLoopSwitch = Not MyAThres = 0
                        Next n
                    Next m

                    ' -----------------------------------------------------------------------------------------------------
                    ' Post-search assignment of result to FnOutput(h,1..K)
                    ' -----------------------------------------------------------------------------------------------------

                    ' We must fill, for each A(h), the output row FnOutput(h, 1..K)
                    '     In general we will have X actual matches, where 0 <= x <= K
                    '     If x < K, fill the right-most (unused) slots of FnOutput(h, x+1..K ) with #N/A
                    '     If x = 0, the entire row simply gets #N/A by extension
                    ' The dequeuing process is detailed in the MaxHeapMechanics code file, however:
                    '     Each dequeue attempt "drops an anchor" into the heap, 
                    '         meaning it just tries to insert the triplet (-1, -1, -1), which always succeds until
                    '             the heap is full of anchors
                    '     Each time an anchor is dropped, the worst match of the heap (the root node) pops out
                    '     Some of first pop-outs may be "idle non matches", meaning just unused slots
                    '         (which will have row/col positions = -1, and translated to #N/A
                    '     Either way, all we need to then do is keep filling the output row FnOutput(h, 1..K) back to front
                    '
                    Dim DequeuedMatch As (Integer, Integer, Integer)
                    For x = FnK - 1 To 0 Step -1
                        DequeuedMatch = MyBestK.TryDequeueNode()
                        FnOutput(h, x) =
                            If(DequeuedMatch.Item2 = -1,
                                ExcelError.ExcelErrorNA,
                                If(get_indexes,
                                    If(lookup_array.GetLength(1) = 1,
                                        DequeuedMatch.Item2 + 1,
                                        If(lookup_array.GetLength(0) = 1,
                                            DequeuedMatch.Item3 + 1,
                                            CObj(String.Format("[{0},{1}]", DequeuedMatch.Item1 + 1, DequeuedMatch.Item2 + 1))
                                        )
                                    ),
                                    lookup_array(DequeuedMatch.Item2, DequeuedMatch.Item3)
                                )
                            )
                    Next x

                Next h

            End Sub) ' Next Slice A(FromH..ToH)

            Return FnOutput ' LSDLOOKUP final return of the FnOutput array, with a match result K-row for each A(h)

        End Function

    End Module
End Namespace
