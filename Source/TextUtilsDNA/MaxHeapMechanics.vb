Option Explicit On
Option Strict On

Imports System.Runtime.CompilerServices

Public Module ImplicitHeapExtensions

    ' "Precedes" is an operation comparing two matches to see which is better:
    ' A match is the Levenstein distance between an input text A(h) and some element B(m,n) of the lookup_array
    ' The "value" of a match is its Levenstein Distance, but the position of the candidate in the lookup_array is a tie-breaker
    '     Therefore, a match is (Lev Distance, m, n)
    ' "Precedes" means it's a better match, with "better" in the sense that 
    '     a worse match can't ever exist in the best K if the better one is not also there
    '     but the opposite can be true
    <Extension()>
    Public Function Precedes(ByVal ThisTuple As (Integer, Integer, Integer), ByVal OtherTuple As (Integer, Integer, Integer)) _
        As Boolean

        Return _
            ThisTuple.Item1 < OtherTuple.Item1 _
            OrElse
            (ThisTuple.Item1 = OtherTuple.Item1 AndAlso ThisTuple.Item2 < OtherTuple.Item2) _
            OrElse
            (ThisTuple.Item1 = OtherTuple.Item1 AndAlso ThisTuple.Item2 = OtherTuple.Item2 AndAlso ThisTuple.Item3 = OtherTuple.Item3)

    End Function

    ' Unitary operation which may result in the targetted node's contents switching palces with a child, or nothing happening
    <Extension()>
    Private Function TryAdvanceNode(ByRef ThisHeap() As (Integer, Integer, Integer), ByVal CurrentNodePos As Integer) As Integer

        Dim K = ThisHeap.Length - 1

        Dim UpdatedCurrentNodePos As Integer = CurrentNodePos

        If CurrentNodePos * 2 < K Then

            Dim Child1NodePos As Integer = CurrentNodePos * 2
            Dim Child2NodePos As Integer = CurrentNodePos * 2 + 1

            Dim Parent = ThisHeap(CurrentNodePos)
            Dim Child1 = ThisHeap(Child1NodePos)
            Dim Child2 = ThisHeap(Child2NodePos)

            If Parent.Precedes(Child1) Then
                If Child2.Precedes(Child1) Then
                    ThisHeap(Child1NodePos) = Parent
                    ThisHeap(CurrentNodePos) = Child1
                    UpdatedCurrentNodePos = Child1NodePos
                Else
                    ThisHeap(Child2NodePos) = Parent
                    ThisHeap(CurrentNodePos) = Child2
                    UpdatedCurrentNodePos = Child2NodePos
                End If
            ElseIf Parent.Precedes(Child2) Then
                ThisHeap(Child2NodePos) = Parent
                ThisHeap(CurrentNodePos) = Child2
                UpdatedCurrentNodePos = Child2NodePos
            End If

        ElseIf CurrentNodePos * 2 = K Then

            Dim SingleChildNodePos As Integer = CurrentNodePos * 2

            Dim Parent = ThisHeap(CurrentNodePos)
            Dim SingleChild = ThisHeap(SingleChildNodePos)

            If Parent.Precedes(SingleChild) Then
                ThisHeap(SingleChildNodePos) = Parent
                ThisHeap(CurrentNodePos) = SingleChild
                UpdatedCurrentNodePos = SingleChildNodePos
            End If

        End If

        Return UpdatedCurrentNodePos

    End Function

    ' Tries to enqueue a node by first comparing against the root node, and if the root is overtaken,
    '     then the re-establishment of the heap property is orchestrated through a sequence of unitary TryAdvanceNode operations
    <Extension()>
    Public Sub TryEnqueueNode(ByRef ThisHeap() As (Integer, Integer, Integer), ByVal CandidateNode As (Integer, Integer, Integer))

        Dim NodeSettled = False
        Dim CurrentNodePos As Integer

        If Not CandidateNode.Precedes(ThisHeap(1)) Then
            NodeSettled = True
        Else
            CurrentNodePos = 1
            ThisHeap(1) = CandidateNode
        End If

        While Not NodeSettled
            Dim NextNodePos As Integer = ThisHeap.TryAdvanceNode(CurrentNodePos)
            If NextNodePos = CurrentNodePos Then
                NodeSettled = True
            Else
                CurrentNodePos = NextNodePos
            End If
        End While

    End Sub

    ' A tweaked version of enqueueing which serves the purpose of actually dequeueing the root node,
    '     whilst ensuring the heap readjusts such that the next-worst match is then placed at the root
    ' To be used K times by the end of LSDLOOKUP looping through the lookup_array B
    ' This works by dropping an achor (ie. enqueueing (-,1-,1-,1))
    <Extension()>
    Public Function TryDequeueNode(ByRef ThisHeap() As (Integer, Integer, Integer)) As (Integer, Integer, Integer)

        If Not (-1, -1, -1).Precedes(ThisHeap(1)) Then Return (-1, -1, -1)

        Dim AnchorSettled = False
        Dim DequeuedNode As (Integer, Integer, Integer) = ThisHeap(1)
        Dim CurrentAnchorPos As Integer = 1
        ThisHeap(1) = (-1, -1, -1)

        While Not AnchorSettled
            Dim NextAnchorPos As Integer = ThisHeap.TryAdvanceNode(CurrentAnchorPos)
            If NextAnchorPos = CurrentAnchorPos Then
                AnchorSettled = True
            Else
                CurrentAnchorPos = NextAnchorPos
            End If
        End While

        Return DequeuedNode

    End Function

End Module