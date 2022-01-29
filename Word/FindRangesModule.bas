Attribute VB_Name = "FindRangesModule"
' This file is released under the CC0 1.0 Universal License.
' See the LICENSE.txt file or https://creativecommons.org/publicdomain/zero/1.0/ for details.
Option Private Module
Option Explicit

Public Function FindRanges(ByVal FindCondition As Find) As Collection
    Dim foundRanges As Collection: Set foundRanges = New Collection
    
    Dim currentSelectionRange As Range
    Dim targetRange As Range
    Dim currentRange As Range
    Select Case TypeName(FindCondition.Parent)
        Case "Selection"
            Set currentSelectionRange = Selection.Range
            Do While FindCondition.Execute
                Call foundRanges.Add(Selection.Range)
            Loop
            With Selection
                .Start = currentSelectionRange.Start
                .End = currentSelectionRange.End
            End With
        Case "Range"
            Set targetRange = FindCondition.Parent
            Set currentRange = targetRange.Parent.Range(targetRange.Start, targetRange.End)
            Do While FindCondition.Execute
                If targetRange.Start >= currentRange.Start And targetRange.End <= currentRange.End Then
                    Call foundRanges.Add(targetRange.Parent.Range(currentRange.Start, currentRange.End))
                End If
            Loop
            With targetRange
                .Start = currentRange.Start
                .End = currentRange.End
            End With
    End Select
    
    Set FindRanges = foundRanges
End Function
