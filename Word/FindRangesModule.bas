Attribute VB_Name = "FindRangesModule"
' This file is released under the CC0 1.0 Universal License.
' See the LICENSE.txt file or https://creativecommons.org/publicdomain/zero/1.0/ for details.
Option Private Module
Option Explicit

Public Function FindRanges(ByVal FindCondition As Find) As Collection
    Dim foundRanges As Collection: Set foundRanges = New Collection
    
    Dim targetRange As Range
    Dim startOfRange As Long
    Dim endOfRange As Long
    Select Case TypeName(FindCondition.Parent)
        Case "Selection"
            Do While FindCondition.Execute
                Call foundRanges.Add(Selection.Range)
            Loop
        Case "Range"
            Set targetRange = FindCondition.Parent
            startOfRange = targetRange.Start
            endOfRange = targetRange.End
            Do While FindCondition.Execute
                If targetRange.Start >= startOfRange And targetRange.End <= endOfRange Then
                    Call foundRanges.Add(targetRange.Parent.Range(targetRange.Start, targetRange.End))
                End If
            Loop
    End Select
    
    Set FindRanges = foundRanges
End Function
