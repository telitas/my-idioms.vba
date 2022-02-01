Attribute VB_Name = "GetParentModule"
' This file is released under the CC0 1.0 Universal License.
' See the LICENSE.txt file or https://creativecommons.org/publicdomain/zero/1.0/ for details.
Option Explicit
Option Private Module

Public Function GetParentDocument(ByVal TargetObject As Object) As Document
    If TypeOf TargetObject Is Document Then
        Set GetParentDocument = TargetObject
        Exit Function
    ElseIf TypeOf TargetObject Is Application Then
        Call Err.Raise(5)
    End If
    Set GetParentDocument = GetParentDocument(TargetObject.Parent)
End Function

Public Function GetParentSection(ByVal TargetObject As Object) As Section
    If TypeOf TargetObject Is Section Then
        Set GetParentSection = TargetObject
        Exit Function
    ElseIf TypeOf TargetObject Is Document Or _
            TypeOf TargetObject Is Application Then
        Call Err.Raise(5)
    End If
    Set GetParentSection = GetParentSection(TargetObject.Parent)
End Function
