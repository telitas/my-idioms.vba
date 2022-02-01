Attribute VB_Name = "GetParentModule"
' This file is released under the CC0 1.0 Universal License.
' See the LICENSE.txt file or https://creativecommons.org/publicdomain/zero/1.0/ for details.
Option Explicit
Option Private Module

Public Function GetParentWorkbook(ByVal TargetObject As Object) As Workbook
    If TypeOf TargetObject Is Workbook Then
        Set GetParentWorkbook = TargetObject
        Exit Function
    ElseIf TypeOf TargetObject Is Application Then
        Call Err.Raise(5)
    End If
    Set GetParentWorkbook = GetParentWorkbook(TargetObject.Parent)
End Function

Public Function GetParentWorksheet(ByVal TargetObject As Object) As Worksheet
    If TypeOf TargetObject Is Worksheet Then
        Set GetParentWorksheet = TargetObject
        Exit Function
    ElseIf TypeOf TargetObject Is Workbook Or _
            TypeOf TargetObject Is Application Then
        Call Err.Raise(5)
    End If
    Set GetParentWorksheet = GetParentWorksheet(TargetObject.Parent)
End Function

Public Function GetParentRange(ByVal TargetObject As Object) As Range
    If TypeOf TargetObject Is Range Then
        Set GetParentRange = TargetObject
        Exit Function
    ElseIf TypeOf TargetObject Is Worksheet Or _
            TypeOf TargetObject Is Workbook Or _
            TypeOf TargetObject Is Application Then
        Call Err.Raise(5)
    End If
    Set GetParentRange = GetParentRange(TargetObject.Parent)
End Function
