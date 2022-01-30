Attribute VB_Name = "GetParentModule"
' This file is released under the CC0 1.0 Universal License.
' See the LICENSE.txt file or https://creativecommons.org/publicdomain/zero/1.0/ for details.
Option Explicit
Option Private Module

Public Function GetWorkbook(ByVal TargetObject As Object) As Workbook
    Select Case TypeName(TargetObject)
        Case "Workbook"
            Set GetWorkbook = TargetObject
            Exit Function
        Case "Application"
            Call Err.Raise(5)
    End Select
    Set GetWorkbook = GetWorkbook(TargetObject.Parent)
End Function

Public Function GetWorksheet(ByVal TargetObject As Object) As Worksheet
    Select Case TypeName(TargetObject)
        Case "Worksheet"
            Set GetWorksheet = TargetObject
            Exit Function
        Case "Application", "Workbook"
            Call Err.Raise(5)
    End Select
    Set GetWorksheet = GetWorksheet(TargetObject.Parent)
End Function
