Attribute VB_Name = "GetParentModule"
' This file is released under the CC0 1.0 Universal License.
' See the LICENSE.txt file or https://creativecommons.org/publicdomain/zero/1.0/ for details.
Option Explicit
Option Private Module

Public Function GetDocument(ByVal TargetObject As Object) As Document
    Select Case TypeName(TargetObject)
        Case "Document"
            Set GetDocument = TargetObject
            Exit Function
        Case "Application"
            Call Err.Raise(5)
    End Select
    Set GetDocument = GetDocument(TargetObject.Parent)
End Function

Public Function GetSection(ByVal TargetObject As Object) As Section
    Select Case TypeName(TargetObject)
        Case "Section"
            Set GetSection = TargetObject
            Exit Function
        Case "Document", "Application"
            Call Err.Raise(5)
    End Select
    Set GetSection = GetSection(TargetObject.Parent)
End Function
