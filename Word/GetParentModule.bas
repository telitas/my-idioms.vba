Attribute VB_Name = "GetParentModule"
' This file is released under the CC0 1.0 Universal License.
' See the LICENSE.txt file or https://creativecommons.org/publicdomain/zero/1.0/ for details.
Option Explicit
Option Private Module

Public Function GetParentDocument(ByVal TargetObject As Object) As Document
    Select Case TypeName(TargetObject)
        Case "Document"
            Set GetParentDocument = TargetObject
            Exit Function
        Case "Application"
            Call Err.Raise(5)
    End Select
    Set GetParentDocument = GetParentDocument(TargetObject.Parent)
End Function

Public Function GetParentSection(ByVal TargetObject As Object) As Section
    Select Case TypeName(TargetObject)
        Case "Section"
            Set GetParentSection = TargetObject
            Exit Function
        Case "Document", "Application"
            Call Err.Raise(5)
    End Select
    Set GetParentSection = GetParentSection(TargetObject.Parent)
End Function
