Attribute VB_Name = "GetParentModule"
' This file is released under the CC0 1.0 Universal License.
' See the LICENSE.txt file or https://creativecommons.org/publicdomain/zero/1.0/ for details.
Option Explicit
Option Private Module

Public Function GetParentPresentation(ByVal TargetObject As Object) As Presentation
    Select Case TypeName(TargetObject)
        Case "Presentation"
            Set GetParentPresentation = TargetObject
            Exit Function
        Case "Application"
            Call Err.Raise(5)
    End Select
    Set GetParentPresentation = GetParentPresentation(TargetObject.Parent)
End Function

Public Function GetParentSlide(ByVal TargetObject As Object) As Slide
    Select Case TypeName(TargetObject)
        Case "Slide"
            Set GetParentSlide = TargetObject
            Exit Function
        Case "Application", "Presentation"
            Call Err.Raise(5)
    End Select
    Set GetParentSlide = GetParentSlide(TargetObject.Parent)
End Function
