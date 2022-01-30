Attribute VB_Name = "GetParentModule"
' This file is released under the CC0 1.0 Universal License.
' See the LICENSE.txt file or https://creativecommons.org/publicdomain/zero/1.0/ for details.
Option Explicit
Option Private Module

Public Function GetPresentation(ByVal TargetObject As Object) As Presentation
    Select Case TypeName(TargetObject)
        Case "Presentation"
            Set GetPresentation = TargetObject
            Exit Function
        Case "Application"
            Call Err.Raise(5)
    End Select
    Set GetPresentation = GetPresentation(TargetObject.Parent)
End Function

Public Function GetSlide(ByVal TargetObject As Object) As Slide
    Select Case TypeName(TargetObject)
        Case "Slide"
            Set GetSlide = TargetObject
            Exit Function
        Case "Application", "Presentation"
            Call Err.Raise(5)
    End Select
    Set GetSlide = GetSlide(TargetObject.Parent)
End Function
