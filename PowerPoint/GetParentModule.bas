Attribute VB_Name = "GetParentModule"
' This file is released under the CC0 1.0 Universal License.
' See the LICENSE.txt file or https://creativecommons.org/publicdomain/zero/1.0/ for details.
Option Explicit
Option Private Module

Public Function GetParentPresentation(ByVal TargetObject As Object) As Presentation
    If TypeOf TargetObject Is Presentation Then
        Set GetParentPresentation = TargetObject
        Exit Function
    ElseIf TypeOf TargetObject Is Application Then
        Call Err.Raise(5)
    End If
    Set GetParentPresentation = GetParentPresentation(TargetObject.Parent)
End Function

Public Function GetParentSlide(ByVal TargetObject As Object) As Slide
    If TypeOf TargetObject Is Slide Then
        Set GetParentSlide = TargetObject
        Exit Function
    ElseIf TypeOf TargetObject Is Presentation Or _
            TypeOf TargetObject Is Application Then
        Call Err.Raise(5)
    End If
    Set GetParentSlide = GetParentSlide(TargetObject.Parent)
End Function
