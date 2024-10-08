Attribute VB_Name = "PerformanceSettingModule"
' This file is released under the CC0 1.0 Universal License.
' See the LICENSE.txt file or https://creativecommons.org/publicdomain/zero/1.0/ for details.
Option Private Module
Option Explicit

Public Function ApplyPerformanceSetting( _
    Optional ByVal Calculation As Variant, _
    Optional ByVal EnableEvents As Variant, _
    Optional ByVal ScreenUpdating As Variant _
) As Object
    Dim currentState As Object
    Set currentState = CreateObject("Scripting.Dictionary")
    currentState("EnableEvents") = Application.EnableEvents
    currentState("ScreenUpdating") = Application.ScreenUpdating
    currentState("Calculation") = Application.Calculation
    
    Dim castedCalculation As Long
    With Application
        If Not (IsMissing(EnableEvents) Or IsEmpty(EnableEvents)) Then
            .EnableEvents = CBool(EnableEvents)
        End If
        If Not (IsMissing(ScreenUpdating) Or IsEmpty(ScreenUpdating)) Then
            .ScreenUpdating = CBool(ScreenUpdating)
        End If
        If Not (IsMissing(Calculation) Or IsEmpty(Calculation)) Then
            castedCalculation = CLng(Calculation)
            Select Case castedCalculation
                Case xlCalculationAutomatic, xlCalculationManual, xlCalculationSemiautomatic
                    .Calculation = castedCalculation
                Case Else
                    Call Err.Raise(13)
            End Select
        End If
    End With
    Set ApplyPerformanceSetting = currentState
End Function

Public Function ApplyPerformanceSettingWithDictionary(ByVal SettingDictionary As Object) As Object
    Set ApplyPerformanceSettingWithDictionary = ApplyPerformanceSetting( _
        Calculation:=SettingDictionary("Calculation"), _
        EnableEvents:=SettingDictionary("EnableEvents"), _
        ScreenUpdating:=SettingDictionary("ScreenUpdating") _
    )
End Function
