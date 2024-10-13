Attribute VB_Name = "RecalcSelectedRangeModule"
' This file is released under the CC0 1.0 Universal License.
' See the LICENSE.txt file or https://creativecommons.org/publicdomain/zero/1.0/ for details.
Option Explicit

Public Sub RecalcSelectedRange()
    Dim targetRange As Range
    Set targetRange = Selection
    targetRange.Formula2 = targetRange.Formula2
End Sub
