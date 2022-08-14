Attribute VB_Name = "DictionaryUtils"
'@Folder("VBAProject")
' This file is released under the CC0 1.0 Universal (CC0 1.0) Public Domain Dedication.
' See the LICENSE.txt file or https://creativecommons.org/publicdomain/zero/1.0/ for details.
Option Private Module
Option Explicit

Public Function DictionaryEquals(ByVal One As Object, ByVal Another As Object) As Boolean
    If One Is Another Then
        If One Is Nothing Or TypeName(One) = "Dictionary" Then
            DictionaryEquals = True
            Exit Function
        Else
            DictionaryEquals = False
            Exit Function
        End If
    End If
    
    If One Is Nothing Or Another Is Nothing Then
        DictionaryEquals = False
        Exit Function
    End If
    
    If TypeName(One) <> "Dictionary" Or TypeName(Another) <> "Dictionary" Then
        DictionaryEquals = False
        Exit Function
    End If
    
    Dim key As Variant
    For Each key In One.Keys
        If Not Another.Exists(key) Then
            DictionaryEquals = False
            Exit Function
        End If
    Next
    For Each key In Another.Keys
        If Not One.Exists(key) Then
            DictionaryEquals = False
            Exit Function
        End If
    Next
    
    For Each key In One.Keys
        If IsObject(One(key)) Then
            If IsObject(Another(key)) Then
                If Not One(key) Is Another(key) Then
                    DictionaryEquals = False
                    Exit Function
                End If
            Else
                DictionaryEquals = False
                Exit Function
            End If
        Else
            If Not IsObject(Another(key)) Then
                If One(key) <> Another(key) Then
                    DictionaryEquals = False
                    Exit Function
                End If
            Else
                DictionaryEquals = False
                Exit Function
            End If
        End If
    Next
    
    DictionaryEquals = True
End Function

Public Sub Test()
    Debug.Print TypeName(Empty)

End Sub
