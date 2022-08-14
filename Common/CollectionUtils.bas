Attribute VB_Name = "CollectionUtils"
'@Folder("VBAProject")
' This file is released under the CC0 1.0 Universal (CC0 1.0) Public Domain Dedication.
' See the LICENSE.txt file or https://creativecommons.org/publicdomain/zero/1.0/ for details.
Option Private Module
Option Explicit

Public Function CollectionContains(ByVal Target As Collection, ByVal value As Variant)
    If Target Is Nothing Then
        Call Err.Raise(5)
    End If
    
    Dim i As Long
    Dim element As Variant
    If IsObject(value) Then
        For Each element In Target
            If IsObject(element) Then
                If element Is value Then
                    CollectionContains = True
                    Exit Function
                End If
            End If
        Next
    Else
        For Each element In Target
            If Not IsObject(element) Then
                If element = value Then
                    CollectionContains = True
                    Exit Function
                End If
            End If
        Next
    End If
    CollectionContains = False
End Function

Public Function CollectionEqualsAsList(ByVal One As Collection, ByVal Another As Collection) As Boolean
    If One Is Another Then
        CollectionEqualsAsList = True
        Exit Function
    End If
    
    If One Is Nothing Or Another Is Nothing Then
        CollectionEqualsAsList = False
        Exit Function
    End If
    
    If One.Count <> Another.Count Then
        CollectionEqualsAsList = False
        Exit Function
    End If
    
    Dim i As Long
    For i = 1 To One.Count
        If IsObject(One(i)) Then
            If IsObject(Another(i)) Then
                If Not One(i) Is Another(i) Then
                    CollectionEqualsAsList = False
                    Exit Function
                End If
            Else
                CollectionEqualsAsList = False
                Exit Function
            End If
        Else
            If Not IsObject(Another(i)) Then
                If One(i) <> Another(i) Then
                    CollectionEqualsAsList = False
                    Exit Function
                End If
            Else
                CollectionEqualsAsList = False
                Exit Function
            End If
        End If
    Next
    CollectionEqualsAsList = True
End Function

Public Function CollectionEqualsAsSet(ByVal One As Collection, ByVal Another As Collection) As Boolean
    If One Is Another Then
        CollectionEqualsAsSet = True
        Exit Function
    End If
    
    If One Is Nothing Or Another Is Nothing Then
        CollectionEqualsAsSet = False
        Exit Function
    End If
    
    Dim contains As Boolean
    Dim element As Variant
    
    For Each element In One
        If Not CollectionContains(Another, element) Then
            CollectionEqualsAsSet = False
            Exit Function
        End If
    Next
    
    For Each element In Another
        If Not CollectionContains(One, element) Then
            CollectionEqualsAsSet = False
            Exit Function
        End If
    Next
    
    CollectionEqualsAsSet = True
End Function

Public Function CollectionDistinct(ByVal Target As Collection) As Collection
    If Target Is Nothing Then
        Call Err.Raise(5)
    End If
    
    Dim i As Long
    Dim j As Long
    Dim element As Variant
    Dim newOne As Collection: Set newOne = New Collection
    Dim contained As Boolean
    For i = 1 To Target.Count
        element = Target(i)
        contained = False
        For j = 1 To i - 1
            If Target(j) = element Then
                contained = True
                Exit For
            End If
        Next
        If Not contained Then
            newOne.Add (element)
        End If
    Next
    
    Set CollectionDistinct = newOne
End Function
