Attribute VB_Name = "FileSystemObjectUtils"
'@Folder("VBAProject")
' This file is released under the CC0 1.0 Universal (CC0 1.0) Public Domain Dedication.
' See the LICENSE.txt file or https://creativecommons.org/publicdomain/zero/1.0/ for details.
Option Private Module
Option Explicit

Public Sub MakeFolder(ByVal Path As String, Optional ByVal IgnoreIfExists As Boolean = False)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FolderExists(Path) Then
        If IgnoreIfExists Then
            Exit Sub
        Else
            Call Err.Raise(58)
        End If
    ElseIf fso.FileExists(Path) Then
        Call Err.Raise(58)
    End If
    
    Dim parentPath As String
    parentPath = fso.GetParentFolderName(Path)
    If parentPath = "" Then
        Call Err.Raise(5)
    End If
    If Not fso.FolderExists(parentPath) Then
        Call MakeFolder(parentPath, True)
    End If
    Call fso.CreateFolder(Path)
End Sub

Public Function ListFiles(ByVal Path As String, Optional ByVal Recurse As Boolean = False, Optional ByVal RegExpFilter As String = "") As Collection
    Dim result As Collection: Set result = New Collection
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folder As Object
    Set folder = fso.GetFolder(Path)
    Dim file As Variant
    Dim filter As Object: Set filter = CreateObject("VBScript.RegExp")
    If RegExpFilter = "" Then
        filter.Pattern = "^.*$"
    Else
        filter.Pattern = RegExpFilter
    End If
    For Each file In folder.files
        If filter.Test(file.Name) Then
            Call result.Add(file.Path)
        End If
    Next
    Dim filePath As String
    If Recurse Then
        For Each folder In folder.SubFolders
            For Each file In ListFiles(folder.Path, Recurse, RegExpFilter)
                Call result.Add(file)
            Next
        Next
    End If
    Set ListFiles = result
End Function
