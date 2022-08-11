Attribute VB_Name = "FileSystemObjectUtilsTest"
Option Explicit
Option Private Module

'@TestModule
'@Folder "VBAProject.Tests"

#Const LateBind = LateBindTests

#If LateBind Then
    Private Assert As Object
    Private Fakes As Object
#Else
    Private Assert As Rubberduck.AssertClass
    Private Fakes As Rubberduck.FakesProvider
#End If

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    #If LateBind Then
        Set Assert = CreateObject("Rubberduck.AssertClass")
        Set Fakes = CreateObject("Rubberduck.FakesProvider")
    #Else
        Set Assert = New Rubberduck.AssertClass
        Set Fakes = New Rubberduck.FakesProvider
    #End If
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

Private Function TestDirectoryName(ByVal Prefix As String) As String
    TestDirectoryName = Prefix & "_" & Format(Now, "yyyymmddhhnnss")
End Function

Private Function CollectionContains(ByVal Target As Collection, ByVal value As Variant)
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

'@TestMethod("MakeFolder")
Private Sub MakeFolder_CorrectCall_Successed()
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error GoTo CATCH
    Dim testRootPath As String
    
    testRootPath = fso.BuildPath(fso.GetSpecialFolder(2), TestDirectoryName("MakeFolder"))
    
    Call MakeFolder(testRootPath)
    Dim testPath As String
    testPath = fso.BuildPath(testRootPath, "path\to\directory")
    Call MakeFolder(testPath)
    
    Call MakeFolder(testPath, True)
    
    GoTo FINALLY
CATCH:
    Assert.Fail
    GoTo FINALLY
FINALLY:
    Call fso.DeleteFolder(testRootPath)
End Sub

'@TestMethod("MakeFolder")
Private Sub MakeFolder_IllegalPath_Successed()
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error GoTo CATCH
    Dim testRootPath As String
    
    testRootPath = "ILLEGALPATH:\"
    
    On Error GoTo ERROR_RAISED
    
    Call MakeFolder(testRootPath, True)
    
    Assert.Fail
    Exit Sub
    
ERROR_RAISED:
    Assert.AreEqual CLng(5), Err.Number
    Resume FINALLY
    
CATCH:
    Assert.Fail Err.Description
    GoTo FINALLY
FINALLY:
End Sub


'@TestMethod("MakeFolder")
Private Sub MakeFolder_FileIsExists_Successed()
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error GoTo CATCH
    Dim testRootPath As String
    
    testRootPath = fso.BuildPath(fso.GetSpecialFolder(2), TestDirectoryName("MakeFolder"))
    Call fso.CreateTextFile(testRootPath)
    
    On Error GoTo ERROR_RAISED
    
    Call MakeFolder(testRootPath, True)
    
    Assert.Fail
    Exit Sub
    
ERROR_RAISED:
    Assert.AreEqual CLng(58), Err.Number
    Resume FINALLY
    
CATCH:
    Assert.Fail Err.Description
    GoTo FINALLY
FINALLY:
    Call fso.DeleteFile(testRootPath)
End Sub

'@TestMethod("MakeFolder")
Private Sub MakeFolder_DirectoryIsExistsAndNotIgnore_Successed()
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error GoTo CATCH
    Dim testRootPath As String
    
    testRootPath = fso.BuildPath(fso.GetSpecialFolder(2), TestDirectoryName("MakeFolder"))
    Call fso.CreateFolder(testRootPath)
    
    On Error GoTo ERROR_RAISED
    
    Call MakeFolder(testRootPath, False)
    
    Assert.Fail
    Exit Sub
    
ERROR_RAISED:
    Assert.AreEqual CLng(58), Err.Number
    Resume FINALLY
    
CATCH:
    Assert.Fail Err.Description
    GoTo FINALLY
FINALLY:
    Call fso.DeleteFolder(testRootPath)
End Sub

'@TestMethod("MakeFolder")
Private Sub MakeFolder_DirectoryIsExists_Successed()
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error GoTo CATCH
    Dim testRootPath As String
    
    testRootPath = fso.BuildPath(fso.GetSpecialFolder(2), TestDirectoryName("MakeFolder"))
    Call fso.CreateFolder(testRootPath)
    
    On Error GoTo ERROR_RAISED
    
    Call MakeFolder(testRootPath)
    
    Assert.Fail
    Exit Sub
    
ERROR_RAISED:
    Assert.AreEqual CLng(58), Err.Number
    Resume FINALLY
    
CATCH:
    Assert.Fail Err.Description
    GoTo FINALLY
FINALLY:
    Call fso.DeleteFolder(testRootPath)
End Sub


'@TestMethod("ListFiles")
Private Sub ListFiles_CorrectCall_Successed()
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error GoTo CATCH
    Dim testRootPath As String
    
    testRootPath = fso.BuildPath(fso.GetSpecialFolder(2), TestDirectoryName("ListFiles"))
    Dim testPath As String
    testPath = testRootPath
    
    Dim result As Collection
    
    Call fso.CreateFolder(testPath)
    Set result = ListFiles(testPath)
    Assert.AreEqual CLng(0), result.Count
    
    Dim filePaths As Collection: Set filePaths = New Collection
        
    Call filePaths.Add(fso.BuildPath(testPath, "file1.txt"))
    Call filePaths.Add(fso.BuildPath(testPath, "file2.txt"))
    Call filePaths.Add(fso.BuildPath(testPath, "file3.txt"))
    Call filePaths.Add(fso.BuildPath(testPath, "file4.txt"))
    Call filePaths.Add(fso.BuildPath(testPath, "file5.txt"))
    
    Dim filePathsInPathFolder As Collection: Set filePathsInPathFolder = New Collection
    Dim filePathsUnderPathFolder As Collection: Set filePathsUnderPathFolder = New Collection
    testPath = fso.BuildPath(testPath, "path")
    Dim pathFolder As String
    pathFolder = testPath
    Call fso.CreateFolder(testPath)
    Call filePathsInPathFolder.Add(fso.BuildPath(testPath, "file1.txt"))
    Call filePathsInPathFolder.Add(fso.BuildPath(testPath, "file2.txt"))
    Call filePathsInPathFolder.Add(fso.BuildPath(testPath, "file3.txt"))
    Call filePathsInPathFolder.Add(fso.BuildPath(testPath, "file4.txt"))
    Call filePathsInPathFolder.Add(fso.BuildPath(testPath, "file5.html"))
    
    Dim filePath As Variant
    For Each filePath In filePathsInPathFolder
        Call filePathsUnderPathFolder.Add(filePath)
    Next
    
    testPath = fso.BuildPath(pathFolder, "to1")
    Call fso.CreateFolder(testPath)
    
    Call filePathsUnderPathFolder.Add(fso.BuildPath(testPath, "file1.txt"))
    Call filePathsUnderPathFolder.Add(fso.BuildPath(testPath, "file2.txt"))
    Call filePathsUnderPathFolder.Add(fso.BuildPath(testPath, "file3.html"))
    Call filePathsUnderPathFolder.Add(fso.BuildPath(testPath, "file4.html"))
    Call filePathsUnderPathFolder.Add(fso.BuildPath(testPath, "file5.css"))
    
    testPath = fso.BuildPath(pathFolder, "to2")
    Call fso.CreateFolder(testPath)
    
    Call filePathsUnderPathFolder.Add(fso.BuildPath(testPath, "file1.txt"))
    Call filePathsUnderPathFolder.Add(fso.BuildPath(testPath, "file2.html"))
    Call filePathsUnderPathFolder.Add(fso.BuildPath(testPath, "file3.html"))
    Call filePathsUnderPathFolder.Add(fso.BuildPath(testPath, "file4.css"))
    Call filePathsUnderPathFolder.Add(fso.BuildPath(testPath, "file5.css"))
    
    For Each filePath In filePathsUnderPathFolder
        Call filePaths.Add(filePath)
    Next
    
    For Each filePath In filePaths
        Call fso.CreateTextFile(filePath)
    Next
    
    Set result = ListFiles(pathFolder)
    Assert.AreEqual filePathsInPathFolder.Count, result.Count
    For Each filePath In result
        Assert.IsTrue CollectionContains(filePathsInPathFolder, filePath)
    Next
    
    Set result = ListFiles(pathFolder, True)
    Assert.AreEqual filePathsUnderPathFolder.Count, result.Count
    For Each filePath In result
        Assert.IsTrue CollectionContains(filePathsUnderPathFolder, filePath)
    Next
    
    Set result = ListFiles(testRootPath, True)
    Assert.AreEqual filePaths.Count, result.Count
    For Each filePath In result
        Assert.IsTrue CollectionContains(filePaths, filePath)
    Next
    
    Dim filter As Object: Set filter = CreateObject("VBScript.RegExp")
    filter.Pattern = "\.(html|css)$"
    Dim webSourcePaths As Collection: Set webSourcePaths = New Collection
    For Each filePath In filePaths
        If filter.Test(filePath) Then
            Call webSourcePaths.Add(filePath)
        End If
    Next
    
    Set result = ListFiles(testRootPath, True, filter.Pattern)
    Assert.AreEqual webSourcePaths.Count, result.Count
    For Each filePath In result
        Assert.IsTrue CollectionContains(webSourcePaths, filePath)
    Next
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    GoTo FINALLY
FINALLY:
    Call fso.DeleteFolder(testRootPath)
End Sub
