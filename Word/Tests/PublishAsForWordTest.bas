Attribute VB_Name = "PublishAsForWordTest"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

#Const LateBind = LateBindTests

#If LateBind Then
    Private Assert As Object
    Private Fakes As Object
#Else
    Private Assert As Rubberduck.AssertClass
    Private Fakes As Rubberduck.FakesProvider
#End If

Private TestDirectory As String
Private fileSystemObject As Object

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
    Set fileSystemObject = CreateObject("Scripting.FileSystemObject")
    TestDirectory = fileSystemObject.GetSpecialFolder(2) & Application.PathSeparator & "PublishAsForWordTest"
    If fileSystemObject.FolderExists(TestDirectory) Then
        Call fileSystemObject.DeleteFolder(TestDirectory, True)
    End If
    Call fileSystemObject.CreateFolder(TestDirectory)
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
    Call fileSystemObject.DeleteFolder(TestDirectory, True)
    Set fileSystemObject = Nothing
End Sub

'@TestMethod("PublishAs")
Public Sub PublishAs_CorrectCall_Successed()
    Dim testDocument As Document
    Set testDocument = Documents.Add
    Dim testDocumentPath As String
    testDocumentPath = TestDirectory & Application.PathSeparator & "TestDocument.docm"
    Dim publishedDocumentPath As String
    publishedDocumentPath = Replace(testDocumentPath, ".docm", ".docx")
    Call testDocument.SaveAs(FileName:=testDocumentPath, FileFormat:=wdFormatXMLDocumentMacroEnabled)
    
    Dim currentDisplayAlerts As Boolean
    currentDisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    
    Call PublishAs(TargetDocument:=testDocument, FileName:=publishedDocumentPath, FileFormat:=wdFormatDocumentDefault)
    
    Application.DisplayAlerts = currentDisplayAlerts
    
    Assert.IsTrue fileSystemObject.FileExists(testDocumentPath)
    Assert.IsTrue fileSystemObject.FileExists(publishedDocumentPath)
    
    Call testDocument.Close
    Call fileSystemObject.DeleteFile(publishedDocumentPath, True)
    Call fileSystemObject.DeleteFile(testDocumentPath, True)
End Sub
