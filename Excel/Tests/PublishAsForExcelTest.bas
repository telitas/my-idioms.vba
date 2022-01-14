Attribute VB_Name = "PublishAsForExcelTest"
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
    '@Ignore VariableNotUsed
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
    TestDirectory = fileSystemObject.GetSpecialFolder(2) & Application.PathSeparator & "PublishAsForExcelTest"
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
    Dim testWorkbook As Workbook
    Set testWorkbook = Workbooks.Add
    Dim testWorkbookPath As String
    testWorkbookPath = TestDirectory & Application.PathSeparator & "TestBook.xlsm"
    Dim publishedWorkbookPath As String
    publishedWorkbookPath = Replace(testWorkbookPath, ".xlsm", ".xlsx")
    Call testWorkbook.SaveAs(Filename:=testWorkbookPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled)
    
    Dim currentDisplayAlerts As Boolean
    currentDisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    
    Call PublishAs(TargetWorkbook:=testWorkbook, Filename:=publishedWorkbookPath, FileFormat:=xlWorkbookDefault)
    
    Application.DisplayAlerts = currentDisplayAlerts
    
    Assert.IsTrue fileSystemObject.FileExists(testWorkbookPath)
    Assert.IsTrue fileSystemObject.FileExists(publishedWorkbookPath)
    
    Call testWorkbook.Close
    Call fileSystemObject.DeleteFile(publishedWorkbookPath, True)
    Call fileSystemObject.DeleteFile(testWorkbookPath, True)
End Sub
