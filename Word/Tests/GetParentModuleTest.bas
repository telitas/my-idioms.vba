Attribute VB_Name = "GetParentModuleTest"
Option Explicit
Option Private Module

'@TestModule
'@Folder "Project.Tests"

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

'@TestMethod("GetParentModule")
Private Sub GetParentDocument_CorrectCall_Successed()
    Dim expect As Document
    Dim actual As Document
    
    Dim testDocument As Document
    Set testDocument = Documents.Add
    
    Set expect = testDocument
    Set actual = GetParentDocument(testDocument)
    Assert.AreSame expect, actual
    
    Set actual = GetParentDocument(expect.Sections(1))
    Assert.AreSame expect, actual
    
    Set actual = GetParentDocument(expect.Sections(1).Range)
    Assert.AreSame expect, actual
    
    On Error GoTo ERROR_1
    Call GetParentDocument(Application)
    Assert.Fail
    GoTo FINALLY
    
ERROR_1:
    Assert.AreEqual 5&, Err.Number
    Resume RESUME_1
RESUME_1:
   
FINALLY:
    Call testDocument.Close
End Sub

'@TestMethod("GetParentModule")
Private Sub GetParentSection_CorrectCall_Successed()
    Dim expect As Section
    Dim actual As Section
    
    Dim testDocument As Document
    Set testDocument = Documents.Add
    
    Set expect = testDocument.Sections(1)
    Set actual = GetParentSection(expect)
    Assert.AreEqual expect.Index, actual.Index
    
    Set actual = GetParentSection(expect)
    Assert.AreEqual expect.Index, actual.Index
    
    Set actual = GetParentSection(expect.Headers(wdHeaderFooterPrimary))
    Assert.AreEqual expect.Index, actual.Index
        
    On Error GoTo ERROR_1
    Call GetParentSection(Application)
    Assert.Fail
    GoTo FINALLY
    
ERROR_1:
    Assert.AreEqual 5&, Err.Number
    Resume RESUME_1
RESUME_1:
    
    On Error GoTo ERROR_2
    Call GetParentSection(testDocument)
    Assert.Fail
    GoTo FINALLY
    
ERROR_2:
    Assert.AreEqual 5&, Err.Number
    Resume RESUME_2
RESUME_2:

FINALLY:
    Call testDocument.Close
End Sub
