Attribute VB_Name = "GetParentModuleTest"
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

'@TestMethod("GetParentModule")
Private Sub GetWorkbook_CorrectCall_Successed()
    Dim expect As Workbook
    Dim actual As Workbook
    
    Dim testWorkbook As Workbook
    Set testWorkbook = Workbooks.Add
    
    Set expect = testWorkbook
    Set actual = GetWorkbook(testWorkbook)
    Assert.AreSame expect, actual
    
    Set actual = GetWorkbook(expect.Worksheets(1))
    Assert.AreSame expect, actual
    
    Set actual = GetWorkbook(expect.Worksheets(1).Cells(1, 1))
    Assert.AreSame expect, actual
    
    On Error GoTo ERROR_1
    Call GetWorkbook(Application)
    Assert.Fail
    GoTo FINALLY
    
ERROR_1:
    Assert.AreEqual 5&, Err.Number
    Resume RESUME_1
RESUME_1:

FINALLY:
    Call testWorkbook.Close
End Sub

'@TestMethod("GetParentModule")
Private Sub GetWorksheet_CorrectCall_Successed()
    Dim expect As Worksheet
    Dim actual As Worksheet
        
    Dim testWorkbook As Workbook
    Set testWorkbook = Workbooks.Add
    
    Set expect = testWorkbook.Worksheets(1)
    Set actual = GetWorksheet(expect)
    Assert.AreSame expect, actual
    
    Set actual = GetWorksheet(expect.Cells(1, 1))
    Assert.AreSame expect, actual
            
    On Error GoTo ERROR_1
    Call GetWorksheet(Application)
    Assert.Fail
    GoTo FINALLY
    
ERROR_1:
    Assert.AreEqual 5&, Err.Number
    Resume RESUME_1
RESUME_1:
    
    On Error GoTo ERROR_2
    Call GetWorksheet(ThisWorkbook)
    Assert.Fail
    GoTo FINALLY
    
ERROR_2:
    Assert.AreEqual 5&, Err.Number
    Resume RESUME_2
RESUME_2:
   
FINALLY:
    Call testWorkbook.Close
End Sub
