Attribute VB_Name = "RecalcSelectedRangeModuleTest"
'@TestModule
'@Folder "VBAProject.Tests"


Option Explicit
Option Private Module

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

'@TestMethod("ReadDataFileViaQuery")
Private Sub RecalcSelectedRange_CorrectCall_Successed()
    On Error GoTo ERROR_1
    Dim testWorkbook As Workbook
    Set testWorkbook = Workbooks.Add
    
    Dim testSheet As Worksheet
    Set testSheet = testWorkbook.Worksheets.Add
    
    Const rowMax As Long = 3
    Const colMax As Long = 4
    
    Dim r As Long
    Dim c As Long
    
    For r = 1 To rowMax
        For c = 1 To colMax
            With testSheet.Cells(r, c)
                .NumberFormat = "@"
                .Formula = "=" & r & "*" & c
                .NumberFormat = "#"
            End With
        Next
    Next
    
    testSheet.Range(testSheet.Cells(1, 1), testSheet.Cells(rowMax, colMax)).Select
    RecalcSelectedRange
    
    For r = 1 To rowMax
        For c = 1 To colMax
            Assert.AreEqual testSheet.Cells(r, c).Value, CDbl(r * c)
        Next
    Next
    
    GoTo FINALLY
ERROR_1:
    Assert.Fail "Error raised: " & Err.Number & ", " & Err.Description
    Resume RESUME_1
RESUME_1:

FINALLY:
    On Error Resume Next
    Call testWorkbook.Close(SaveChanges:=False)
End Sub
