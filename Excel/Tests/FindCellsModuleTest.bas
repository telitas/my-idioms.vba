Attribute VB_Name = "FindCellsModuleTest"
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
'@Ignore EmptyMethod
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
'@Ignore EmptyMethod
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

Private Sub CloseBook(ByVal TargetBook As Workbook)
    Dim currentDisplayAlerts As Boolean
    currentDisplayAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Call TargetBook.Close
    Application.DisplayAlerts = currentDisplayAlerts
End Sub

Private Function CreateTestBookToFindCellsTest() As Workbook
    Dim testBook As Workbook
    Set testBook = Workbooks.Add
    Do While testBook.Worksheets.Count < 3
        testBook.Worksheets.Add
    Loop
    
    Dim sheet As Worksheet
    Const findString As String = "What"
    Dim i As Long
    For Each sheet In testBook.Worksheets
        For i = 1 To 10
            sheet.Cells(i, i).Value = findString
            sheet.Cells(1, i).Value = findString
            sheet.Cells(i, 1).Value = findString
        Next
    Next
    Set CreateTestBookToFindCellsTest = testBook
End Function

'@TestMethod
Public Sub FindCellsInRange_Successed()
    Dim testBook As Workbook
    Set testBook = CreateTestBookToFindCellsTest
    Dim testSheet As Worksheet
    Set testSheet = testBook.Worksheets(1)
    Dim testRange As Range
    Set testRange = testSheet.Range("A1:E5")
    
    Dim expect As Long
    expect = 13
    Dim actual As Collection
    Set actual = FindCellsInRange( _
        TargetRange:=testRange, _
        What:="What", _
        LookIn:=xlValues, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=True, _
        MatchByte:=True, _
        SearchFormat:=False _
    )
    Assert.AreEqual actual.Count, expect
    Call CloseBook(testBook)
End Sub

'@TestMethod
Public Sub FindCellsInWorksheet_Successed()
    Dim testBook As Workbook
    Set testBook = CreateTestBookToFindCellsTest
    Dim testSheet As Worksheet
    Set testSheet = testBook.Worksheets(1)
    
    Dim expect As Long
    expect = 28
    Dim actual As Collection
    Set actual = FindCellsInWorksheet( _
        TargetWorksheet:=testSheet, _
        What:="What", _
        LookIn:=xlValues, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=True, _
        MatchByte:=True, _
        SearchFormat:=False _
    )
    Assert.AreEqual actual.Count, expect
    Call CloseBook(testBook)
End Sub

'@TestMethod
Public Sub FindCellsInWorkbook_Successed()
    Dim testBook As Workbook
    Set testBook = CreateTestBookToFindCellsTest
    
    Dim expect As Long
    expect = 28 * testBook.Worksheets.Count
    Dim actual As Collection
    Set actual = FindCellsInWorkbook( _
        TargetWorkbook:=testBook, _
        What:="What", _
        LookIn:=xlValues, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=True, _
        MatchByte:=True, _
        SearchFormat:=False _
    )
    Assert.AreEqual actual.Count, expect
    Call CloseBook(testBook)
End Sub
