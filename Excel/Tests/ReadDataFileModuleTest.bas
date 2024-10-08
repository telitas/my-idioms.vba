Attribute VB_Name = "ReadDataFileModuleTest"
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

'@TestMethod("ReadDataFileViaQuery")
Private Sub ReadDataFile_CSV_PromoteHeaders_CorrectCall_Successed()
    Dim cols(3, 1) As String
    cols(0, 0) = "text_column": cols(0, 1) = "text"
    cols(1, 0) = "number_column": cols(1, 1) = "number"
    cols(2, 0) = "datetime_column": cols(2, 1) = "datetime"
    cols(3, 0) = "Int64_column": cols(3, 1) = "Int64"

    Dim testDataWorkbook As Workbook
    Set testDataWorkbook = Workbooks.Add
    Dim testDataworksheet As Worksheet
    Set testDataworksheet = testDataWorkbook.Worksheets(1)
    
    Dim expect As Variant
    expect = testDataworksheet.Range( _
        testDataworksheet.Cells(1, 1), _
        testDataworksheet.Cells(1 + 2, 1 + UBound(cols, 1)) _
    ).Value
    
    Dim i As Long
    For i = 0 To UBound(cols, 1)
        expect(1, 1 + i) = cols(i, 0)
    Next
    expect(2, 1) = "a": expect(2, 2) = 0: expect(2, 3) = #1/1/2000#: expect(2, 4) = 0
    expect(3, 1) = "z": expect(3, 2) = 1.1: expect(3, 3) = #12/31/9999 11:59:59 PM#: expect(3, 4) = 1
    
    Dim cellFormat As String
    For i = LBound(cols, 1) To UBound(cols, 1)
        Select Case cols(i, 1)
            Case "datetime"
                cellFormat = "yyyy/mm/dd hh:mm:ss"
            Case Else
                cellFormat = "@"
        End Select
        testDataworksheet.Range( _
            testDataworksheet.Cells(LBound(expect, 1), 1 + i), _
            testDataworksheet.Cells(UBound(expect, 1), 1 + i) _
        ).NumberFormatLocal = cellFormat
    Next
    testDataworksheet.Range( _
        testDataworksheet.Cells(1, 1), _
        testDataworksheet.Cells(UBound(expect, 1), UBound(expect, 2)) _
    ).Value = expect
    
    Dim testDataFilePath As String
    testDataFilePath = CreateObject("Scripting.FileSystemObject").GetSpecialFolder(2).Path & "\data.csv"
    Call testDataWorkbook.SaveAs(Filename:=testDataFilePath, FileFormat:=xlCSVUTF8)
    Call testDataWorkbook.Close(SaveChanges:=False)
    
    Dim actual As Variant
    
    On Error GoTo ERROR_1
    Dim testWorkbook As Workbook
    Set testWorkbook = Workbooks.Add
    
    Dim outputSheet As Worksheet
    
    Set outputSheet = ReadDataFile( _
        TargetWorkbook:=testWorkbook, _
        FilePath:=testDataFilePath, _
        FileType:="csv", _
        CodePage:=65001, _
        Columns:=cols _
    )
    actual = outputSheet.Range( _
        outputSheet.Cells(1, 1), _
        outputSheet.Cells(outputSheet.Cells(outputSheet.Rows.Count, 1).End(xlUp).Row, outputSheet.Cells(1, outputSheet.Columns.Count).End(xlToLeft).Column) _
    ).Value
    
    Dim j As Long
    For i = LBound(expect, 1) To UBound(expect, 1)
        For j = LBound(expect, 2) To UBound(expect, 2)
            Assert.IsTrue expect(i, j) = actual(i, j)
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
    Kill testDataFilePath
End Sub

'@TestMethod("ReadDataFileViaQuery")
Private Sub ReadDataFile_CSV_NoHeaders_CorrectCall_Successed()
    Dim cols(3, 1) As String
    cols(0, 0) = "text_column": cols(0, 1) = "text"
    cols(1, 0) = "number_column": cols(1, 1) = "number"
    cols(2, 0) = "datetime_column": cols(2, 1) = "datetime"
    cols(3, 0) = "Int64_column": cols(3, 1) = "Int64"

    Dim testDataWorkbook As Workbook
    Set testDataWorkbook = Workbooks.Add
    Dim testDataworksheet As Worksheet
    Set testDataworksheet = testDataWorkbook.Worksheets(1)
    
    Dim expect As Variant
    expect = testDataworksheet.Range( _
        testDataworksheet.Cells(1, 1), _
        testDataworksheet.Cells(1 + 2, 1 + UBound(cols, 1)) _
    ).Value
    
    Dim i As Long
    For i = 0 To UBound(cols, 1)
        expect(1, 1 + i) = cols(i, 0)
    Next
    expect(2, 1) = "a": expect(2, 2) = 0: expect(2, 3) = #1/1/2000#: expect(2, 4) = 0
    expect(3, 1) = "z": expect(3, 2) = 1.1: expect(3, 3) = #12/31/9999 11:59:59 PM#: expect(3, 4) = 1
    
    Dim cellFormat As String
    For i = LBound(cols, 1) To UBound(cols, 1)
        Select Case cols(i, 1)
            Case "datetime"
                cellFormat = "yyyy/mm/dd hh:mm:ss"
            Case Else
                cellFormat = "@"
        End Select
        testDataworksheet.Range( _
            testDataworksheet.Cells(LBound(expect, 1), 1 + i), _
            testDataworksheet.Cells(UBound(expect, 1), 1 + i) _
        ).NumberFormatLocal = cellFormat
    Next
    
    testDataworksheet.Range( _
        testDataworksheet.Cells(1, 1), _
        testDataworksheet.Cells(UBound(expect, 1), UBound(expect, 2)) _
    ).Value = expect
    Call testDataworksheet.Range("1:1").Delete
    
    Dim testDataFilePath As String
    testDataFilePath = CreateObject("Scripting.FileSystemObject").GetSpecialFolder(2).Path & "\data.csv"
    Call testDataWorkbook.SaveAs(Filename:=testDataFilePath, FileFormat:=xlCSVUTF8)
    Call testDataWorkbook.Close(SaveChanges:=False)
    
    Dim actual As Variant
    
    On Error GoTo ERROR_1
    Dim testWorkbook As Workbook
    Set testWorkbook = Workbooks.Add
    
    Dim outputSheet As Worksheet
    
    Set outputSheet = ReadDataFile( _
        TargetWorkbook:=testWorkbook, _
        FilePath:=testDataFilePath, _
        FileType:="csv", _
        CodePage:=65001, _
        Columns:=cols, _
        CsvPromoteHeaders:=False _
    )
    actual = outputSheet.Range( _
        outputSheet.Cells(1, 1), _
        outputSheet.Cells(outputSheet.Cells(outputSheet.Rows.Count, 1).End(xlUp).Row, outputSheet.Cells(1, outputSheet.Columns.Count).End(xlToLeft).Column) _
    ).Value
    
    Dim j As Long
    For i = LBound(expect, 1) To UBound(expect, 1)
        For j = LBound(expect, 2) To UBound(expect, 2)
            Assert.IsTrue expect(i, j) = actual(i, j)
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
    Kill testDataFilePath
End Sub

'@TestMethod("ReadDataFileViaQuery")
Private Sub ReadDataFile_CSV_ParseTSV_CorrectCall_Successed()
    Dim cols(3, 1) As String
    cols(0, 0) = "text_column": cols(0, 1) = "text"
    cols(1, 0) = "number_column": cols(1, 1) = "number"
    cols(2, 0) = "datetime_column": cols(2, 1) = "datetime"
    cols(3, 0) = "Int64_column": cols(3, 1) = "Int64"

    Dim testDataWorkbook As Workbook
    Set testDataWorkbook = Workbooks.Add
    Dim testDataworksheet As Worksheet
    Set testDataworksheet = testDataWorkbook.Worksheets(1)
    
    Dim expect As Variant
    expect = testDataworksheet.Range( _
        testDataworksheet.Cells(1, 1), _
        testDataworksheet.Cells(1 + 2, 1 + UBound(cols, 1)) _
    ).Value
    
    Dim i As Long
    For i = 0 To UBound(cols, 1)
        expect(1, 1 + i) = cols(i, 0)
    Next
    expect(2, 1) = "a": expect(2, 2) = 0: expect(2, 3) = #1/1/2000#: expect(2, 4) = 0
    expect(3, 1) = "z": expect(3, 2) = 1.1: expect(3, 3) = #12/31/9999 11:59:59 PM#: expect(3, 4) = 1
    
    Dim cellFormat As String
    For i = LBound(cols, 1) To UBound(cols, 1)
        Select Case cols(i, 1)
            Case "datetime"
                cellFormat = "yyyy/mm/dd hh:mm:ss"
            Case Else
                cellFormat = "@"
        End Select
        testDataworksheet.Range( _
            testDataworksheet.Cells(LBound(expect, 1), 1 + i), _
            testDataworksheet.Cells(UBound(expect, 1), 1 + i) _
        ).NumberFormatLocal = cellFormat
    Next
    
    testDataworksheet.Range( _
        testDataworksheet.Cells(1, 1), _
        testDataworksheet.Cells(UBound(expect, 1), UBound(expect, 2)) _
    ).Value = expect
    
    Dim testDataFilePath As String
    testDataFilePath = CreateObject("Scripting.FileSystemObject").GetSpecialFolder(2).Path & "\data.csv"
    Call testDataWorkbook.SaveAs(Filename:=testDataFilePath, FileFormat:=xlText)
    Call testDataWorkbook.Close(SaveChanges:=False)
    
    Dim actual As Variant
    
    On Error GoTo ERROR_1
    Dim testWorkbook As Workbook
    Set testWorkbook = Workbooks.Add
    
    Dim outputSheet As Worksheet
    
    Set outputSheet = ReadDataFile( _
        TargetWorkbook:=testWorkbook, _
        FilePath:=testDataFilePath, _
        FileType:="csv", _
        CodePage:=65001, _
        Columns:=cols, _
        CsvDelimiter:=vbTab _
    )
    actual = outputSheet.Range( _
        outputSheet.Cells(1, 1), _
        outputSheet.Cells(outputSheet.Cells(outputSheet.Rows.Count, 1).End(xlUp).Row, outputSheet.Cells(1, outputSheet.Columns.Count).End(xlToLeft).Column) _
    ).Value
    
    Dim j As Long
    For i = LBound(expect, 1) To UBound(expect, 1)
        For j = LBound(expect, 2) To UBound(expect, 2)
            Assert.IsTrue expect(i, j) = actual(i, j)
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
    Kill testDataFilePath
End Sub



'@TestMethod("ReadDataFileViaQuery")
Private Sub ReadDataFile_JSON_CorrectCall_Successed()
    Dim cols(3, 1) As String
    cols(0, 0) = "text_column": cols(0, 1) = "text"
    cols(1, 0) = "number_column": cols(1, 1) = "number"
    cols(2, 0) = "datetime_column": cols(2, 1) = "datetime"
    cols(3, 0) = "Int64_column": cols(3, 1) = "Int64"

    Dim testDataWorkbook As Workbook
    Set testDataWorkbook = Workbooks.Add
    Dim testDataworksheet As Worksheet
    Set testDataworksheet = testDataWorkbook.Worksheets(1)
    
    Dim expect As Variant
    expect = testDataworksheet.Range( _
        testDataworksheet.Cells(1, 1), _
        testDataworksheet.Cells(1 + 2, 1 + UBound(cols, 1)) _
    ).Value
    Call testDataWorkbook.Close(SaveChanges:=False)
    
    Dim i As Long
    For i = 0 To UBound(cols, 1)
        expect(1, 1 + i) = cols(i, 0)
    Next
    expect(2, 1) = "a": expect(2, 2) = 0: expect(2, 3) = #1/1/2000#: expect(2, 4) = 0
    expect(3, 1) = "z": expect(3, 2) = 1.1: expect(3, 3) = #12/31/9999 11:59:59 PM#: expect(3, 4) = 1
    
    Dim testDataFilePath As String
    testDataFilePath = CreateObject("Scripting.FileSystemObject").GetSpecialFolder(2).Path & "\data.json"

    Open testDataFilePath For Output As #1
    Print #1, "[{""text_column"": ""a"",""number_column"": 0,""datetime_column"": ""2000-01-01T00:00:00"",""Int64_column"": 0},{""text_column"": ""z"",""number_column"": 1.1,""datetime_column"": ""9999-12-31T23:59:59"",""Int64_column"": 1}]"
    Close #1
    
    Dim actual As Variant
    
    On Error GoTo ERROR_1
    Dim testWorkbook As Workbook
    Set testWorkbook = Workbooks.Add
    
    Dim outputSheet As Worksheet
    
    Set outputSheet = ReadDataFile( _
        TargetWorkbook:=testWorkbook, _
        FilePath:=testDataFilePath, _
        FileType:="json", _
        CodePage:=65001, _
        Columns:=cols _
    )
    actual = outputSheet.Range( _
        outputSheet.Cells(1, 1), _
        outputSheet.Cells(outputSheet.Cells(outputSheet.Rows.Count, 1).End(xlUp).Row, outputSheet.Cells(1, outputSheet.Columns.Count).End(xlToLeft).Column) _
    ).Value
    
    Dim j As Long
    For i = LBound(expect, 1) To UBound(expect, 1)
        For j = LBound(expect, 2) To UBound(expect, 2)
            Assert.IsTrue expect(i, j) = actual(i, j)
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
    Kill testDataFilePath
End Sub


'@TestMethod("ReadDataFileViaQuery")
Private Sub ReadDataFile_XML_CorrectCall_Successed()
    Dim cols(3, 1) As String
    cols(0, 0) = "text_column": cols(0, 1) = "text"
    cols(1, 0) = "number_column": cols(1, 1) = "number"
    cols(2, 0) = "datetime_column": cols(2, 1) = "datetime"
    cols(3, 0) = "Int64_column": cols(3, 1) = "Int64"

    Dim testDataWorkbook As Workbook
    Set testDataWorkbook = Workbooks.Add
    Dim testDataworksheet As Worksheet
    Set testDataworksheet = testDataWorkbook.Worksheets(1)
    
    Dim expect As Variant
    expect = testDataworksheet.Range( _
        testDataworksheet.Cells(1, 1), _
        testDataworksheet.Cells(1 + 2, 1 + UBound(cols, 1)) _
    ).Value
    Call testDataWorkbook.Close(SaveChanges:=False)
    
    Dim i As Long
    For i = 0 To UBound(cols, 1)
        expect(1, 1 + i) = cols(i, 0)
    Next
    expect(2, 1) = "a": expect(2, 2) = 0: expect(2, 3) = #1/1/2000#: expect(2, 4) = 0
    expect(3, 1) = "z": expect(3, 2) = 1.1: expect(3, 3) = #12/31/9999 11:59:59 PM#: expect(3, 4) = 1
    
    Dim testDataFilePath As String
    testDataFilePath = CreateObject("Scripting.FileSystemObject").GetSpecialFolder(2).Path & "\data.json"

    Open testDataFilePath For Output As #1
    Print #1, "<?xml version=""1.0"" encoding=""utf-8""?><rows><row><text_column>a</text_column><number_column>0</number_column><datetime_column>2000-01-01T00:00:00</datetime_column><Int64_column>0</Int64_column></row><row><text_column>z</text_column><number_column>1.1</number_column><datetime_column>9999-12-31T23:59:59</datetime_column><Int64_column>1</Int64_column></row></rows>"
    Close #1
    
    Dim actual As Variant
    
    On Error GoTo ERROR_1
    Dim testWorkbook As Workbook
    Set testWorkbook = Workbooks.Add
    
    Dim outputSheet As Worksheet
    
    Set outputSheet = ReadDataFile( _
        TargetWorkbook:=testWorkbook, _
        FilePath:=testDataFilePath, _
        FileType:="xml", _
        CodePage:=65001, _
        Columns:=cols _
    )
    actual = outputSheet.Range( _
        outputSheet.Cells(1, 1), _
        outputSheet.Cells(outputSheet.Cells(outputSheet.Rows.Count, 1).End(xlUp).Row, outputSheet.Cells(1, outputSheet.Columns.Count).End(xlToLeft).Column) _
    ).Value
    
    Dim j As Long
    For i = LBound(expect, 1) To UBound(expect, 1)
        For j = LBound(expect, 2) To UBound(expect, 2)
            Assert.IsTrue expect(i, j) = actual(i, j)
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
    Kill testDataFilePath
End Sub

'@TestMethod("ReadDataFileViaQuery")
Private Sub ReadDataFile_InvalidFileType_Failed()
    Dim cols(3, 1) As String
    cols(0, 0) = "text_column": cols(0, 1) = "text"
    cols(1, 0) = "number_column": cols(1, 1) = "number"
    cols(2, 0) = "datetime_column": cols(2, 1) = "datetime"
    cols(3, 0) = "Int64_column": cols(3, 1) = "Int64"
    
    On Error GoTo ERROR_1
    
    Dim testDataFilePath As String
    testDataFilePath = CreateObject("Scripting.FileSystemObject").GetSpecialFolder(2).Path & "\data.json"
    
    Dim testWorkbook As Workbook
    Set testWorkbook = Workbooks.Add
    
    Dim outputSheet As Worksheet
    
    Set outputSheet = ReadDataFile( _
        TargetWorkbook:=testWorkbook, _
        FilePath:=testDataFilePath, _
        FileType:="yaml", _
        CodePage:=65001, _
        Columns:=cols, _
        CsvDelimiter:=vbTab _
    )
    
    Assert.Fail "This should not be reached."
    
    GoTo FINALLY
ERROR_1:
    Assert.AreEqual 5&, Err.Number
    Resume RESUME_1
RESUME_1:

FINALLY:
    On Error Resume Next
    Call testWorkbook.Close(SaveChanges:=False)
    Kill testDataFilePath
End Sub
