Attribute VB_Name = "ReadDataFileModule"
' This file is released under the CC0 1.0 Universal License.
' See the LICENSE.txt file or https://creativecommons.org/publicdomain/zero/1.0/ for details.
Option Private Module
Option Explicit

Public Function ReadDataFile( _
    ByVal TargetWorkbook As Workbook, _
    ByVal FilePath As String, _
    ByVal FileType As String, _
    ByVal CodePage As Long, _
    ByRef Columns() As String, _
    Optional ByVal CsvDelimiter As String = ",", _
    Optional ByVal CsvPromoteHeaders As Boolean = True _
) As Worksheet
    Dim readDataQueryName As String
    Dim query As WorkbookQuery
    
    Randomize
    Do
        readDataQueryName = "__ReadCsvFileViaQuery_" & Int(Rnd() * (2 ^ 31) - 1)
        For Each query In TargetWorkbook.Queries
            If readDataQueryName = query.Name Then
                readDataQueryName = vbNullString
            End If
        Next
    Loop While readDataQueryName = vbNullString
    
    Dim columnDefinition As String
    columnDefinition = "{"
    Dim i As Long
    Dim typeString As String
    For i = LBound(Columns) To UBound(Columns)
        Select Case Columns(i, 1)
            Case "any", _
                "anynonnull", _
                "binary", _
                "date", _
                "datetime", _
                "datetimezone", _
                "duration", _
                "function", _
                "list", _
                "logical", _
                "none", _
                "null", _
                "number", _
                "record", _
                "table", _
                "text", _
                "time", _
                "type":
                typeString = "type " + Columns(i, 1)
            Case Else
                typeString = Columns(i, 1) & ".Type"
        End Select
        columnDefinition = columnDefinition & _
            "{""" & Columns(i, 0) + """, " & _
            vbNullString & typeString & "}"
        If i < UBound(Columns) Then
            columnDefinition = columnDefinition & ", "
        End If
    Next
    columnDefinition = columnDefinition + "}"
    
    Dim queryFormula As String
    Dim columnNames As String
    Select Case FileType
        Case "csv"
            columnNames = "{"
            For i = LBound(Columns) To UBound(Columns)
                columnNames = columnNames & """" & Columns(i, 0) & """"
                If i < UBound(Columns) Then
                    columnNames = columnNames & ", "
                End If
            Next
            columnNames = columnNames & "}"
            queryFormula = "Csv.Document(File.Contents(""" & FilePath & """), [Delimiter=""" & CsvDelimiter & """, Columns=" & columnNames & ", Encoding=" & CodePage & "])"
            If CsvPromoteHeaders Then
                queryFormula = "Table.PromoteHeaders(" & queryFormula & ", [PromoteAllScalars=true])"
            End If
        Case "json"
            queryFormula = "Table.FromRecords(Json.Document(File.Contents(""" & FilePath & """), " & CodePage & "))"
        Case "xml"
            queryFormula = "Xml.Tables(File.Contents(""" & FilePath & """), [Encoding=" & CodePage & "]){0}[Table]"
        Case Else
            Call Err.Raise(Number:=5, Description:="FileType must be in ""csv"", ""json"" or ""xml"".")
    End Select
    queryFormula = "Table.TransformColumnTypes(" & queryFormula & ", " & columnDefinition & ")"
    
    Dim readDataQuery As WorkbookQuery
    Set readDataQuery = TargetWorkbook.Queries.Add( _
        Name:=readDataQueryName, _
        Formula:=queryFormula _
    )
    
    Dim outputSheet As Worksheet
    Set outputSheet = TargetWorkbook.Worksheets.Add
    
    Dim listObject As listObject
    Set listObject = outputSheet.ListObjects.Add( _
        SourceType:=0, _
        Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & readDataQuery.Name & ";Extended Properties=""""", _
        Destination:=outputSheet.Range("$A$1") _
    )
    With listObject.QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & readDataQuery.Name & "]")
        .RowNumbers = False
    End With
    
    Dim queryConnString As String
    queryConnString = listObject.QueryTable.Connection
    Dim readDataConnection As WorkbookConnection
    Dim conn As WorkbookConnection
    Dim connString As String
    For Each conn In TargetWorkbook.Connections
        connString = vbNullString
        On Error Resume Next
        connString = conn.OLEDBConnection.Connection
        On Error GoTo 0
        If connString = queryConnString Then
            Set readDataConnection = conn
        End If
    Next
    
    Call listObject.QueryTable.Refresh(BackgroundQuery:=False)
    
    listObject.TableStyle = vbNullString
    Call listObject.Unlist
    Call readDataQuery.Delete
    Call readDataConnection.Delete
    
    Set ReadDataFile = outputSheet
End Function
