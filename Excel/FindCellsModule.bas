Attribute VB_Name = "FindCellsModule"
' This file is released under the CC0 1.0 Universal License.
' See the LICENSE.txt file or https://creativecommons.org/publicdomain/zero/1.0/ for details.
Option Explicit
Option Private Module

Public Function FindCellsInRange( _
    ByVal TargetRange As Range, _
    ByVal What As Variant, _
    Optional ByVal LookIn As Variant, _
    Optional ByVal LookAt As Variant, _
    Optional ByVal SearchOrder As Variant, _
    Optional ByVal SearchDirection As Variant, _
    Optional ByVal MatchCase As Variant, _
    Optional ByVal MatchByte As Variant, _
    Optional ByVal SearchFormat As Variant _
) As Collection
    Dim listedCells As Collection: Set listedCells = New Collection
    Dim foundCell As Range
    With TargetRange
        Set foundCell = .Find( _
            What:=What, _
            LookIn:=LookIn, _
            LookAt:=LookAt, _
            SearchOrder:=SearchOrder, _
            SearchDirection:=SearchDirection, _
            MatchCase:=MatchCase, _
            MatchByte:=MatchByte, _
            SearchFormat:=SearchFormat _
        )
        If foundCell Is Nothing Then
            Set FindCellsInRange = listedCells
            Exit Function
        End If
        
        Call listedCells.Add(foundCell)
        Dim firstFoundCellAddress As String
        firstFoundCellAddress = foundCell.Address
        Do
            Set foundCell = .FindNext(foundCell)
            If foundCell.Address = firstFoundCellAddress Then
                Exit Do
            Else
                Call listedCells.Add(foundCell)
            End If
        Loop
    End With
    
    Set FindCellsInRange = listedCells
End Function

Public Function FindCellsInWorksheet( _
    ByVal TargetWorksheet As Worksheet, _
    ByVal What As Variant, _
    Optional ByVal LookIn As Variant, _
    Optional ByVal LookAt As Variant, _
    Optional ByVal SearchOrder As Variant, _
    Optional ByVal SearchDirection As Variant, _
    Optional ByVal MatchCase As Variant, _
    Optional ByVal MatchByte As Variant, _
    Optional ByVal SearchFormat As Variant _
) As Collection
    Set FindCellsInWorksheet = FindCellsInRange( _
        TargetRange:=TargetWorksheet.Cells, _
        What:=What, _
        LookIn:=LookIn, _
        LookAt:=LookAt, _
        SearchOrder:=SearchOrder, _
        SearchDirection:=SearchDirection, _
        MatchCase:=MatchCase, _
        MatchByte:=MatchByte, _
        SearchFormat:=SearchFormat _
    )
End Function

Public Function FindCellsInWorkbook( _
    ByVal TargetWorkbook As Workbook, _
    ByVal What As Variant, _
    Optional ByVal LookIn As Variant, _
    Optional ByVal LookAt As Variant, _
    Optional ByVal SearchOrder As Variant, _
    Optional ByVal SearchDirection As Variant, _
    Optional ByVal MatchCase As Variant, _
    Optional ByVal MatchByte As Variant, _
    Optional ByVal SearchFormat As Variant _
) As Collection
    Dim listedCells As Collection: Set listedCells = New Collection
    Dim sheet As Worksheet
    Dim cell_ As Range
    For Each sheet In TargetWorkbook.Worksheets
        For Each cell_ In FindCellsInWorksheet( _
            TargetWorksheet:=sheet, _
            What:=What, _
            LookIn:=LookIn, _
            LookAt:=LookAt, _
            SearchOrder:=SearchOrder, _
            SearchDirection:=SearchDirection, _
            MatchCase:=MatchCase, _
            MatchByte:=MatchByte, _
            SearchFormat:=SearchFormat _
        )
            Call listedCells.Add(cell_)
        Next
    Next
    Set FindCellsInWorkbook = listedCells
End Function
