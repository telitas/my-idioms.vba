Attribute VB_Name = "PublishAsForExcel"
' This file is released under the CC0 1.0 Universal License.
' See the LICENSE.txt file or https://creativecommons.org/publicdomain/zero/1.0/ for details.
Option Private Module
Option Explicit

Public Sub PublishAs( _
    TargetWorkbook As Workbook, _
    Filename As String, _
    Optional FileFormat As Variant, _
    Optional Password As Variant, _
    Optional WriteResPassword As Variant, _
    Optional ReadOnlyRecommended As Variant, _
    Optional CreateBackup As Variant, _
    Optional AccessMode As XlSaveAsAccessMode, _
    Optional ConflictResolution As XlSaveConflictResolution, _
    Optional AddToMru As Variant, _
    Optional TextCodepage As Variant, _
    Optional TextVisualLayout As Variant, _
    Optional Locale As Variant _
)
    Dim fileSystemObject As Object
    Set fileSystemObject = CreateObject("Scripting.FileSystemObject")
    Dim targetWorkbookPath As String
    targetWorkbookPath = TargetWorkbook.Path & Application.PathSeparator & TargetWorkbook.Name
    
    Dim temporaryWorkbookPath As String
    temporaryWorkbookPath = fileSystemObject.GetSpecialFolder(2) & Application.PathSeparator & Format$(Now, "yyyymmddhhnnss") & "_" & TargetWorkbook.Name
    
    Call TargetWorkbook.Save
    
    Call fileSystemObject.CopyFile(targetWorkbookPath, temporaryWorkbookPath)
    
    Dim publishedWorkbook As Workbook
    Set publishedWorkbook = Workbooks.Open(temporaryWorkbookPath)
    Call publishedWorkbook.SaveAs( _
        Filename:=Filename, _
        FileFormat:=FileFormat, _
        Password:=Password, _
        WriteResPassword:=WriteResPassword, _
        ReadOnlyRecommended:=ReadOnlyRecommended, _
        CreateBackup:=CreateBackup, _
        AccessMode:=AccessMode, _
        ConflictResolution:=ConflictResolution, _
        AddToMru:=AddToMru, _
        TextCodepage:=TextCodepage, _
        TextVisualLayout:=TextVisualLayout, _
        Local:=Locale _
    )
    publishedWorkbook.Close
    Call fileSystemObject.DeleteFile(temporaryWorkbookPath, True)
End Sub
