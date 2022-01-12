Attribute VB_Name = "PublishAsForWord"
' This file is released under the CC0 1.0 Universal License.
' See the LICENSE.txt file or https://creativecommons.org/publicdomain/zero/1.0/ for details.
Option Private Module
Option Explicit

Public Sub PublishAs( _
    TargetDocument As Document, _
    FileName As String, _
    Optional FileFormat As Variant, _
    Optional LockComments As Variant, _
    Optional Password As Variant, _
    Optional AddToRecentFiles As Variant, _
    Optional WritePassword As Variant, _
    Optional ReadOnlyRecommended As Variant, _
    Optional EmbedTrueTypeFonts As Variant, _
    Optional SaveNativePictureFormat As Variant, _
    Optional SaveFormsData As Variant, _
    Optional SaveAsAOCELetter As Variant, _
    Optional Encoding As Variant, _
    Optional InsertLineBreaks As Variant, _
    Optional AllowSubstitutions As Variant, _
    Optional LineEnding As Variant, _
    Optional AddBiDiMarks As Variant, _
    Optional CompatibilityMode As Variant _
)
    Dim fileSystemObject As Object
    Set fileSystemObject = CreateObject("Scripting.FileSystemObject")
    Dim targetDocumentPath As String
    targetDocumentPath = TargetDocument.Path & Application.PathSeparator & TargetDocument.Name
    
    Dim temporaryDocumentPath As String
    temporaryDocumentPath = fileSystemObject.GetSpecialFolder(2) & Application.PathSeparator & Format$(Now, "yyyymmddhhnnss") & "_" & TargetDocument.Name
    
    Call TargetDocument.Save
    
    Call fileSystemObject.CopyFile(targetDocumentPath, temporaryDocumentPath)
    
    Dim publishedDocument As Document
    Set publishedDocument = Documents.Open(temporaryDocumentPath)
    Call publishedDocument.SaveAs2( _
        FileName:=FileName, _
        FileFormat:=FileFormat, _
        LockComments:=LockComments, _
        Password:=Password, _
        AddToRecentFiles:=AddToRecentFiles, _
        WritePassword:=WritePassword, _
        ReadOnlyRecommended:=ReadOnlyRecommended, _
        EmbedTrueTypeFonts:=EmbedTrueTypeFonts, _
        SaveNativePictureFormat:=SaveNativePictureFormat, _
        SaveFormsData:=SaveFormsData, _
        SaveAsAOCELetter:=SaveAsAOCELetter, _
        Encoding:=Encoding, _
        InsertLineBreaks:=InsertLineBreaks, _
        AllowSubstitutions:=AllowSubstitutions, _
        LineEnding:=LineEnding, _
        AddBiDiMarks:=AddBiDiMarks, _
        CompatibilityMode:=CompatibilityMode _
    )
    publishedDocument.Close
    Call fileSystemObject.DeleteFile(temporaryDocumentPath, True)
End Sub
