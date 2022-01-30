Attribute VB_Name = "FindRangesModuleTest"
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
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("FindRanges")
Public Sub FindRanges_CorrectCall_Successed()
    Dim testDocument As Document
    Set testDocument = Documents.Add
    
    Dim expectCount As Long
    expectCount = 15
    Dim i As Long
    With testDocument.Content
        For i = 1 To expectCount
            Call .InsertAfter("Paragraph " & i)
            Call .InsertParagraphAfter
        Next
        Call .Paragraphs(11).Range.InsertBreak(wdSectionBreakContinuous)
    End With
    
    Dim currentSelectionRange As Range
    Set currentSelectionRange = Selection.Range
    Dim find_ As Find
    Set find_ = Selection.Find
    With find_
        .ClearFormatting
        .Text = "Paragraph [0-9]{1,2}"
        .MatchFuzzy = False
        .MatchWildcards = True
    End With
    
    Dim actual As Collection
    Set actual = FindRanges(find_)
    Assert.AreEqual expectCount, actual.Count
    Assert.AreEqual Selection.Start, currentSelectionRange.Start
    Assert.AreEqual Selection.End, currentSelectionRange.End
    
    Dim findRange As Range
    Dim currentRange As Range
    
    Set findRange = testDocument.Content
    Set currentRange = findRange.Parent.Range(findRange.Start, findRange.End)
    Set find_ = findRange.Find
    With find_
        .ClearFormatting
        .Text = "Paragraph [0-9]{1,2}"
        .MatchFuzzy = False
        .MatchWildcards = True
    End With
    
    Set actual = FindRanges(find_)
    Assert.AreEqual expectCount, actual.Count
    Assert.AreEqual findRange.Start, currentRange.Start
    Assert.AreEqual findRange.End, currentRange.End
    
    Set findRange = testDocument.Content.Sections(1).Range
    Set currentRange = findRange.Parent.Range(findRange.Start, findRange.End)
    Set find_ = findRange.Find
    With find_
        .ClearFormatting
        .Text = "Paragraph [0-9]{1,2}"
        .MatchFuzzy = False
        .MatchWildcards = True
    End With
    
    expectCount = 10
    Set actual = FindRanges(find_)
    Assert.AreEqual expectCount, actual.Count
    Assert.AreEqual findRange.Start, currentRange.Start
    Assert.AreEqual findRange.End, currentRange.End
    
    Set findRange = testDocument.Content.Paragraphs(3).Range
    Set currentRange = findRange.Parent.Range(findRange.Start, findRange.End)
    Set find_ = findRange.Find
    With find_
        .ClearFormatting
        .Text = "Paragraph [0-9]{1,2}"
        .MatchFuzzy = False
        .MatchWildcards = True
    End With
    
    expectCount = 1
    Set actual = FindRanges(find_)
    Assert.AreEqual expectCount, actual.Count
    Assert.AreEqual findRange.Start, currentRange.Start
    Assert.AreEqual findRange.End, currentRange.End
    
    Call testDocument.Close(wdDoNotSaveChanges)
End Sub
