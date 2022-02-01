Attribute VB_Name = "GetParentModuleTest"
'@TestModule
'@Folder "VBAProject.Tests"
'@IgnoreModule RedundantByRefModifier, ObsoleteCallStatement, FunctionReturnValueDiscarded, FunctionReturnValueAlwaysDiscarded
'@IgnoreModule IndexedDefaultMemberAccess, ImplicitDefaultMemberAccess, IndexedUnboundDefaultMemberAccess, DefaultMemberRequired
'WhitelistedIdentifiers i, j
Option Explicit
Option Private Module

'@Ignore VariableNotUsed
Private Assert As Object
'@Ignore VariableNotUsed
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
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

'@TestMethod("GetParentModule")
Private Sub GetParentPresentation_CorrectCall_Successed()
    Dim expect As Presentation
    Dim actual As Presentation
    
    Dim testPresentation As Presentation
    Set testPresentation = Presentations.Add
    
    Call testPresentation.Slides.AddSlide(1, testPresentation.SlideMaster.CustomLayouts(7))
    
    Set expect = testPresentation
    Set actual = GetParentPresentation(expect)
    Assert.AreSame expect, actual
    
    Set actual = GetParentPresentation(expect.Slides(1))
    Assert.AreSame expect, actual
    
    Set actual = GetParentPresentation(expect.Slides(1).Shapes.AddShape(msoShapeRectangle, 0, 0, 1, 1))
    Assert.AreSame expect, actual
    
    On Error GoTo ERROR_1
    Call GetParentPresentation(Application)
    Assert.Fail
    GoTo FINALLY
    
ERROR_1:
    Assert.AreEqual 5&, Err.Number
    Resume RESUME_1
RESUME_1:
   
FINALLY:
    Call testPresentation.Close
End Sub

'@TestMethod("GetParentModule")
Private Sub GetParentSlide_CorrectCall_Successed()
    Dim expect As Slide
    Dim actual As Slide
    
    Dim testPresentation As Presentation
    Set testPresentation = Presentations.Add
    
    Set expect = testPresentation.Slides.AddSlide(1, testPresentation.SlideMaster.CustomLayouts(7))
    Set actual = GetParentSlide(expect)
    Assert.AreSame expect, actual
    
    Set actual = GetParentSlide(expect.Shapes.AddShape(msoShapeRectangle, 0, 0, 1, 1))
    Assert.AreSame expect, actual
            
    On Error GoTo ERROR_1
    Call GetParentSlide(Application)
    Assert.Fail
    GoTo FINALLY
    
ERROR_1:
    Assert.AreEqual 5&, Err.Number
    Resume RESUME_1
RESUME_1:
    
    On Error GoTo ERROR_2
    Call GetParentSlide(testPresentation)
    Assert.Fail
    GoTo FINALLY
    
ERROR_2:
    Assert.AreEqual 5&, Err.Number
    Resume RESUME_2
RESUME_2:

FINALLY:
    Call testPresentation.Close
End Sub
