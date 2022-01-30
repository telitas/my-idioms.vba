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

Private Const DebugObjectSheetName As String = "Sheet1"
Private Const DummyCodeString As String = _
"Private Sub Dummy()" & vbCrLf & _
"End Sub"
Private Const SkipTaskpaneAppsTest As Boolean = True

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
Private Sub GetPresentation_CorrectCall_Successed()
    Dim expect As Presentation
    Dim actual As Presentation
    
    Dim testPresentation As Presentation
    Set testPresentation = Presentations.Add
    
    Call testPresentation.Slides.AddSlide(1, testPresentation.SlideMaster.CustomLayouts(7))
    
    Set expect = testPresentation
    Set actual = GetPresentation(expect)
    Assert.AreSame expect, actual
    
    Set actual = GetPresentation(expect.Slides(1))
    Assert.AreSame expect, actual
    
    Set actual = GetPresentation(expect.Slides(1).Shapes.AddShape(msoShapeRectangle, 0, 0, 1, 1))
    Assert.AreSame expect, actual
    
    On Error GoTo ERROR_1
    Call GetPresentation(Application)
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
Private Sub GetSlide_CorrectCall_Successed()
    Dim expect As Slide
    Dim actual As Slide
    
    Dim testPresentation As Presentation
    Set testPresentation = Presentations.Add
    
    Set expect = testPresentation.Slides.AddSlide(1, testPresentation.SlideMaster.CustomLayouts(7))
    Set actual = GetSlide(expect)
    Assert.AreSame expect, actual
    
    Set actual = GetSlide(expect.Shapes.AddShape(msoShapeRectangle, 0, 0, 1, 1))
    Assert.AreSame expect, actual
            
    On Error GoTo ERROR_1
    Call GetSlide(Application)
    Assert.Fail
    GoTo FINALLY
    
ERROR_1:
    Assert.AreEqual 5&, Err.Number
    Resume RESUME_1
RESUME_1:
    
    On Error GoTo ERROR_2
    Call GetSlide(testPresentation)
    Assert.Fail
    GoTo FINALLY
    
ERROR_2:
    Assert.AreEqual 5&, Err.Number
    Resume RESUME_2
RESUME_2:

FINALLY:
    Call testPresentation.Close
End Sub
