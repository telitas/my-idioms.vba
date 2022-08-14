Attribute VB_Name = "DictionaryUtilsTest"
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

'@TestMethod("DictionaryEquals")
Private Sub DictionaryEquals_CorrectCall_Successed()
    Dim One As Object
    Dim Another As Object
    
    Dim obj1 As Object: Set obj1 = New Collection
    Dim obj2 As Object: Set obj2 = New Collection
    
    Set One = CreateObject("Scripting.Dictionary")
    Set Another = CreateObject("Scripting.Dictionary")
    Assert.IsTrue DictionaryEquals(Nothing, Nothing)
    Assert.IsFalse DictionaryEquals(One, Nothing)
    Assert.IsFalse DictionaryEquals(Nothing, Another)
    Assert.IsTrue DictionaryEquals(One, One)
    Assert.IsTrue DictionaryEquals(Another, Another)
    Assert.IsFalse DictionaryEquals(obj1, obj2)
    Assert.IsFalse DictionaryEquals(obj1, obj1)
    
    Call One.Add("a", 1)
    Assert.IsFalse DictionaryEquals(One, Another)
    Call Another.Add("a", 1)
    Assert.IsTrue DictionaryEquals(One, Another)
    Call One.Add("b", 2)
    Assert.IsFalse DictionaryEquals(One, Another)
    Call Another.Add("b", 2)
    Assert.IsTrue DictionaryEquals(One, Another)
    Call One.Add("c", 3)
    Call Another.Add("c", 4)
    Assert.IsFalse DictionaryEquals(One, Another)
    
    Set One = CreateObject("Scripting.Dictionary")
    Set Another = CreateObject("Scripting.Dictionary")
    Call One.Add("a", 1)
    Call Another.Add("a", obj1)
    Assert.IsFalse DictionaryEquals(One, Another)
    Assert.IsFalse DictionaryEquals(Another, One)

    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    GoTo FINALLY
FINALLY:
End Sub
