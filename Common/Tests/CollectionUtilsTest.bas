Attribute VB_Name = "CollectionUtilsTest"
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

'@TestMethod("CollectionContains")
Private Sub CollectionContains_CorrectCall_Successed()
    Dim Target As Collection
    Dim value As Variant
    
    Set Target = New Collection
    Assert.IsFalse CollectionContains(Target, 1)
    Call Target.Add(1)
    Assert.IsTrue CollectionContains(Target, 1)
    Assert.IsFalse CollectionContains(Target, 2)
    Call Target.Add(2)
    Assert.IsTrue CollectionContains(Target, 1)
    Assert.IsTrue CollectionContains(Target, 2)
    
    Dim obj1 As Object: Set obj1 = New Collection
    Dim obj2 As Object: Set obj2 = New Collection
    
    Set Target = New Collection
    Call Target.Add(obj1)
    Assert.IsFalse CollectionContains(Target, Nothing)
    Assert.IsTrue CollectionContains(Target, obj1)
    Assert.IsFalse CollectionContains(Target, obj2)
    Assert.IsFalse CollectionContains(Target, 1)
    
    Set Target = New Collection
    Call Target.Add(1)
    Assert.IsFalse CollectionContains(Target, obj1)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    GoTo FINALLY
FINALLY:
End Sub

'@TestMethod("CollectionContains")
Private Sub CollectionContains_IllegalCall_Error()
    On Error GoTo CATCH
    
    On Error GoTo ERROR_RAISED
    Call CollectionContains(Nothing, 1)
    
    Assert.Fail
    GoTo FINALLY
    
ERROR_RAISED:
    Assert.AreEqual CLng(5), Err.Number
    Resume FINALLY
    
CATCH:
    Assert.Fail Err.Description
    GoTo FINALLY
FINALLY:
End Sub

'@TestMethod("CollectionEqualsAsList")
Private Sub CollectionEqualsAsList_CorrectCall_Successed()
    Dim One As Collection
    Dim Another As Collection
    
    Set One = New Collection
    Set Another = New Collection
    Assert.IsTrue CollectionEqualsAsList(Nothing, Nothing)
    
    Assert.IsFalse CollectionEqualsAsList(One, Nothing)
    Assert.IsFalse CollectionEqualsAsList(Nothing, Another)
    
    Assert.IsTrue CollectionEqualsAsList(One, One)
    Assert.IsTrue CollectionEqualsAsList(Another, Another)
    
    Assert.IsTrue CollectionEqualsAsList(One, Another)
    
    Call One.Add(1)
    Assert.IsFalse CollectionEqualsAsList(One, Another)
    Call Another.Add(1)
    Assert.IsTrue CollectionEqualsAsList(One, Another)
    Call One.Add(2)
    Assert.IsFalse CollectionEqualsAsList(One, Another)
    Call Another.Add(2)
    Assert.IsTrue CollectionEqualsAsList(One, Another)
    
    Set One = New Collection
    Call One.Add(1)
    Set Another = New Collection
    Call Another.Add(2)
    Assert.IsFalse CollectionEqualsAsList(One, Another)
    
    Set One = New Collection
    Call One.Add(1)
    Call One.Add(2)
    Set Another = New Collection
    Call Another.Add(2)
    Call Another.Add(1)
    Assert.IsFalse CollectionEqualsAsList(One, Another)
    
    Set One = New Collection
    Call One.Add(1)
    Call One.Add(1)
    Call One.Add(2)
    Set Another = New Collection
    Call Another.Add(1)
    Call Another.Add(2)
    Call Another.Add(2)
    Assert.IsFalse CollectionEqualsAsList(One, Another)
    
    Dim obj1 As Object: Set obj1 = New Collection
    Dim obj2 As Object: Set obj2 = New Collection
    
    Set One = New Collection
    Call One.Add(obj1)
    Set Another = New Collection
    Call Another.Add(obj1)
    Assert.IsTrue CollectionEqualsAsList(One, Another)
    
    Set One = New Collection
    Call One.Add(obj1)
    Set Another = New Collection
    Call Another.Add(obj2)
    Assert.IsFalse CollectionEqualsAsList(One, Another)
    
    Set One = New Collection
    Call One.Add(obj1)
    Set Another = New Collection
    Call Another.Add(1)
    Assert.IsFalse CollectionEqualsAsList(One, Another)
    
    Set One = New Collection
    Call One.Add(1)
    Set Another = New Collection
    Call Another.Add(obj2)
    Assert.IsFalse CollectionEqualsAsList(One, Another)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    GoTo FINALLY
FINALLY:
End Sub

'@TestMethod("CollectionEqualsAsSet")
Private Sub CollectionEqualsAsSet_CorrectCall_Successed()
    Dim One As Collection
    Dim Another As Collection
    
    Set One = New Collection
    Set Another = New Collection
    Assert.IsTrue CollectionEqualsAsSet(Nothing, Nothing)
    
    Assert.IsFalse CollectionEqualsAsSet(One, Nothing)
    Assert.IsFalse CollectionEqualsAsSet(Nothing, Another)
    
    Assert.IsTrue CollectionEqualsAsSet(One, One)
    Assert.IsTrue CollectionEqualsAsSet(Another, Another)
    
    Assert.IsTrue CollectionEqualsAsSet(One, Another)
    
    Call One.Add(1)
    Assert.IsFalse CollectionEqualsAsSet(One, Another)
    Call Another.Add(1)
    Assert.IsTrue CollectionEqualsAsSet(One, Another)
    Call One.Add(2)
    Assert.IsFalse CollectionEqualsAsSet(One, Another)
    Call Another.Add(2)
    Assert.IsTrue CollectionEqualsAsSet(One, Another)
    
    Set One = New Collection
    Call One.Add(1)
    Set Another = New Collection
    Call Another.Add(2)
    Assert.IsFalse CollectionEqualsAsSet(One, Another)
    
    Set One = New Collection
    Call One.Add(1)
    Call One.Add(2)
    Set Another = New Collection
    Call Another.Add(2)
    Call Another.Add(1)
    Assert.IsTrue CollectionEqualsAsSet(One, Another)
    
    Set One = New Collection
    Call One.Add(1)
    Call One.Add(1)
    Call One.Add(2)
    Set Another = New Collection
    Call Another.Add(1)
    Call Another.Add(2)
    Call Another.Add(2)
    Assert.IsTrue CollectionEqualsAsSet(One, Another)
    
    Set One = New Collection
    Call One.Add(1)
    Set Another = New Collection
    Call Another.Add(1)
    Call Another.Add(2)
    Assert.IsFalse CollectionEqualsAsSet(One, Another)
    
    Dim obj1 As Object: Set obj1 = New Collection
    Dim obj2 As Object: Set obj2 = New Collection
    
    Set One = New Collection
    Call One.Add(obj1)
    Set Another = New Collection
    Call Another.Add(obj1)
    Assert.IsTrue CollectionEqualsAsSet(One, Another)
    
    Set One = New Collection
    Call One.Add(obj1)
    Set Another = New Collection
    Call Another.Add(obj2)
    Assert.IsFalse CollectionEqualsAsSet(One, Another)
    
    Set One = New Collection
    Call One.Add(obj1)
    Set Another = New Collection
    Call Another.Add(1)
    Assert.IsFalse CollectionEqualsAsSet(One, Another)
    
    Set One = New Collection
    Call One.Add(1)
    Set Another = New Collection
    Call Another.Add(obj2)
    Assert.IsFalse CollectionEqualsAsSet(One, Another)
    
    Set One = New Collection
    Call One.Add(1)
    Call One.Add(obj1)
    Set Another = New Collection
    Call Another.Add(1)
    Assert.IsFalse CollectionEqualsAsSet(One, Another)
    
    Set One = New Collection
    Call One.Add(1)
    Set Another = New Collection
    Call Another.Add(1)
    Call Another.Add(obj2)
    Assert.IsFalse CollectionEqualsAsSet(One, Another)
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    GoTo FINALLY
FINALLY:
End Sub

'@TestMethod("CollectionDistinct")
Private Sub CollectionDistinct_CorrectCall_Successed()
    Dim Target As Collection
    Dim newOne As Collection
    
    On Error GoTo CATCH
    
    Set Target = New Collection
    Set newOne = CollectionDistinct(Target)
    Target.Add (1)
    Set newOne = CollectionDistinct(Target)
    Assert.AreEqual newOne.Count, Target.Count
    Target.Add (2)
    Set newOne = CollectionDistinct(Target)
    Assert.AreEqual newOne.Count, Target.Count
    
    Set Target = New Collection
    Target.Add (1)
    Target.Add (1)
    Set newOne = CollectionDistinct(Target)
    Assert.AreEqual newOne.Count, Target.Count - 1
    Target.Add (2)
    Target.Add (2)
    Set newOne = CollectionDistinct(Target)
    Assert.AreEqual newOne.Count, Target.Count - 2
    
    GoTo FINALLY
CATCH:
    Assert.Fail Err.Description
    GoTo FINALLY
FINALLY:
End Sub

'@TestMethod("CollectionDistinct")
Private Sub CollectionDistinct_IllegalCall_Error()
    On Error GoTo CATCH
    
    On Error GoTo ERROR_RAISED
    Call CollectionDistinct(Nothing)
    
    Assert.Fail
    GoTo FINALLY
    
ERROR_RAISED:
    Assert.AreEqual CLng(5), Err.Number
    Resume FINALLY
    
CATCH:
    Assert.Fail Err.Description
    GoTo FINALLY
FINALLY:
End Sub
