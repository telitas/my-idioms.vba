Attribute VB_Name = "PerformanceSettingModuleTest"
'@IgnoreModule IndexedUnboundDefaultMemberAccess
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

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
'@Ignore EmptyMethod
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
'@Ignore EmptyMethod
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("ApplyPerformanceSetting")
Private Sub ApplyPerfSetting_CorrectCall_Successed()
    Dim initialEnableEvents As Boolean
    Dim initialScreenUpdating As Boolean
    Dim initialCalculation As XlCalculation
    With Application
        .EnableEvents = True
        .ScreenUpdating = False
        .Calculation = xlCalculationSemiautomatic
    End With
    initialEnableEvents = Application.EnableEvents
    initialScreenUpdating = Application.ScreenUpdating
    initialCalculation = Application.Calculation
    
    Dim expectEnableEvents As Boolean
    Dim expectScreenUpdating As Boolean
    Dim expectCalculation As XlCalculation
    expectEnableEvents = False
    expectScreenUpdating = True
    expectCalculation = xlCalculationManual
    
    Dim initialState As Object
    Set initialState = ApplyPerformanceSetting( _
        Calculation:=expectCalculation, _
        EnableEvents:=expectEnableEvents, _
        ScreenUpdating:=expectScreenUpdating _
    )
    Assert.AreEqual Application.Calculation, expectCalculation
    Assert.AreEqual Application.EnableEvents, expectEnableEvents
    Assert.AreEqual Application.ScreenUpdating, expectScreenUpdating
    Assert.AreEqual initialState("Calculation"), initialCalculation
    Assert.AreEqual initialState("EnableEvents"), initialEnableEvents
    Assert.AreEqual initialState("ScreenUpdating"), initialScreenUpdating
    
    Dim secondState As Object
    Set secondState = ApplyPerformanceSetting
    
    Assert.AreEqual Application.Calculation, expectCalculation
    Assert.AreEqual Application.EnableEvents, expectEnableEvents
    Assert.AreEqual Application.ScreenUpdating, expectScreenUpdating
    Assert.AreEqual secondState("Calculation"), expectCalculation
    Assert.AreEqual secondState("EnableEvents"), expectEnableEvents
    Assert.AreEqual secondState("ScreenUpdating"), expectScreenUpdating
End Sub

'@TestMethod("ApplyPerformanceSettingWithDictionary")
Private Sub ApplyPerformanceSettingWithDictionary_CorrectCall_Successed()
    Dim initialEnableEvents As Boolean
    Dim initialScreenUpdating As Boolean
    Dim initialCalculation As XlCalculation
    With Application
        .EnableEvents = True
        .ScreenUpdating = False
        .Calculation = xlCalculationSemiautomatic
    End With
    initialEnableEvents = Application.EnableEvents
    initialScreenUpdating = Application.ScreenUpdating
    initialCalculation = Application.Calculation
    
    Dim expectSetting As Object
    Set expectSetting = CreateObject("Scripting.Dictionary")
    expectSetting("EnableEvents") = False
    expectSetting("ScreenUpdating") = True
    expectSetting("Calculation") = xlCalculationManual
    
    Dim initialState As Object
    Set initialState = ApplyPerformanceSettingWithDictionary(expectSetting)
    Assert.AreEqual Application.Calculation, expectSetting("Calculation")
    Assert.AreEqual Application.EnableEvents, expectSetting("EnableEvents")
    Assert.AreEqual Application.ScreenUpdating, expectSetting("ScreenUpdating")
    Assert.AreEqual initialState("Calculation"), initialCalculation
    Assert.AreEqual initialState("EnableEvents"), initialEnableEvents
    Assert.AreEqual initialState("ScreenUpdating"), initialScreenUpdating
    
    Dim secondState As Object
    Set secondState = ApplyPerformanceSettingWithDictionary(CreateObject("Scripting.Dictionary"))
    
    Assert.AreEqual Application.Calculation, expectSetting("Calculation")
    Assert.AreEqual Application.EnableEvents, expectSetting("EnableEvents")
    Assert.AreEqual Application.ScreenUpdating, expectSetting("ScreenUpdating")
    Assert.AreEqual secondState("Calculation"), expectSetting("Calculation")
    Assert.AreEqual secondState("EnableEvents"), expectSetting("EnableEvents")
    Assert.AreEqual secondState("ScreenUpdating"), expectSetting("ScreenUpdating")
End Sub
