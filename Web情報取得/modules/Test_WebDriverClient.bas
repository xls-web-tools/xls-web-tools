Attribute VB_Name = "Test_WebDriverClient"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! WebDriver 呼び出し境界のユニット テストです。
'! Lib_UnitTest.UnitTestMain() によって実行されます。
'!
' #############################################################################

Public Sub Test_WebDriverClient_TestDouble_WebDriverコマンドを差し替えられる(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' --- Arrange ---
    Dim client_double As WebDriverClientTestDouble
    Set client_double = New WebDriverClientTestDouble
    Call client_double.Store.SetReturn("Execute", "{""value"":{""sessionId"":""abc""}}", "POST", "/session", "{""capabilities"":{}}")

    Dim client As IWebDriverClient
    Set client = client_double

    ' --- Act ---
    Dim actual_response As String
    actual_response = client.Execute("POST", "/session", "{""capabilities"":{}}")

    ' --- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "{""value"":{""sessionId"":""abc""}}", actual_response
    Assert.EqualsNumeric 1, client_double.Store.GetCallCount("Execute", "POST", "/session", "{""capabilities"":{}}")
End Sub
