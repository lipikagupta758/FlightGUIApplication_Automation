Dim qtApp, qtTest

Set qtApp = CreateObject("QuickTest.Application")

' Launch UFT if not already open
If Not qtApp.Launched Then
    qtApp.Launch
End If

qtApp.Visible = True

' Open your test (adjust path as needed)
testDir = createobject("Scripting.Filesystemobject").GetAbsolutePathName(".")
WScript.Echo "Resolved test path: " & testDir & "\TestScripts\MainTest"

qtApp.Open testDir& "\TestScripts\MainTest", True

' Set test object
Set qtTest = qtApp.Test

' Run the test (set to false to hide UFT during execution)
qtTest.Run

' Close the test after run
qtTest.Close

qtApp.Quit

' Cleanup
Set qtTest = Nothing
Set qtApp = Nothing
