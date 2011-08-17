' tiaVSM Example module
' Created by Andras Tim @ 2011

' DESCRIPTION:
' The modules work like a class with additional module framework

' Optional constructor
Sub Module_Initialize()
    WScript.StdOut.WriteLine "module2_example: construct"
End Sub

Public Sub test(text1, test2)
    WScript.StdOut.WriteLine "module2_example: test - " & text1 & " " & text2
End Sub