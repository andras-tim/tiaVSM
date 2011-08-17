' tiaVSM Example module
' Created by Andras Tim @ 2011

' DESCRIPTION:
' The modules work like a class with additional module framework

' Optional constructor
Private Sub Module_Initialize()
    ' Optional import more module
    Import "module2_example"

    WScript.StdOut.WriteLine "module_example: construct"
End Sub

' Module's private/public methods and functions
Public Sub test(text, num)
    WScript.StdOut.WriteLine "module_example: test - " & text & " " & num
    modules("module2_example").test "devil", "angel"
End Sub
