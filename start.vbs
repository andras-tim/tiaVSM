'''''' TIAVSM BEGIN ''''''
' tiaVSM (tia VisuabBasicScript ModuleFramework)
' Version: 0.1.110818
' Created by: Andras Tim @ 2011
'
Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim startPath: startPath = fso.GetParentFolderName(wscript.ScriptFullName)

Dim modules: Set modules = CreateObject("Scripting.Dictionary")
Dim moduleClassTemplate: moduleClassTemplate = _
"Class %CLASS%" & vbCrLf & _
"%CODE%" & vbCrLf & _
"Dim modulePath" & vbCrLf & _
"Private Sub Class_Initialize()" & vbCrLf & _
"    modulePath = ""%PATH%""" & vbCrLf & _
"    Dim existInit: On Error Resume Next: IsObject Module_Initialize: " & _
    "existInit = (Err.Number = 13): Err.Clear: On Error Goto 0" & vbCrLf & _
"    If existInit Then Module_Initialize" & vbCrLf & _
"End Sub" & vbCrLf & _
"End Class"

Public Sub Import(moduleName)
    Dim moduleFullPath: moduleFullPath =  startPath & "\" & moduleName & ".vbs"
    If Not modules.Exists(moduleName) Then
        'Load class contain from file
        Dim strCode, fo: Set fo = fso.OpenTextFile(moduleFullPath): strCode = fo.ReadAll: fo.Close
        'Get a new class template and substitue it
        Dim classCode: classCode = moduleClassTemplate
        classCode = Replace(classCode, "%CLASS%", "mod_" & moduleName)
        classCode = Replace(classCode, "%PATH%", fso.GetParentFolderName(moduleFullPath))
        classCode = Replace(classCode, "%CODE%", strCode)
        'Eval class
        On Error Resume Next
        ExecuteGlobal classCode
        If Not Err.Number = 0 Then
            WScript.StdOut.WriteLine "====== CODE ======" & vbCrLf & classCode & vbCrLf & "====== CODE ======" & vbCrLf
            WScript.StdOut.WriteLine "Microsoft VBScript runtime error: " & Err.Description & vbCrLf
            WScript.StdOut.WriteLine "MODULE: " & moduleName & vbCrLf & "METHOD: Code evaluate"
            WScript.Quit(1)
        End If
        On Error Goto 0
        'Register a new instance of class into modules
        ExecuteGlobal "modules.Add """ & moduleName & """, New mod_" & moduleName
        If Not Err.Number = 0 Then
            WScript.StdOut.WriteLine "MODULE: " & moduleName & vbCrLf & "METHOD: Module_Initialize"
            WScript.Quit(1)
        End If
    End If
End Sub
Public Sub Include(fileName)
    'Load VBScript from file
    Dim strCode, fo: Set fo = fso.OpenTextFile(startPath & "\" & fileName): strCode = fo.ReadAll: fo.Close
    Execute strCode
End Sub
'''''' TIAVSM END ''''''

' MAIN
WScript.StdOut.WriteLine vbCrLf & "### Before import"
Import "module_example"

WScript.StdOut.WriteLine  vbCrLf & "### Before run method from a module"
modules("module_example").test "foo", 2

WScript.StdOut.WriteLine vbCrLf & "### Before run method from an another module (linked module)"
modules("module2_example").test "cain", "abel"
