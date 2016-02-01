Attribute VB_Name = "modMain"
Option Explicit

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Const GWL_STYLE = -16

Public Const ES_NUMBER = &H2000

Public Const IDANI_OPEN = &H1
Public Const IDANI_CLOSE = &H2
Public Const IDANI_CAPTION = &H3

Public Const SW_SHOWNORMAL As Long = 1

Public Declare Sub InitCommonControls Lib "comctl32" ()

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function DrawAnimatedRects Lib "user32" (ByVal hWnd As Long, ByVal idani As Long, lprcFrom As RECT, lprcTo As RECT) As Long

Public Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function EnableScroll Lib "VCScroll.dll" () As Long
Public Declare Function DisableScroll Lib "VCScroll.dll" () As Long
Public Declare Sub ScrollLines Lib "VCScroll.dll" (ByVal Number As Long)
Public Declare Sub SetWheelButton Lib "VCScroll.dll" (ByVal Press As Long, ByVal Hold As Long)
Public Declare Sub SetCtrlKey Lib "VCScroll.dll" (ByVal Value As Long)

Public Function ShortcutPath(ByVal Profile As Variant) As String
    ShortcutPath = CreateObject("WScript.Shell").SpecialFolders(Profile) & "\" & App.Title & ".lnk"
End Function

Public Sub CreateShortcut(ByVal Profile As Variant)
    Dim objShell As Object
    Dim objShortcut As Object
    Dim exePath As String
    Dim lnkPath As String
    
    exePath = App.Path & "\" & App.EXEName & ".exe"
    lnkPath = ShortcutPath(Profile)
    
    Set objShell = CreateObject("WScript.Shell")
    Set objShortcut = objShell.CreateShortcut(lnkPath)
    
    objShortcut.TargetPath = exePath
    objShortcut.WindowStyle = 1
    objShortcut.IconLocation = exePath & ", 0"
    objShortcut.Description = App.Title
    objShortcut.WorkingDirectory = App.Path
    
    objShortcut.Save
    
    Set objShortcut = Nothing
    Set objShell = Nothing
End Sub

Public Sub DeleteShortcut(ByVal Profile As Variant)
    On Error Resume Next
    Kill ShortcutPath(Profile)
End Sub
