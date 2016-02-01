VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3615
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VBScroll.XPFrame xfrCtrlKey 
      Height          =   2055
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   3625
      Caption         =   "Ctrl key"
      Begin VB.OptionButton optCtrlKey 
         Caption         =   "switch opened project windows"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Top             =   1680
         Width           =   2535
      End
      Begin VB.OptionButton optCtrlKey 
         Caption         =   "scroll procedures (code view)"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   2655
      End
      Begin VB.OptionButton optCtrlKey 
         Caption         =   "perform horizontal scrolling"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   2175
      End
      Begin VB.OptionButton optCtrlKey 
         Caption         =   "ignore keystroke"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1575
      End
      Begin VB.Image imgCtrlKey 
         Height          =   480
         Left            =   2760
         Picture         =   "Main.frx":1CFA
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblCtrlKey 
         AutoSize        =   -1  'True
         Caption         =   "When Ctrl key is pressed:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1800
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1440
      TabIndex        =   21
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   22
      Top             =   4920
      Width           =   975
   End
   Begin VBScroll.XPFrame xfrStartup 
      Height          =   615
      Left            =   120
      TabIndex        =   16
      Top             =   4080
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1085
      Caption         =   "Startup"
      Begin VB.OptionButton optCurrentUser 
         Caption         =   "Current user"
         Height          =   255
         Left            =   1140
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optAllUsers 
         Caption         =   "All users"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optManual 
         Caption         =   "Manual"
         Height          =   255
         Left            =   2400
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
   End
   Begin VBScroll.XPFrame xfrWheel 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   2990
      Caption         =   "Wheel"
      Begin VB.ComboBox cmbAction 
         Height          =   315
         ItemData        =   "Main.frx":25C4
         Left            =   1920
         List            =   "Main.frx":25CB
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1260
         Width           =   1335
      End
      Begin VB.OptionButton optScrollPage 
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   255
      End
      Begin VB.OptionButton optScrollLines 
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   255
      End
      Begin VB.VScrollBar sclLines 
         Height          =   220
         Left            =   810
         Max             =   99
         Min             =   1
         TabIndex        =   4
         Top             =   615
         Value           =   1
         Width           =   255
      End
      Begin VB.TextBox txtLines 
         Height          =   285
         Left            =   480
         MaxLength       =   2
         TabIndex        =   3
         Top             =   585
         Width           =   615
      End
      Begin VB.Label lblButton 
         AutoSize        =   -1  'True
         Caption         =   "Wheel button action:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   1485
      End
      Begin VB.Image imgWheel 
         Height          =   480
         Left            =   2760
         Picture         =   "Main.frx":25D8
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblScrollPage 
         AutoSize        =   -1  'True
         Caption         =   "one page at time"
         Height          =   195
         Left            =   480
         TabIndex        =   7
         Top             =   990
         Width           =   1185
      End
      Begin VB.Label lblScrollLines 
         AutoSize        =   -1  'True
         Caption         =   "lines at time"
         Height          =   195
         Left            =   1200
         TabIndex        =   5
         Top             =   630
         Width           =   825
      End
      Begin VB.Label lblScroll 
         AutoSize        =   -1  'True
         Caption         =   "Roll the wheel one notch to scroll:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2400
      End
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      Caption         =   "Version"
      ForeColor       =   &H80000011&
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   5010
      Width           =   525
   End
   Begin VB.Menu mnuTray 
      Caption         =   "empty"
      Visible         =   0   'False
      Begin VB.Menu mnuConfigure 
         Caption         =   "Configure"
      End
      Begin VB.Menu mnuVisit 
         Caption         =   "Visit my site"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private iniPath As String
Private WithEvents astMain As AdvSysTray
Attribute astMain.VB_VarHelpID = -1
Private pntTray As POINTAPI
Private rctTray As RECT, rctForm As RECT

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHand
    
    Dim myCaption As String, buffer As String
    Dim textStyle As Long, result As Long, pressKey As Long
    Dim i As Integer
    
    myCaption = App.Title & " by " & App.CompanyName
    
    If App.PrevInstance Then
        MsgBox App.Title & " is already running!", vbExclamation
        End
    End If
    
    Me.Caption = myCaption
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    iniPath = App.Path & "\" & App.EXEName & ".ini"
    
    For i = 1 To 10
        buffer = String$(100, vbNullChar)
        result = GetPrivateProfileString("WheelButtonAction", "Caption" & CStr(i), "Action" & CStr(i), buffer, Len(buffer), iniPath)
        pressKey = GetPrivateProfileInt("WheelButtonAction", "PressKey" & CStr(i), 0, iniPath)
        
        If pressKey = 0 Then Exit For
        
        cmbAction.AddItem Left$(buffer, result)
        cmbAction.ItemData(i) = pressKey
    Next
    
    ReadSettings
    
    textStyle = GetWindowLong(txtLines.hWnd, GWL_STYLE)
    SetWindowLong txtLines.hWnd, GWL_STYLE, textStyle Or ES_NUMBER
    
    sclLines.Value = txtLines.text
    
    EnableScroll
    ScrollLines IIf(optScrollLines.Value, CLng(txtLines.text), 0)
    SetWheelButton cmbAction.ItemData(cmbAction.ListIndex), CLng(cmbAction.Tag)
    
    For i = optCtrlKey.LBound To optCtrlKey.UBound
        If optCtrlKey.Item(i).Value Then
            SetCtrlKey i
            Exit For
        End If
    Next
    
    Set astMain = New AdvSysTray
    
    astMain.Create Me
    astMain.Tooltip = myCaption
    If GetSetting(App.Title, "Settings", "FirstTime", True) Then
        astMain.ShowBalloon "Program started successfully. Click here to configure it.", Me.Caption, NIIF_INFO
        SaveSetting App.Title, "Settings", "FirstTime", False
    End If
    
    Exit Sub
ErrHand:
    MsgBox Err.Description, vbCritical, Err.Source
    End
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        AnimateHide
        Cancel = True
    Else
        DisableScroll
        
        astMain.Destroy
        
        Set astMain = Nothing
    End If
End Sub

Private Sub sclLines_Change()
    txtLines.text = sclLines.Value
End Sub

Private Sub optScrollLines_Click()
    txtLines.Enabled = optScrollLines.Value
    sclLines.Enabled = txtLines.Enabled
End Sub

Private Sub optScrollPage_Click()
    txtLines.Enabled = optScrollLines.Value
    sclLines.Enabled = txtLines.Enabled
End Sub

Private Sub lblScrollLines_Click()
    optScrollLines.Value = True
    optScrollLines_Click
End Sub

Private Sub lblScrollPage_Click()
    optScrollPage.Value = True
    optScrollPage_Click
End Sub

Private Sub cmdOk_Click()
    On Error GoTo ErrHand
    
    Dim i As Integer
    
    If CInt(txtLines.text) < 1 Then Exit Sub
    
    SaveSetting App.Title, "Settings", "ScrollLines", txtLines.text
    SaveSetting App.Title, "Settings", "ScrollPage", optScrollPage.Value
    SaveSetting App.Title, "Settings", "WheelButton", cmbAction.ListIndex
    
    ScrollLines IIf(optScrollLines.Value, CLng(txtLines.text), 0)
    SetWheelButton cmbAction.ItemData(cmbAction.ListIndex), CLng(cmbAction.Tag)
    
    For i = optCtrlKey.LBound To optCtrlKey.UBound
        If optCtrlKey.Item(i).Value Then
            SaveSetting App.Title, "Settings", "CtrlKey", i
            SetCtrlKey i
            Exit For
        End If
    Next
    
    If optAllUsers.Value Then
        DeleteShortcut "Startup"
        CreateShortcut "AllUsersStartup"
    ElseIf optCurrentUser.Value Then
        DeleteShortcut "AllUsersStartup"
        CreateShortcut "Startup"
    Else
        DeleteShortcut "Startup"
        DeleteShortcut "AllUsersStartup"
    End If
    
    AnimateHide
    
    Exit Sub
ErrHand:
    MsgBox Err.Description, vbCritical, Err.Source
    Err.Clear
End Sub

Private Sub cmdCancel_Click()
    ReadSettings
    AnimateHide
End Sub

Private Sub astMain_LButtonDown()
    GetCursorPos pntTray
End Sub

Private Sub astMain_RButtonDown()
    GetCursorPos pntTray
End Sub

Private Sub astMain_LButtonDblClk()
    AnimateShow
End Sub

Private Sub astMain_RButtonUp()
    SetForegroundWindow Me.hWnd
    If Not Me.Visible Then
        Me.PopupMenu mnuTray, , , , mnuConfigure
    End If
End Sub

Private Sub astMain_BalloonUserClick()
    GetCursorPos pntTray
    AnimateShow
End Sub

Private Sub mnuConfigure_Click()
    astMain_LButtonDblClk
End Sub

Private Sub mnuVisit_Click()
    ShellExecute Me.hWnd, vbNullString, "http://www.gasanov.net", vbNullString, vbNullString, SW_SHOWNORMAL
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub ReadSettings()
    txtLines.text = GetSetting(App.Title, "Settings", "ScrollLines", 3)
    optScrollPage.Value = GetSetting(App.Title, "Settings", "ScrollPage", False)
    optScrollLines.Value = Not optScrollPage.Value
    cmbAction.ListIndex = GetSetting(App.Title, "Settings", "WheelButton", 0)
    cmbAction.Tag = GetPrivateProfileInt("WheelButtonAction", "HoldKey" & cmbAction.ListIndex, 0, iniPath)
    
    optCtrlKey.Item(GetSetting(App.Title, "Settings", "CtrlKey", 1)).Value = True
    
    optAllUsers.Value = (Dir$(ShortcutPath("AllUsersStartup")) <> "")
    optCurrentUser.Value = (Dir$(ShortcutPath("Startup")) <> "")
    optManual.Value = Not (optAllUsers.Value Or optCurrentUser.Value)
End Sub

Private Sub AnimateShow()
    If Me.Visible Then
        SetForegroundWindow Me.hWnd
    Else
        rctTray.Left = pntTray.x
        rctTray.Top = pntTray.y
        rctTray.Right = pntTray.x
        rctTray.Bottom = pntTray.y
        
        Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
        GetWindowRect Me.hWnd, rctForm
        DrawAnimatedRects Me.hWnd, IDANI_OPEN Or IDANI_CAPTION, rctTray, rctForm
        Me.Show
    End If
End Sub

Private Sub AnimateHide()
    GetWindowRect Me.hWnd, rctForm
    DrawAnimatedRects Me.hWnd, IDANI_CLOSE Or IDANI_CAPTION, rctForm, rctTray
    Me.Hide
End Sub
