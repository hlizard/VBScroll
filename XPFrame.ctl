VERSION 5.00
Begin VB.UserControl XPFrame 
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1995
   ControlContainer=   -1  'True
   ForwardFocus    =   -1  'True
   HitBehavior     =   0  'None
   LockControls    =   -1  'True
   ScaleHeight     =   100
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   133
End
Attribute VB_Name = "XPFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type Size
    cx As Long
    cy As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const BDR_SUNKENOUTER = &H2
Private Const BDR_RAISEDINNER = &H4
Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)

Private Const BF_LEFT = &H1
Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_BOTTOM = &H8
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Declare Function IsAppThemed Lib "uxtheme.dll" () As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As String) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pClipRect As Any) As Long
Private Declare Function DrawThemeText Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal pszText As String, ByVal iCharCount As Long, ByVal dwTextFlags As Long, ByVal dwTextFlags2 As Long, pRect As RECT) As Long

Private Declare Function GetTextExtentPoint32 Lib "gdi32.dll" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long

Private Declare Function SetRect Lib "user32.dll" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawEdge Lib "user32.dll" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawState Lib "user32.dll" Alias "DrawStateA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As String, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long

Private Const BP_GROUPBOX = 4

Private Const DSS_NORMAL = &H0
Private Const DSS_DISABLED = &H20

Private Const DST_PREFIXTEXT = &H2

Private comctl32 As Object
Private myEnabled As Boolean
Private myCaption As String
Private captionLeft As Single, captionSpace As Single

Private Sub UserControl_Initialize()
    On Error Resume Next
    Set comctl32 = CreateObject("COMCTL.ImageListCtrl")
End Sub

Private Sub UserControl_Terminate()
    Set comctl32 = Nothing
End Sub

Private Sub UserControl_InitProperties()
    myEnabled = True
    myCaption = Ambient.DisplayName
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    myEnabled = PropBag.ReadProperty("Enabled", True)
    myCaption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Enabled", myEnabled, True
    PropBag.WriteProperty "Caption", myCaption, Ambient.DisplayName
End Sub

Private Sub UserControl_Paint()
    DrawFrame
End Sub

Public Property Get Enabled() As Boolean
    Enabled = myEnabled
End Property

Public Property Let Enabled(ByVal Value As Boolean)
    myEnabled = Value
    PropertyChanged "Enabled"
    UserControl.Enabled = myEnabled
    DrawFrame
End Property

Public Property Get Caption() As String
    Caption = myCaption
End Property

Public Property Let Caption(ByVal Value As String)
    myCaption = Value
    PropertyChanged "Caption"
    DrawFrame
End Property

Public Property Get Themed() As Boolean
    Themed = False
    On Error Resume Next
    Themed = CBool(IsAppThemed)
    Themed = Themed And Not (comctl32 Is Nothing)
End Property

Private Sub DrawFrame()
    On Error GoTo ErrHand
    
    Dim text As String
    Dim sizeText As Size
    Dim rectEdge As RECT, rectText As RECT
    Dim theme As Long
    
    UserControl.Cls
    
    If Themed Then
        captionLeft = 9
        captionSpace = 2
    Else
        captionLeft = 4
        captionSpace = 1
    End If
    
    text = IIf(myCaption = "", " ", Replace(myCaption, "&", ""))
    GetTextExtentPoint32 UserControl.hDC, text, Len(text), sizeText
    SetRect rectEdge, 0, sizeText.cy / 2, UserControl.ScaleWidth, UserControl.ScaleHeight
    SetRect rectText, captionLeft, 0, captionLeft + sizeText.cx, sizeText.cy
    
    If Themed Then
        theme = OpenThemeData(UserControl.hWnd, StrConv("Button", vbUnicode))
        DrawThemeBackground theme, UserControl.hDC, BP_GROUPBOX, 0, rectEdge, ByVal 0
        If myCaption <> "" Then HideLine sizeText
        DrawThemeText theme, UserControl.hDC, BP_GROUPBOX, 1, StrConv(myCaption, vbUnicode), Len(myCaption), 0, IIf(myEnabled, DSS_NORMAL, DSS_DISABLED), rectText
        CloseThemeData theme
    Else
        DrawEdge UserControl.hDC, rectEdge, EDGE_ETCHED, BF_RECT
        If myCaption <> "" Then HideLine sizeText
        DrawState UserControl.hDC, 0, 0, myCaption, Len(myCaption), captionLeft, 0, sizeText.cx, sizeText.cy, IIf(myEnabled, DSS_NORMAL, DSS_DISABLED) Or DST_PREFIXTEXT
    End If
    
    Exit Sub
ErrHand:
    Err.Clear
End Sub

Private Sub HideLine(ByRef XY As Size)
    Line (captionLeft - captionSpace, 0)-(captionLeft + XY.cx + captionSpace, XY.cy), UserControl.BackColor, BF
End Sub
