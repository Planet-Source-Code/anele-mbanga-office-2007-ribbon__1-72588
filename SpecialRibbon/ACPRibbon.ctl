VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl ACPRibbon 
   BackColor       =   &H00404040&
   ClientHeight    =   3705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7095
   EditAtDesignTime=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   ScaleHeight     =   3705
   ScaleWidth      =   7095
   ToolboxBitmap   =   "ACPRibbon.ctx":0000
   Begin VB.ComboBox cboMenus1 
      Height          =   315
      Left            =   3480
      TabIndex        =   18
      Text            =   "Combo1"
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cboMenus 
      Height          =   315
      Left            =   2760
      TabIndex        =   17
      Text            =   "Combo1"
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComctlLib.ProgressBar progBar 
      Height          =   315
      Index           =   0
      Left            =   1440
      TabIndex        =   15
      Top             =   2280
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.ComboBox cboMaster 
      Height          =   315
      Left            =   2760
      TabIndex        =   14
      Text            =   "Combo1"
      Top             =   2640
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComCtl2.DTPicker datePick 
      Height          =   315
      Index           =   0
      Left            =   5640
      TabIndex        =   13
      Top             =   2400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/mm/yyyy"
      Format          =   120061955
      CurrentDate     =   40111
   End
   Begin VB.ComboBox cboBox 
      Height          =   315
      Index           =   0
      Left            =   6240
      TabIndex        =   12
      Text            =   "Combo1"
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox txtBox 
      Height          =   315
      Index           =   0
      Left            =   2400
      TabIndex        =   11
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Button_Text 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label ButMouse 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   0
      Left            =   4320
      TabIndex        =   9
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Glip_on 
      Height          =   60
      Index           =   0
      Left            =   4560
      Top             =   2280
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Image Glip_off 
      Height          =   60
      Index           =   0
      Left            =   4440
      Top             =   2280
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Image Button_left_over 
      Height          =   990
      Index           =   0
      Left            =   4800
      Top             =   2520
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image Button_center_over 
      Height          =   990
      Index           =   0
      Left            =   4920
      Stretch         =   -1  'True
      Top             =   2520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Button_right_over 
      Height          =   990
      Index           =   0
      Left            =   5760
      Top             =   2520
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image Cat_Dlg_over 
      Height          =   210
      Index           =   0
      Left            =   4800
      Top             =   1920
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image Cat_Dlg_on 
      Height          =   210
      Index           =   0
      Left            =   4560
      Top             =   1920
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image Cat_Dlg 
      Height          =   210
      Index           =   0
      Left            =   4320
      Top             =   1920
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image Button_Icon 
      Appearance      =   0  'Flat
      Height          =   495
      Index           =   0
      Left            =   3600
      Top             =   2640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Button_Caption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   3735
      TabIndex        =   10
      Top             =   3240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image RibbonTopCustom_over 
      Height          =   390
      Left            =   4680
      Top             =   480
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image RibbonTopCustom 
      Height          =   390
      Left            =   4440
      Top             =   480
      Width           =   225
   End
   Begin VB.Image Button_right 
      Height          =   990
      Index           =   0
      Left            =   4200
      Top             =   2520
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image Button_center 
      Height          =   990
      Index           =   0
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   2520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Button_left 
      Height          =   990
      Index           =   0
      Left            =   3240
      Top             =   2520
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label TBMouse 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   0
      Left            =   4080
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image RibbonTopImage 
      Height          =   390
      Index           =   0
      Left            =   3360
      Top             =   480
      Width           =   270
   End
   Begin VB.Image RibbonTop_over 
      Height          =   390
      Index           =   0
      Left            =   3720
      Top             =   480
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label TabMouse 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   1800
      TabIndex        =   5
      Top             =   2760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Tab_caption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aba 01"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   2820
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image Tab_right 
      Height          =   360
      Index           =   0
      Left            =   1560
      Top             =   2760
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image Tab_center 
      Height          =   360
      Index           =   0
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   2760
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image Tab_left 
      Height          =   360
      Index           =   0
      Left            =   960
      Top             =   2760
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image Tab_left_over 
      Height          =   360
      Index           =   0
      Left            =   960
      Top             =   3240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image Tab_center_over 
      Height          =   360
      Index           =   0
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   3240
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image Tab_right_over 
      Height          =   360
      Index           =   0
      Left            =   1560
      Top             =   3240
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label CatMouse 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1350
      Index           =   0
      Left            =   5280
      TabIndex        =   7
      Top             =   750
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Cat_Caption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   5760
      TabIndex        =   6
      Tag             =   "sadf"
      Top             =   1800
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Cat_Right_on 
      Height          =   1335
      Index           =   0
      Left            =   6840
      Top             =   750
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Image Cat_Center_on 
      Height          =   1335
      Index           =   0
      Left            =   6600
      Stretch         =   -1  'True
      Top             =   750
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image Cat_Left_on 
      Height          =   1335
      Index           =   0
      Left            =   6480
      Top             =   750
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image Cat_Right_off 
      Height          =   1335
      Index           =   0
      Left            =   6120
      Top             =   750
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Image Cat_Left_off 
      Height          =   1335
      Index           =   0
      Left            =   5760
      Top             =   750
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image Cat_Center_off 
      Height          =   1335
      Index           =   0
      Left            =   5880
      Stretch         =   -1  'True
      Top             =   750
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label ButtonRibbon 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2760
      TabIndex        =   3
      Top             =   960
      Width           =   690
   End
   Begin VB.Image Endon 
      Height          =   345
      Left            =   6240
      Top             =   0
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image Maxon 
      Height          =   345
      Left            =   5520
      Top             =   0
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image Minon 
      Height          =   345
      Left            =   4800
      Top             =   0
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image Endoff 
      Height          =   345
      Left            =   3960
      Top             =   0
      Width           =   600
   End
   Begin VB.Image Maxoff 
      Height          =   345
      Left            =   3240
      Top             =   0
      Width           =   600
   End
   Begin VB.Image Minoff 
      Height          =   345
      Left            =   2520
      Top             =   0
      Width           =   600
   End
   Begin VB.Label Barra 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Titulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   465
   End
   Begin VB.Label Titulo2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FFD18A&
      Height          =   195
      Left            =   840
      TabIndex        =   1
      Top             =   75
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image RibbonTopRight 
      Height          =   390
      Left            =   3120
      Top             =   480
      Width           =   195
   End
   Begin VB.Image RibbonTop 
      Height          =   390
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   480
      Width           =   270
   End
   Begin VB.Image Logo 
      Height          =   360
      Left            =   2760
      Top             =   1680
      Width           =   360
   End
   Begin VB.Image ButtonRibbonon 
      Height          =   675
      Left            =   1800
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image ButtonRibbonover 
      Height          =   675
      Left            =   1800
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image ButtonRibbonoff 
      Height          =   675
      Left            =   1800
      Top             =   480
      Width           =   735
   End
   Begin VB.Image BarraLeft 
      Height          =   2130
      Left            =   0
      Top             =   0
      Width           =   105
   End
   Begin VB.Image BarraRight 
      Height          =   2130
      Left            =   960
      Top             =   0
      Width           =   105
   End
   Begin VB.Image Barra2 
      Height          =   2130
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   405
   End
End
Attribute VB_Name = "ACPRibbon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
'#######################################
'#                                     #
'#           ACP Ribbon 2007           #
'#                  origianlby                 #
'#      adrianopaladini@gmail.com      #
'#               'enhanced by
'#             Anele Mbanga (anele@safiri.co.za)
'#                                     #
'#                                     #con
'#  Visual from Office 2007 Beta 2 TR  #
'#                                     #
'#   Please Don´t Remove Author & Enhancer Info!  #
'#                                     #
'#######################################
'------------------------------------------------
' TO DO:
'
' A) Update when Resize, resolve flicks
' B) Optimize Code
' C) Insert Mini Buttons, Combos and Checkbox on Each Categories
' D) Option to Show Menu Under the Ribbon and Hide Ribbon
' E) Make Menu
' F) Option to user customize the menu
' G) Group Tabs
' H) Add Comment to All code
' I) FINISHED this project!
'
'------------------------------------------------
'------------------------------------------------
' Bugs:
'
' Please report to:
'
'         adrianopaladini@gmail.com
'
'------------------------------------------------
' enhancements done by Anele Mbanga (anelem@rocketmail.com) are the following
' the enhancement made include the following
' addition of textbox, combobox, datepicker, progress bar, label on buttons
' animation of main icon, see timer functions on form 1
' ability to edit the top button, edit the tab caption and button caption including icons thereof
' buttons now have menus that can be assigned to them
' buttons and tabs are no longer limited to 90 buttons, a redimensionable array has been used across the board
' menus, buttons can be added depending on the permissions per button, the permission string must contain the id of a button separated by ;
' if you like please vote for me
Private TotalTopButton As Integer
Private TotalButton As Integer
Private TotalTabs As Integer
Private TotalCats As Integer
Private Type TabButton
    TabID As String
    TabCaption As String
    TabVisible As Boolean
End Type
Private TabButtons() As TabButton
Private Type CategoryButton
    CatsID As String
    CatsC As String
    CatsT As String
    CatsD As String
    CatsTool As String
End Type
Private CategoryButtons() As CategoryButton
Private Type TopButton
    TopBID As String
    TopBC As String
End Type
Private TopButtons() As TopButton
'Private mvarHandle As Long
Private TabSelected As String
Private DefFont As StdFont
Private Type RibbonButton
    TopBuID As String
    TopBuS As String
    TopBuC As String
    TopBuI As Picture
    TopBuT As String
    TopBuG As Boolean
    TopBuX As String
    TopTxt As String
    TopWdt As Long
    TopType As String
    TopFormat As String
    TopMin As Long
    TopMax As Long
    menuName As String
End Type
Private sPermissions As String
Private RibbonButtons() As RibbonButton
'Private Type RECT
'    Left As Long
''    Top As Long
''    Right As Long
'    Bottom As Long
'End Type
Private MS As Boolean
Private Mx As Integer
Private My As Integer
Private iImgLType As Integer
Private sCaption As String
Private Const m_def_Caption = ""
Private Const m_def_ShowCustomMenu = False
Private m_ShowCustomMenu As Boolean
Private mvarUsePermissions As Boolean
Public Event MainMenuClick()
Public Event MenuClick(ByVal Id As String, ByVal Caption As String)
Public Event TabClick(ByVal Id As String, ByVal Caption As String)
Public Event CatClick(ByVal Id As String, ByVal Caption As String)
Public Event ButtonClick(ByVal Id As String, ByVal Caption As String)
Public Event ComboClick(ByVal ComboName As String, ByVal Text As String)
Public Event DatePickClick(ByVal DatePickName As String, ByVal DatePicked As String)
Public Event CustomClick()
Public Event CloseForm()
Public Event MaxForm()
Public Event MinForm()
Private zImg As Variant
Private TAB_NORMAL As Long
Private TAB_SELECTED As Long
Public Enum ThemeEnum
    Black = 0
    Blue = 1
    Silver = 2
End Enum
Public Enum ImageSizeEnum
    SizeNormal = 0
    Size160 = 1
    Size240 = 2
    Size320 = 3
End Enum
Private m_Theme As ThemeEnum
Private m_ImageSize As Integer
Private mParent As Variant
'Private Const WM_SETREDRAW = &HB
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
'Private Const RDW_INVALIDATE = &H1
'Private Const RDW_INTERNALPAINT = &H2
'Private Const RDW_UPDATENOW = &H100
'Private Const RDW_ALLCHILDREN = &H80
'Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Sub FreezeWindow(ObjSource As Variant, Optional boolAction As Boolean = True)
    On Error Resume Next
    If boolAction = True Then
        LockWindowUpdate ObjSource.hwnd
    Else
        LockWindowUpdate 0&
    End If
    Err.Clear
End Sub
Private Sub Barra_DblClick()
    On Error Resume Next
    Maxon_Click
    Err.Clear
End Sub
Private Sub Barra_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Mx = x
    My = y
    MS = True
    Err.Clear
End Sub
Private Sub Barra_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    'Freeze
    Dim i As Integer
    Dim iTot As Integer
    Dim KL As Integer
    Dim KLTot As Integer
    If MS = True Then
        mParent.Move mParent.Left - (Mx - x), mParent.Top - (My - y)
    End If
    iTot = TabMouse.UBound
    For i = 0 To iTot
        Tab_center_over(i).Visible = False
        Tab_left_over(i).Visible = False
        Tab_right_over(i).Visible = False
    Next
    iTot = CatMouse.UBound
    For i = 0 To iTot
        Cat_Center_on(i).Visible = False
        Cat_Left_on(i).Visible = False
        Cat_Right_on(i).Visible = False
        If Cat_Dlg(i).Visible = True Then
            Cat_Dlg_on(i).Visible = False
            Cat_Dlg_over(i).Visible = False
        End If
    Next
    KLTot = ButMouse.UBound
    For KL = 0 To KLTot
        Button_left(KL).Visible = False
        Button_right(KL).Visible = False
        Button_center(KL).Visible = False
    Next
    iTot = TBMouse.UBound
    For i = 0 To iTot
        RibbonTop_over(i).Visible = False
    Next
    RibbonTopCustom_over.Visible = False
    Endon.Visible = False
    Maxon.Visible = False
    Minon.Visible = False
    ButtonRibbonover.Visible = False
    ButtonRibbonon.Visible = False
    'Freeze False
    Err.Clear
End Sub
Private Sub Barra_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    MS = False
    Err.Clear
End Sub
Private Sub Barra2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    'Freeze
    Dim i As Integer
    Dim KL As Integer
    Dim iTot As Integer
    Dim KLTot As Integer
    iTot = TabMouse.UBound
    For i = 0 To iTot
        Tab_center_over(i).Visible = False
        Tab_left_over(i).Visible = False
        Tab_right_over(i).Visible = False
    Next
    iTot = CatMouse.UBound
    For i = 0 To iTot
        Cat_Center_on(i).Visible = False
        Cat_Left_on(i).Visible = False
        Cat_Right_on(i).Visible = False
        If Cat_Dlg(i).Visible = True Then
            Cat_Dlg_on(i).Visible = False
            Cat_Dlg_over(i).Visible = False
        End If
    Next
    KLTot = ButMouse.UBound
    For KL = 0 To KLTot
        Button_left(KL).Visible = False
        Button_right(KL).Visible = False
        Button_center(KL).Visible = False
    Next
    iTot = TBMouse.UBound
    For i = 0 To iTot
        RibbonTop_over(i).Visible = False
    Next
    RibbonTopCustom_over.Visible = False
    Endon.Visible = False
    Maxon.Visible = False
    Minon.Visible = False
    ButtonRibbonover.Visible = False
    ButtonRibbonon.Visible = False
    'Freeze False
    Err.Clear
End Sub
Private Sub BarraLeft_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    'Freeze
    Dim i As Integer
    Dim KL As Integer
    Dim iTot As Integer
    Dim KLTot As Integer
    iTot = TabMouse.UBound
    For i = 0 To iTot
        Tab_center_over(i).Visible = False
        Tab_left_over(i).Visible = False
        Tab_right_over(i).Visible = False
    Next
    iTot = CatMouse.UBound
    For i = 0 To iTot
        Cat_Center_on(i).Visible = False
        Cat_Left_on(i).Visible = False
        Cat_Right_on(i).Visible = False
        If Cat_Dlg(i).Visible = True Then
            Cat_Dlg_on(i).Visible = False
            Cat_Dlg_over(i).Visible = False
        End If
    Next
    KLTot = ButMouse.UBound
    For KL = 0 To KLTot
        Button_left(KL).Visible = False
        Button_right(KL).Visible = False
        Button_center(KL).Visible = False
    Next
    iTot = TBMouse.UBound
    For i = 0 To iTot
        RibbonTop_over(i).Visible = False
    Next
    RibbonTopCustom_over.Visible = False
    Endon.Visible = False
    Maxon.Visible = False
    Minon.Visible = False
    ButtonRibbonover.Visible = False
    ButtonRibbonon.Visible = False
    'Freeze False
    Err.Clear
End Sub
Private Sub BarraRight_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    'Freeze
    Dim i As Integer
    Dim KL As Integer
    Dim iTot As Integer
    Dim KLTot As Integer
    iTot = TabMouse.UBound
    For i = 0 To iTot
        Tab_center_over(i).Visible = False
        Tab_left_over(i).Visible = False
        Tab_right_over(i).Visible = False
    Next
    iTot = CatMouse.UBound
    For i = 0 To iTot
        Cat_Center_on(i).Visible = False
        Cat_Left_on(i).Visible = False
        Cat_Right_on(i).Visible = False
        If Cat_Dlg(i).Visible = True Then
            Cat_Dlg_on(i).Visible = False
            Cat_Dlg_over(i).Visible = False
        End If
    Next
    KLTot = ButMouse.UBound
    For KL = 0 To KLTot
        Button_left(KL).Visible = False
        Button_right(KL).Visible = False
        Button_center(KL).Visible = False
    Next
    iTot = TBMouse.UBound
    For i = 0 To iTot
        RibbonTop_over(i).Visible = False
    Next
    RibbonTopCustom_over.Visible = False
    Endon.Visible = False
    Maxon.Visible = False
    Minon.Visible = False
    ButtonRibbonover.Visible = False
    ButtonRibbonon.Visible = False
    'Freeze False
    Err.Clear
End Sub
Private Sub ButMouse_Click(Index As Integer)
    On Error Resume Next
    Dim butPos As Integer
    Dim strMenu As String
    Dim zID As String
    zID = ButMouse(Index).Tag
    butPos = ButtonSearch(ButMouse(Index).Tag)
    strMenu = RibbonButtons(butPos).menuName
    Select Case Len(strMenu)
    Case 0
        ' raise the normal button click
        RaiseEvent ButtonClick(ButMouse(Index).Tag, Button_Caption(Index).Caption)
    Case Else
        ' a menu has been selected
        Dim theMenu As clsMenu
        Dim subMenu As clsMenu
        Dim subTot As Integer
        Dim subCnt As Integer
        Dim menuCnt As Long
        Dim menuTot As Long
        Dim strLine As String
        Dim menuParent As String
        Dim menuID As String
        Dim menuCaption As String
        Dim menuSelect As String
        Dim lngChildren As Long
        Dim spChildren() As String
        Dim curChild As String
        Dim cntChildren As Integer
        Dim totChildren As Integer
        Dim spTot As Long
        Dim spCnt As Long
        Dim subMenuID As String
        Dim psubMenuID As String
        Dim hasSubMenus As String
        Dim strSimilar As String
        ' this is used to group the menus and extract those relevant to this menu
        cboMenus1.Clear
        menuTot = cboMenus.ListCount - 1
        For menuCnt = 0 To menuTot
            strLine = cboMenus.List(menuCnt)
            menuParent = MvField(strLine, 1, "|")
            menuID = MvField(strLine, 2, "|")
            menuID = RemDelim(menuID, "\")
            menuCaption = MvField(strLine, 3, "|")
            hasSubMenus = MvField(strLine, 4, "|")
            If LCase$(menuParent) = LCase$(zID) Then
                strSimilar = MenuSimilar(menuID)
                spTot = StrParse(spChildren, strSimilar, ";")
                For spCnt = 1 To spTot
                    cboMenus1.AddItem spChildren(spCnt)
                Next
            End If
        Next
        Set theMenu = New clsMenu
        theMenu.Reset
        menuTot = cboMenus1.ListCount - 1
        For menuCnt = 0 To menuTot
            strLine = cboMenus1.List(menuCnt)
            menuParent = MvField(strLine, 1, "|")
            menuID = MvField(strLine, 2, "|")
            menuID = RemDelim(menuID, "\")
            menuCaption = MvField(strLine, 3, "|")
            hasSubMenus = MvField(strLine, 4, "|")
            If LCase$(menuParent) <> LCase$(zID) Then GoTo NextMenu
            ' does this have submenus
            If InStr(1, menuID, "\") = 0 Then
                If hasSubMenus = "1" Then
                    Set subMenu = New clsMenu
                    subMenu.Caption = menuCaption
                    theMenu.AddMenu menuID, subMenu
                Else
                    theMenu.AddMenu menuID, menuCaption
                End If
            Else
                ' the menu has children
                spTot = StrParse(spChildren, menuID, "\")
                For spCnt = 2 To spTot
                    subMenuID = MvFromMv(menuID, spCnt, 1, "\")
                    subMenu.AddMenu subMenuID, menuCaption
                Next
            End If
NextMenu:
        Next
        If theMenu.MenuCount >= 1 Then
            menuSelect = theMenu.TrackMenu
            RaiseEvent ButtonClick(menuSelect, "")
        End If
    End Select
    Err.Clear
End Sub
Private Sub ButMouse_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Button_left_over(Index).Visible = True
    Button_center_over(Index).Visible = True
    Button_right_over(Index).Visible = True
    Err.Clear
End Sub
Private Sub ButMouse_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Button_left_over(Index).Visible = False
    Button_center_over(Index).Visible = False
    Button_right_over(Index).Visible = False
    Err.Clear
End Sub
Private Sub ButtonRibbon_Click()
    On Error Resume Next
    RaiseEvent MainMenuClick
    Err.Clear
End Sub
Private Sub ButtonRibbon_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    ButtonRibbonover.Visible = False
    ButtonRibbonon.Visible = True
    Err.Clear
End Sub
Private Sub ButtonRibbon_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    ButtonRibbonover.Visible = True
    ButtonRibbonon.Visible = False
    Err.Clear
End Sub
Private Sub Cat_Dlg_over_Click(Index As Integer)
    On Error Resume Next
    RaiseEvent CatClick(Cat_Caption(Index).Tag, Cat_Caption(Index).Caption)
    Err.Clear
End Sub
Private Sub cboBox_Click(Index As Integer)
    On Error Resume Next
    RaiseEvent ComboClick(cboBox(Index).Tag, cboBox(Index).Text)
    Err.Clear
End Sub
Private Sub datePick_CloseUp(Index As Integer)
    On Error Resume Next
    RaiseEvent DatePickClick(datePick(Index).Tag, Format$(datePick(Index).Value, datePick(Index).CustomFormat))
    Err.Clear
End Sub
Private Sub Endon_Click()
    On Error Resume Next
    Endon.Visible = False
    RaiseEvent CloseForm
    Unload mParent
    Err.Clear
End Sub
Private Sub Maxon_Click()
    On Error Resume Next
    Maxon.Visible = False
    RaiseEvent MaxForm
    Err.Clear
End Sub
Private Sub Minon_Click()
    On Error Resume Next
    Minon.Visible = False
    RaiseEvent MinForm
    Err.Clear
End Sub
Private Sub TBMouse_Click(Index As Integer)
    On Error Resume Next
    RaiseEvent MenuClick(TopButtons(Index).TopBID, TopButtons(Index).TopBC)
    Err.Clear
End Sub
'Public Sub SetRibbon()
'
'    UserControl_Initialize
'
'End Sub
Private Sub UserControl_Initialize()
    On Error Resume Next
    Theme = Blue
    ImageList = Nothing
    TotalTopButton = 0
    TotalButton = 0
    TotalTabs = 0
    TotalCats = 0
    Caption = "Ribbon Control"
    TabSelected = ""
    ImageSize = SizeNormal
    Barra.BackStyle = 0
    ButtonRibbon.BackStyle = 0
    TabMouse(0).BackStyle = 0
    CatMouse(0).BackStyle = 0
    TBMouse(0).BackStyle = 0
    ButMouse(0).BackStyle = 0
    Err.Clear
End Sub
Public Property Get ImageSize() As ImageSizeEnum
    On Error Resume Next
    ImageSize = m_ImageSize
    Err.Clear
End Property
Public Property Let ImageSize(ByVal New_Size As ImageSizeEnum)
    On Error Resume Next
    m_ImageSize = New_Size
    PropertyChanged "ImageSize"
    Err.Clear
End Property
'Public Property Get Handle() As Long
'
'    Handle = mvarHandle
'
'End Property
'Public Property Let Handle(ByVal New_Handle As Long)
'
'    mvarHandle = New_Handle
'    PropertyChanged "Handle"
'
'End Property
Public Property Get UsePermissions() As Boolean
    On Error Resume Next
    UsePermissions = mvarUsePermissions
    Err.Clear
End Property
Public Property Let UsePermissions(ByVal New_Handle As Boolean)
    On Error Resume Next
    mvarUsePermissions = New_Handle
    PropertyChanged "UsePermissions"
    Err.Clear
End Property
Public Property Get Caption() As String
    On Error Resume Next
    Caption = sCaption
    Err.Clear
End Property
Public Property Let Caption(ByVal New_Caption As String)
    On Error Resume Next
    Dim InicioArea As Long
    Dim area As Long
    FreezeWindow Me
    sCaption = New_Caption
    PropertyChanged "Caption"
    If m_ShowCustomMenu = True Then
        InicioArea = (RibbonTopCustom.Left + RibbonTopCustom.Width)
    Else
        InicioArea = (RibbonTopRight.Left + RibbonTopRight.Width)
    End If
    area = UserControl.Width - (InicioArea + (Endoff.Width * 3))
    'pos = InStr(sCaption, " - ")
    'If pos > 0 Then
    '    Titulo.Caption = Mid$(sCaption, 1, pos + 2)
    '    Titulo2.Caption = Mid$(sCaption, pos + 3)
    '    Titulo.Left = ((area - (Titulo.Width + Titulo2.Width)) / 2) + InicioArea
    '    Titulo2.Left = Titulo.Left + Titulo.Width
    '    Titulo2.Visible = True
    'Else
    Titulo.Caption = sCaption
    Titulo.Left = ((area - Titulo.Width) / 2) + InicioArea
    Titulo2.Visible = False
    'End If
    FreezeWindow Me, False
    Err.Clear
End Property
Public Property Get Permissions() As String
    On Error Resume Next
    Permissions = sPermissions
    Err.Clear
End Property
Public Property Let Permissions(ByVal vData As String)
    On Error Resume Next
    sPermissions = vData
    PropertyChanged "Permissions"
    Err.Clear
End Property
Public Sub Refresh()
    On Error Resume Next
    FreezeWindow Me
    TabsUpdate
    CatsUpdate
    FreezeWindow Me, False
    Err.Clear
End Sub
Private Sub UserControl_InitProperties()
    On Error Resume Next
    sCaption = m_def_Caption
    m_ShowCustomMenu = m_def_ShowCustomMenu
    Theme = 1
    Err.Clear
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    'Freeze
    Set DefFont = New StdFont
    DefFont.Name = "Tahoma"
    DefFont.Size = 8
    ImageSize = PropBag.ReadProperty("ImageSize", m_ImageSize)
    Set Font = PropBag.ReadProperty("Font", DefFont)
    Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    ShowCustomMenu = PropBag.ReadProperty("ShowCustomMenu", m_def_ShowCustomMenu)
    Theme = PropBag.ReadProperty("Theme", 1)
    UsePermissions = PropBag.ReadProperty("UsePermissions", True)
    'Freeze False
    Err.Clear
End Sub
Private Sub UserControl_Show()
    On Error Resume Next
    Resize
    Err.Clear
End Sub
Private Sub UserControl_Terminate()
    On Error Resume Next
    Me.Clear
    Err.Clear
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error Resume Next
    Set DefFont = New StdFont
    DefFont.Name = "Tahoma"
    DefFont.Size = 8
    Call PropBag.WriteProperty("ImageSize", m_ImageSize)
    Call PropBag.WriteProperty("Font", Font, DefFont)
    Call PropBag.WriteProperty("Caption", sCaption, m_def_Caption)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("ShowCustomMenu", m_ShowCustomMenu, m_def_ShowCustomMenu)
    Call PropBag.WriteProperty("Theme", m_Theme, 1)
    Call PropBag.WriteProperty("UsePermissions", mvarUsePermissions, True)
    Err.Clear
End Sub
Public Function AddTab(zID As String, zCaption As String, zVisible As Boolean) As Long
    On Error Resume Next
    If UsePermissions = False Then GoTo AddIt
    Dim strPrefix As String
    Dim strSuffix As String
    strPrefix = MvField(zID, 1, "_")
    Select Case strPrefix
    Case "openportfolio"
        strSuffix = MvField(zID, -1, "-")
        If IsNumeric(strSuffix) = True Then
        Else
            strSuffix = MvField(zID, 3, "_")
            If MvSearch(Permissions, "openportfolio_" & strSuffix, ";") = 0 Then Exit Function
        End If
    Case Else
        If MvSearch(Permissions, zID, ";") = 0 Then Exit Function
    End Select
AddIt:
    zCaption = ProperCase(Replace$(zCaption, "&", "&&"))
    TotalTabs = TotalTabs + 1
    ReDim Preserve TabButtons(TotalTabs - 1)
    TabButtons(TotalTabs - 1).TabID = zID
    zCaption = Replace$(zCaption, vbNewLine, " ")
    TabButtons(TotalTabs - 1).TabCaption = zCaption
    TabButtons(TotalTabs - 1).TabVisible = zVisible
    If Len(TabSelected) = 0 Then
        TabSelected = zID
    End If
    AddTab = TotalTabs - 1
    TabsUpdate
    Err.Clear
End Function
Public Sub EditTab(tabPos As Integer, ByVal zCaption As String)
    On Error Resume Next
    zCaption = ProperCase(Replace$(zCaption, "&", "&&"))
    zCaption = Replace$(zCaption, vbNewLine, " ")
    TabButtons(tabPos).TabCaption = zCaption
    TabsUpdate
    Err.Clear
End Sub
Public Sub RenameTab(ByVal zTab As String, ByVal zCaption As String)
    On Error Resume Next
    Dim tabPos As Integer
    tabPos = TabSearch(zTab)
    If tabPos >= 0 Then EditTab tabPos, zCaption
    Err.Clear
End Sub
Public Sub AddCat(zID As String, zTab As String, zCaption As String, zDlgButton As Boolean, ByVal zToolTip As String)
    On Error Resume Next
    If UsePermissions = False Then GoTo AddIt
    Dim strPrefix As String
    Dim strSuffix As String
    strPrefix = MvField(zID, 1, "_")
    Select Case strPrefix
    Case "openportfolio"
        strSuffix = MvField(zID, -1, "-")
        If IsNumeric(strSuffix) = True Then
        Else
            strSuffix = MvField(zID, 3, "_")
            If MvSearch(Permissions, "openportfolio_" & strSuffix, ";") = 0 Then Exit Sub
        End If
    Case Else
        If MvSearch(Permissions, zID, ";") = 0 Then Exit Sub
    End Select
AddIt:
    zCaption = ProperCase(Replace$(zCaption, "&", "&&"))
    TotalCats = TotalCats + 1
    ReDim Preserve CategoryButtons(TotalCats - 1)
    CategoryButtons(TotalCats - 1).CatsID = zID
    CategoryButtons(TotalCats - 1).CatsT = zTab
    zCaption = Replace$(zCaption, vbNewLine, " ")
    CategoryButtons(TotalCats - 1).CatsC = zCaption
    CategoryButtons(TotalCats - 1).CatsD = zDlgButton
    CategoryButtons(TotalCats - 1).CatsTool = zToolTip
    CatsUpdate
    Err.Clear
End Sub
Public Sub AddTopButton(zID As String, zCaption As String, zPicture As Variant, zToolTip As String)
    On Error Resume Next
    If UsePermissions = False Then GoTo AddIt
    Dim strPrefix As String
    Dim strSuffix As String
    strPrefix = MvField(zID, 1, "_")
    Select Case strPrefix
    Case "openportfolio"
        strSuffix = MvField(zID, -1, "-")
        If IsNumeric(strSuffix) = True Then
        Else
            strSuffix = MvField(zID, 3, "_")
            If MvSearch(Permissions, "openportfolio_" & strSuffix, ";") = 0 Then Exit Sub
        End If
    Case Else
        If MvSearch(Permissions, zID, ";") = 0 Then Exit Sub
    End Select
AddIt:
    TotalTopButton = TotalTopButton + 1
    ReDim Preserve TopButtons(TotalTopButton - 1)
    TopButtons(TotalTopButton - 1).TopBID = zID
    TopButtons(TotalTopButton - 1).TopBC = zCaption
    If TotalTopButton <> 1 Then
        Load RibbonTopImage(TotalTopButton - 1)
        Load RibbonTop_over(TotalTopButton - 1)
        Load TBMouse(TotalTopButton - 1)
    End If
    TBMouse(TotalTopButton - 1).Top = 0
    RibbonTop_over(TotalTopButton - 1).Top = 0
    RibbonTop_over(TotalTopButton - 1).Left = RibbonTop.Left + (330 * (TotalTopButton - 1))
    TBMouse(TotalTopButton - 1).Left = RibbonTop_over(TotalTopButton - 1).Left
    Set RibbonTopImage(TotalTopButton - 1).Picture = zImg.ListImages.Item(GetIconIndex(zImg, zPicture)).Picture
    RibbonTopImage(TotalTopButton - 1).Top = (RibbonTop.Height - RibbonTopImage(TotalTopButton - 1).Height) / 2
    'ct = (RibbonTop_over(TotalTopButton - 1).Width - RibbonTopImage(TotalTopButton - 1).Width) / 2
    ' for some reasons, the picture for the first top button
    ' is not always correct, this is a fix
    If TotalTopButton - 1 = 0 Then
        RibbonTopImage(TotalTopButton - 1).Height = RibbonTop_over(TotalTopButton - 1).Height - 60
    Else
        RibbonTopImage(TotalTopButton - 1).Height = RibbonTop_over(TotalTopButton - 1).Height - 120
    End If
    RibbonTopImage(TotalTopButton - 1).Left = RibbonTop_over(TotalTopButton - 1).Left + 30
    RibbonTopImage(TotalTopButton - 1).Width = RibbonTop_over(TotalTopButton - 1).Width - 60
    RibbonTopImage(TotalTopButton - 1).Top = RibbonTop_over(TotalTopButton - 1).Top + 60
    RibbonTop_over(TotalTopButton - 1).Visible = False
    RibbonTop_over(TotalTopButton - 1).ZOrder 0
    RibbonTopImage(TotalTopButton - 1).Visible = True
    RibbonTopImage(TotalTopButton - 1).ZOrder 0
    RibbonTopImage(TotalTopButton - 1).Stretch = True
    If Len(zToolTip) = 0 Or zToolTip = Null Then
        If InStr(zCaption, vbNewLine) Then
            zCaption = Replace$(zCaption, vbNewLine, " ")
        End If
        TBMouse(TotalTopButton - 1).ToolTipText = zCaption
    Else
        zToolTip = Replace$(zToolTip, vbNewLine, " ")
        TBMouse(TotalTopButton - 1).ToolTipText = zToolTip
    End If
    TBMouse(TotalTopButton - 1).Visible = True
    TBMouse(TotalTopButton - 1).ZOrder 0
    CatsUpdate
    Err.Clear
End Sub
Public Sub ResizeLogo(lngSize As Long)
    On Error Resume Next
    Logo.Stretch = True
    Logo.Height = lngSize
    Logo.Width = lngSize
    Logo.Refresh
    Err.Clear
End Sub
Public Property Get ShowCustomMenu() As Boolean
    On Error Resume Next
    ShowCustomMenu = m_ShowCustomMenu
    Err.Clear
End Property
Public Property Let ShowCustomMenu(ByVal New_ShowCustomMenu As Boolean)
    On Error Resume Next
    m_ShowCustomMenu = New_ShowCustomMenu
    PropertyChanged "ShowCustomMenu"
    Err.Clear
End Property
Private Sub RibbonTopCustom_over_Click()
    On Error Resume Next
    RaiseEvent CustomClick
    Err.Clear
End Sub
Public Sub AddButton(zID As String, zSubCat As String, zCaption As String, zPicture As Variant, zMore As Boolean, zToolTip As String, SplitCaption As Boolean)
    On Error Resume Next
    If UsePermissions = False Then GoTo AddIt
    Dim strPrefix As String
    Dim strSuffix As String
    strPrefix = MvField(zID, 1, "_")
    Select Case strPrefix
    Case "openportfolio"
        strSuffix = MvField(zID, -1, "-")
        If IsNumeric(strSuffix) = True Then
        Else
            strSuffix = MvField(zID, 3, "_")
            If MvSearch(Permissions, "openportfolio_" & strSuffix, ";") = 0 Then Exit Sub
        End If
    Case Else
        If MvSearch(Permissions, zID, ";") = 0 Then Exit Sub
    End Select
AddIt:
    zCaption = ProperCase(zCaption)
    If SplitCaption = True Then zCaption = Replace$(zCaption, " ", vbNewLine)
    TotalButton = TotalButton + 1
    ReDim Preserve RibbonButtons(TotalButton - 1)
    RibbonButtons(TotalButton - 1).TopBuID = zID
    RibbonButtons(TotalButton - 1).TopBuS = zSubCat
    RibbonButtons(TotalButton - 1).TopBuC = zCaption
    If Len(zToolTip) = 0 Or zToolTip = Null Then
        If InStr(zCaption, vbNewLine) Then
            zCaption = Replace$(zCaption, vbNewLine, " ")
        End If
        RibbonButtons(TotalButton - 1).TopBuT = zCaption
    Else
        zToolTip = Replace$(zToolTip, vbNewLine, " ")
        RibbonButtons(TotalButton - 1).TopBuT = zToolTip
    End If
    Set RibbonButtons(TotalButton - 1).TopBuI = Nothing
    If Len(zPicture) > 0 Then Set RibbonButtons(TotalButton - 1).TopBuI = zImg.ListImages.Item(GetIconIndex(zImg, zPicture)).Picture
    RibbonButtons(TotalButton - 1).TopBuG = zMore
    RibbonButtons(TotalButton - 1).TopTxt = ""
    RibbonButtons(TotalButton - 1).TopWdt = 0
    RibbonButtons(TotalButton - 1).TopType = ""
    RibbonButtons(TotalButton - 1).TopFormat = ""
    RibbonButtons(TotalButton - 1).TopBuX = ""
    CatsUpdate
    Err.Clear
End Sub
Public Sub AddComboBox(zID As String, zSubCat As String, zCaption As String, zToolTip As String, ByVal cboName As String, ByVal cboWidth As Long)
    On Error Resume Next
    If UsePermissions = False Then GoTo AddIt
    Dim strPrefix As String
    Dim strSuffix As String
    strPrefix = MvField(zID, 1, "_")
    Select Case strPrefix
    Case "openportfolio"
        strSuffix = MvField(zID, -1, "-")
        If IsNumeric(strSuffix) = True Then
        Else
            strSuffix = MvField(zID, 3, "_")
            If MvSearch(Permissions, "openportfolio_" & strSuffix, ";") = 0 Then Exit Sub
        End If
    Case Else
        If MvSearch(Permissions, zID, ";") = 0 Then Exit Sub
    End Select
AddIt:
    zCaption = ProperCase(zCaption)
    TotalButton = TotalButton + 1
    ReDim Preserve RibbonButtons(TotalButton - 1)
    RibbonButtons(TotalButton - 1).TopBuID = zID
    RibbonButtons(TotalButton - 1).TopBuS = zSubCat
    RibbonButtons(TotalButton - 1).TopBuC = zCaption
    If Len(zToolTip) = 0 Or zToolTip = Null Then
        If InStr(zCaption, vbNewLine) Then
            zCaption = Replace$(zCaption, vbNewLine, " ")
        End If
        RibbonButtons(TotalButton - 1).TopBuT = zCaption
    Else
        zToolTip = Replace$(zToolTip, vbNewLine, " ")
        RibbonButtons(TotalButton - 1).TopBuT = zToolTip
    End If
    Set RibbonButtons(TotalButton - 1).TopBuI = Nothing
    RibbonButtons(TotalButton - 1).TopBuG = False
    RibbonButtons(TotalButton - 1).TopTxt = cboName
    RibbonButtons(TotalButton - 1).TopWdt = cboWidth
    RibbonButtons(TotalButton - 1).TopType = "c"
    RibbonButtons(TotalButton - 1).TopFormat = ""
    CatsUpdate
    Err.Clear
End Sub
Public Sub TabShow(ByVal zID As String)
    On Error Resume Next
    Dim myLocation As Integer
    myLocation = TabSearch(zID)
    If myLocation <> -1 Then
        SaveSetting App.Title, "click", "tab", zID
        TabButtons(myLocation).TabVisible = True
        Me.Refresh
        TabMouse_Click myLocation
    End If
    Err.Clear
End Sub
Public Sub TabHide(ByVal zID As String)
    On Error Resume Next
    Dim myLocation As Integer
    myLocation = TabSearch(zID)
    If myLocation <> -1 Then
        TabButtons(myLocation).TabVisible = False
        Me.Refresh
        TabMouse_Click 0
    End If
    Err.Clear
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
'Public Property Get Theme() As ThemeEnum
'
'    Theme = m_Theme
'
'End Property
Public Property Let Theme(ByVal New_Theme As ThemeEnum)
    On Error Resume Next
    'Freeze
    m_Theme = New_Theme
    PropertyChanged "Theme"
    LoadTheme m_Theme
    'Freeze False
    Err.Clear
End Property
Public Sub LoadTheme(iTema As ThemeEnum)
    On Error Resume Next
    Dim Id As String
    Select Case iTema
    Case 0
        Id = "BLACK"
        Titulo.ForeColor = &HFFFFFF
        Titulo2.ForeColor = &HFFD18A
        Cat_Caption(0).ForeColor = &HFFFFFF
        TAB_NORMAL = vbWhite
        TAB_SELECTED = vbBlack
        Button_Caption(0).ForeColor = &H80000008
    Case 1
        Id = "BLUE"
        Titulo.ForeColor = &H797069
        Titulo2.ForeColor = &HB86A3E
        Cat_Caption(0).ForeColor = &HB86A3E
        TAB_NORMAL = &H8B4215
        TAB_SELECTED = &H8B4215
        Button_Caption(0).ForeColor = &H8B4215
    Case 2
        Id = "SILVER"
        Titulo.ForeColor = &H6A625C
        Titulo2.ForeColor = &HB86A3E
        Cat_Caption(0).ForeColor = &H6A625C
        TAB_NORMAL = &H6A625C
        TAB_SELECTED = &H6A625C
        Button_Caption(0).ForeColor = &H6A625C
    Case Else
        Id = "BLACK"
    End Select
    Set Barra2.Picture = LoadResourcePicture(101, Id)
    Set BarraLeft.Picture = LoadResourcePicture(102, Id)
    Set BarraRight.Picture = LoadResourcePicture(103, Id)
    Set Minoff.Picture = LoadResourcePicture(104, Id)
    Set Minon.Picture = LoadResourcePicture(105, Id)
    Set Maxoff.Picture = LoadResourcePicture(106, Id)
    Set Maxon.Picture = LoadResourcePicture(107, Id)
    Set Endoff.Picture = LoadResourcePicture(108, Id)
    Set Endon.Picture = LoadResourcePicture(109, Id)
    Set ButtonRibbonoff.Picture = LoadResourcePicture(110, Id)
    Set ButtonRibbonover.Picture = LoadResourcePicture(111, Id)
    Set ButtonRibbonon.Picture = LoadResourcePicture(112, Id)
    Set RibbonTop.Picture = LoadResourcePicture(113, Id)
    Set RibbonTopRight.Picture = LoadResourcePicture(114, Id)
    Set RibbonTopCustom.Picture = LoadResourcePicture(115, Id)
    Set RibbonTopCustom_over.Picture = LoadResourcePicture(116, Id)
    Set RibbonTop_over(0).Picture = LoadResourcePicture(117, Id)
    Set Cat_Dlg(0).Picture = LoadResourcePicture(118, Id)
    Set Cat_Dlg_on(0).Picture = LoadResourcePicture(119, Id)
    Set Cat_Dlg_over(0).Picture = LoadResourcePicture(120, Id)
    Set Cat_Left_off(0).Picture = LoadResourcePicture(121, Id)
    Set Cat_Center_off(0).Picture = LoadResourcePicture(122, Id)
    Set Cat_Right_off(0).Picture = LoadResourcePicture(123, Id)
    Set Cat_Left_on(0).Picture = LoadResourcePicture(124, Id)
    Set Cat_Center_on(0).Picture = LoadResourcePicture(125, Id)
    Set Cat_Right_on(0).Picture = LoadResourcePicture(126, Id)
    Set Tab_left(0).Picture = LoadResourcePicture(127, Id)
    Set Tab_center(0).Picture = LoadResourcePicture(128, Id)
    Set Tab_right(0).Picture = LoadResourcePicture(129, Id)
    Set Tab_left_over(0).Picture = LoadResourcePicture(130, Id)
    Set Tab_center_over(0).Picture = LoadResourcePicture(131, Id)
    Set Tab_right_over(0).Picture = LoadResourcePicture(132, Id)
    Set Glip_off(0).Picture = LoadResourcePicture(133, Id)
    Set Glip_on(0).Picture = LoadResourcePicture(134, Id)
    Set Button_left_over(0).Picture = LoadResourcePicture(135, Id)
    Set Button_center_over(0).Picture = LoadResourcePicture(136, Id)
    Set Button_right_over(0).Picture = LoadResourcePicture(137, Id)
    Set Button_left(0).Picture = LoadResourcePicture(138, Id)
    Set Button_center(0).Picture = LoadResourcePicture(139, Id)
    Set Button_right(0).Picture = LoadResourcePicture(140, Id)
    Err.Clear
End Sub
Private Function TempFileName(ByVal strPrefix As String) As String
    On Error Resume Next
    Dim fs As FileSystemObject
    Set fs = New FileSystemObject
    Dim strTempFolder As String
    strTempFolder = fs.GetSpecialFolder(Scripting.TemporaryFolder)
    TempFileName = strTempFolder & "\" & strPrefix & fs.GetTempName()
    Set fs = Nothing
    Err.Clear
End Function
Public Function LoadResourcePicture(ByVal Id As Variant, ByVal Format As Variant) As IPicture
    On Error Resume Next
    Dim sFile As String
    Dim b() As Byte
    Dim iFile As Integer
    On Error GoTo ErrorHandler
    b = LoadResData(Id, Format)
    sFile = TempFileName("LRP")
    iFile = FreeFile
    Open sFile For Binary Access Write Lock Read As #iFile
    Put #iFile, , b
    Close #iFile
    iFile = 0
    Set LoadResourcePicture = LoadPicture(sFile)
    KillFile sFile
    Erase b
    Err.Clear
    Exit Function
ErrorHandler:
    Dim lErr As Long
    Dim sErr As String
    lErr = Err.Number:   sErr = Err.Description
    If Not iFile = 0 Then Close #iFile
    KillFile sFile
    Err.Raise Err.Number, App.EXEName & ".LoadResourcePicture", Err.Description
    Err.Clear
    Exit Function
    Err.Clear
End Function
Private Sub KillFile(ByVal sFile As String)
    On Error Resume Next
    Kill sFile
    Err.Clear
End Sub
Public Sub Resize()
    On Error Resume Next
    Dim InicioArea As Long
    Dim area As Long
    Dim pos As Long
    UserControl.Height = Barra2.Height
    UserControl.Width = UserControl.ParentControls.Item(0).ScaleWidth
    'If TypeName(mParent) <> "Nothing" Then UserControl.Width = mParent.ScaleWidth
    Barra2.Width = UserControl.Width
    BarraRight.Left = Barra2.Width - BarraRight.Width
    ButtonRibbon.Top = 0
    ButtonRibbon.Left = 0
    ButtonRibbonoff.Top = 0
    ButtonRibbonover.Top = 0
    ButtonRibbonon.Top = 0
    ButtonRibbonoff.Left = 0
    ButtonRibbonover.Left = 0
    ButtonRibbonon.Left = 0
    Logo.Top = (ButtonRibbonoff.Height - Logo.Height) / 2
    Logo.Left = Logo.Top
    RibbonTop.Top = 0
    RibbonTop.Left = ButtonRibbonoff.Width
    If TotalTopButton >= 1 Then RibbonTopImage(TotalTopButton - 1).Top = (RibbonTop.Height - RibbonTopImage(TotalTopButton - 1).Height) / 2
    RibbonTop.Width = 330 * TotalTopButton
    RibbonTopRight.Top = 0
    RibbonTopRight.Left = RibbonTop.Left + RibbonTop.Width
    RibbonTopCustom.Top = 0
    RibbonTopCustom.Left = RibbonTopRight.Left + RibbonTopRight.Width
    RibbonTopCustom_over.Top = 0
    RibbonTopCustom_over.Left = RibbonTopCustom.Left
    If m_ShowCustomMenu = True Then
        RibbonTopCustom.Visible = True
        InicioArea = (RibbonTopCustom.Left + RibbonTopCustom.Width)
    Else
        RibbonTopCustom.Visible = False
        InicioArea = (RibbonTopRight.Left + RibbonTopRight.Width)
    End If
    area = UserControl.Width - (InicioArea + (Endoff.Width * 3))
    Barra.Left = InicioArea
    If area >= 0 Then Barra.Width = area
    pos = InStr(sCaption, " - ")
    If pos > 0 Then
        Titulo.Caption = Mid$(sCaption, 1, pos + 2)
        Titulo2.Caption = Mid$(sCaption, pos + 3)
        Titulo.Left = ((area - (Titulo.Width + Titulo2.Width)) / 2) + InicioArea
        Titulo2.Left = Titulo.Left + Titulo.Width
        Titulo2.Visible = True
    Else
        Titulo.Caption = sCaption
        Titulo.Left = ((area - Titulo.Width) / 2) + InicioArea
        Titulo2.Visible = False
    End If
    Endoff.Left = Barra2.Width - Endoff.Width
    Endon.Left = Endoff.Left
    Maxoff.Left = Endoff.Left - Maxoff.Width
    Maxon.Left = Maxoff.Left
    Minoff.Left = Maxoff.Left - Minoff.Width
    Minon.Left = Minoff.Left
    Err.Clear
End Sub
Public Property Let ImageList(ByVal zImageList As Variant)
    On Error Resume Next
    Set zImg = zImageList
    If TypeName(zImg) = "ImageList" Then
        iImgLType = 1
    ElseIf TypeName(zImageList) = "vbalImageList" Then
        iImgLType = 2
    Else
        iImgLType = 0
    End If
    Err.Clear
End Property
Public Property Let Icon(ByVal zPicture As Variant)
    On Error Resume Next
    Set Logo.Picture = zImg.ListImages.Item(GetIconIndex(zImg, zPicture)).Picture
    Err.Clear
End Property
Private Function GetIconIndex(zImg As Variant, iIcon As Variant) As Integer
    On Error Resume Next
    Dim i As Integer
    Dim iLCnt As Integer
    'Parameter NOT string or integer?
    If (VarType(iIcon) <> vbInteger) And (VarType(iIcon) <> vbString) Then
        GetIconIndex = -1
        Err.Clear
        Exit Function
    End If
    iLCnt = zImg.ListImages.Count
    'Key was passed
    If VarType(iIcon) = vbString Then
        'get icon index
        For i = 1 To iLCnt
            If LCase$(zImg.ListImages(i).Key) = LCase$(iIcon) Then
                'we did find the Icons index
                GetIconIndex = i
                Err.Clear
                Exit Function
            End If
        Next
        'when we got here the string doesn't match
        GetIconIndex = -1
        Err.Clear
        Exit Function
    End If
    'Index was passed
    If iIcon >= 1 Or iIcon <= iLCnt Then
        GetIconIndex = iIcon
    Else
        'RaiseWarning "GetIconIndex", "GetIconIndex: invalid Image Index."
        GetIconIndex = -1
    End If
    Err.Clear
    Exit Function
NoImage:
    'No imagelist was selected
    GetIconIndex = -1
    Err.Clear
End Function
'Private Sub RaiseError(sErrorDescription As String)
'
'    MsgBox "An Error has occurred." & vbCrLf & sErrorDescription, vbCritical, "Ribbon"
'
'End Sub
Public Property Set ParentForm(newForm As Variant)
    On Error Resume Next
    Set mParent = newForm
    Err.Clear
End Property
Public Property Set Font(newFont As StdFont)
    On Error Resume Next
    Dim tmpCtl As Control
    UserControl.Font.Name = newFont.Name
    UserControl.Font.Size = newFont.Size
    UserControl.Font.Charset = newFont.Charset
    For Each tmpCtl In UserControl.Controls
        tmpCtl.Font = newFont.Font
    Next
    UserControl.Refresh
    Err.Clear
End Property
Public Property Get Font() As StdFont
    On Error Resume Next
    Set Font = UserControl.Font
    Err.Clear
End Property
Private Sub ButMouse_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    'Freeze
    Dim i As Long
    For i = 0 To ButMouse.UBound
        If i <> Index Then
            Button_left(i).Visible = False
            Button_center(i).Visible = False
            Button_right(i).Visible = False
            If Glip_off(i).Visible = True Then
                Glip_on(i).Visible = False
            End If
        End If
    Next
    If Button_left(Index).Visible = False Then
        Button_left(Index).Visible = True
        Button_center(Index).Visible = True
        Button_right(Index).Visible = True
        If Glip_off(Index).Visible = True Then
            Glip_on(Index).Visible = True
        End If
    End If
    For i = 0 To CatMouse.UBound
        If Cat_Dlg(i).Visible = True Then
            Cat_Dlg_over(i).Visible = False
        End If
    Next
    For i = 0 To TabMouse.UBound
        Tab_center_over(i).Visible = False
        Tab_left_over(i).Visible = False
        Tab_right_over(i).Visible = False
    Next
    For i = 0 To TBMouse.UBound
        RibbonTop_over(i).Visible = False
    Next
    RibbonTopCustom_over.Visible = False
    Endon.Visible = False
    Maxon.Visible = False
    Minon.Visible = False
    ButtonRibbonover.Visible = False
    ButtonRibbonon.Visible = False
    'Freeze False
    Err.Clear
End Sub
Private Sub ButtonRibbon_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    'Freeze
    Dim i As Integer
    Dim KL As Integer
    ButtonRibbonover.Visible = True
    ButtonRibbonon.Visible = False
    For i = 0 To CatMouse.UBound
        Cat_Center_on(i).Visible = False
        Cat_Left_on(i).Visible = False
        Cat_Right_on(i).Visible = False
        If Cat_Dlg(i).Visible = True Then
            Cat_Dlg_on(i).Visible = False
            Cat_Dlg_over(i).Visible = False
        End If
    Next
    For i = 0 To TabMouse.UBound
        Tab_center_over(i).Visible = False
        Tab_left_over(i).Visible = False
        Tab_right_over(i).Visible = False
    Next
    For KL = 0 To ButMouse.UBound
        Button_left(KL).Visible = False
        Button_right(KL).Visible = False
        Button_center(KL).Visible = False
    Next
    For i = 0 To TBMouse.UBound
        RibbonTop_over(i).Visible = False
    Next
    RibbonTopCustom_over.Visible = False
    Endon.Visible = False
    Maxon.Visible = False
    Minon.Visible = False
    'Freeze False
    Err.Clear
End Sub
Private Sub Cat_Dlg_on_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    'Freeze
    Dim KL As Integer
    Cat_Dlg_over(Index).Visible = True
    For KL = 0 To ButMouse.UBound
        Button_left(KL).Visible = False
        Button_right(KL).Visible = False
        Button_center(KL).Visible = False
    Next
    'Freeze False
    Err.Clear
End Sub
Private Sub CatMouse_Click(Index As Integer)
    On Error Resume Next
    'Freeze
    Dim i As Integer
    For i = 0 To CatMouse.UBound
        Cat_Center_on(i).Visible = False
        Cat_Left_on(i).Visible = False
        Cat_Right_on(i).Visible = False
    Next
    Cat_Center_on(Index).Visible = True
    Cat_Left_on(Index).Visible = True
    Cat_Right_on(Index).Visible = True
    'Freeze False
    Err.Clear
End Sub
Private Sub CatMouse_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    'Freeze
    Dim i As Integer
    Dim KL As Integer
    For i = 0 To CatMouse.UBound
        If i = Index Then
            If Cat_Center_on(i).Visible = False Then
                Cat_Center_on(Index).Visible = True
                Cat_Left_on(Index).Visible = True
                Cat_Right_on(Index).Visible = True
                If Cat_Dlg(i).Visible = True Then
                    Cat_Dlg_on(Index).Visible = True
                End If
            End If
        Else
            Cat_Center_on(i).Visible = False
            Cat_Left_on(i).Visible = False
            Cat_Right_on(i).Visible = False
            If Cat_Dlg(i).Visible = True Then
                Cat_Dlg_on(i).Visible = False
                Cat_Dlg_over(i).Visible = False
            End If
        End If
        If Cat_Dlg(i).Visible = True Then
            Cat_Dlg_over(i).Visible = False
        End If
    Next
    For KL = 0 To ButMouse.UBound
        Button_left(KL).Visible = False
        Button_right(KL).Visible = False
        Button_center(KL).Visible = False
    Next
    For i = 0 To TabMouse.UBound
        Tab_center_over(i).Visible = False
        Tab_left_over(i).Visible = False
        Tab_right_over(i).Visible = False
    Next
    For i = 0 To TBMouse.UBound
        RibbonTop_over(i).Visible = False
    Next
    RibbonTopCustom_over.Visible = False
    Endon.Visible = False
    Maxon.Visible = False
    Minon.Visible = False
    ButtonRibbonover.Visible = False
    ButtonRibbonon.Visible = False
    'Freeze False
    Err.Clear
End Sub
Private Sub Minoff_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    'Freeze
    Dim i As Integer
    Dim KL As Integer
    Endon.Visible = False
    Maxon.Visible = False
    Minon.Visible = True
    For i = 0 To TabMouse.UBound
        Tab_center_over(i).Visible = False
        Tab_left_over(i).Visible = False
        Tab_right_over(i).Visible = False
    Next
    For i = 0 To CatMouse.UBound
        Cat_Center_on(i).Visible = False
        Cat_Left_on(i).Visible = False
        Cat_Right_on(i).Visible = False
        If Cat_Dlg(i).Visible = True Then
            Cat_Dlg_on(i).Visible = False
            Cat_Dlg_over(i).Visible = False
        End If
    Next
    For KL = 0 To ButMouse.UBound
        Button_left(KL).Visible = False
        Button_right(KL).Visible = False
        Button_center(KL).Visible = False
    Next
    For i = 0 To TBMouse.UBound
        RibbonTop_over(i).Visible = False
    Next
    RibbonTopCustom_over.Visible = False
    ButtonRibbonover.Visible = False
    ButtonRibbonon.Visible = False
    'Freeze False
    Err.Clear
End Sub
Private Sub Maxoff_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    'Freeze
    Dim i As Integer
    Dim KL As Integer
    Endon.Visible = False
    Maxon.Visible = True
    Minon.Visible = False
    For i = 0 To TabMouse.UBound
        Tab_center_over(i).Visible = False
        Tab_left_over(i).Visible = False
        Tab_right_over(i).Visible = False
    Next
    For i = 0 To CatMouse.UBound
        Cat_Center_on(i).Visible = False
        Cat_Left_on(i).Visible = False
        Cat_Right_on(i).Visible = False
        If Cat_Dlg(i).Visible = True Then
            Cat_Dlg_on(i).Visible = False
            Cat_Dlg_over(i).Visible = False
        End If
    Next
    For KL = 0 To ButMouse.UBound
        Button_left(KL).Visible = False
        Button_right(KL).Visible = False
        Button_center(KL).Visible = False
    Next
    For i = 0 To TBMouse.UBound
        RibbonTop_over(i).Visible = False
    Next
    RibbonTopCustom_over.Visible = False
    ButtonRibbonover.Visible = False
    ButtonRibbonon.Visible = False
    'Freeze False
    Err.Clear
End Sub
Private Sub Endoff_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    'Freeze
    Dim i As Integer
    Dim KL As Integer
    Endon.Visible = True
    Maxon.Visible = False
    Minon.Visible = False
    For i = 0 To TabMouse.UBound
        Tab_center_over(i).Visible = False
        Tab_left_over(i).Visible = False
        Tab_right_over(i).Visible = False
    Next
    For i = 0 To CatMouse.UBound
        Cat_Center_on(i).Visible = False
        Cat_Left_on(i).Visible = False
        Cat_Right_on(i).Visible = False
        If Cat_Dlg(i).Visible = True Then
            Cat_Dlg_on(i).Visible = False
            Cat_Dlg_over(i).Visible = False
        End If
    Next
    For KL = 0 To ButMouse.UBound
        Button_left(KL).Visible = False
        Button_right(KL).Visible = False
        Button_center(KL).Visible = False
    Next
    For i = 0 To TBMouse.UBound
        RibbonTop_over(i).Visible = False
    Next
    RibbonTopCustom_over.Visible = False
    ButtonRibbonover.Visible = False
    ButtonRibbonon.Visible = False
    'Freeze False
    Err.Clear
End Sub
Private Sub RibbonTopCustom_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    'Freeze
    Dim i As Integer
    Dim KL As Integer
    RibbonTopCustom_over.Visible = True
    For i = 0 To CatMouse.UBound
        Cat_Center_on(i).Visible = False
        Cat_Left_on(i).Visible = False
        Cat_Right_on(i).Visible = False
        If Cat_Dlg(i).Visible = True Then
            Cat_Dlg_on(i).Visible = False
            Cat_Dlg_over(i).Visible = False
        End If
    Next
    For i = 0 To TabMouse.UBound
        Tab_center_over(i).Visible = False
        Tab_left_over(i).Visible = False
        Tab_right_over(i).Visible = False
    Next
    For KL = 0 To ButMouse.UBound
        Button_left(KL).Visible = False
        Button_right(KL).Visible = False
        Button_center(KL).Visible = False
    Next
    For i = 0 To TBMouse.UBound
        RibbonTop_over(i).Visible = False
    Next
    ButtonRibbonover.Visible = False
    ButtonRibbonon.Visible = False
    Endon.Visible = False
    Maxon.Visible = False
    Minon.Visible = False
    'Freeze False
    Err.Clear
End Sub
Private Sub RibbonTopRight_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    'Freeze
    Dim i As Integer
    Dim KL As Integer
    RibbonTopCustom_over.Visible = False
    For i = 0 To CatMouse.UBound
        Cat_Center_on(i).Visible = False
        Cat_Left_on(i).Visible = False
        Cat_Right_on(i).Visible = False
        If Cat_Dlg(i).Visible = True Then
            Cat_Dlg_on(i).Visible = False
            Cat_Dlg_over(i).Visible = False
        End If
    Next
    For i = 0 To TabMouse.UBound
        Tab_center_over(i).Visible = False
        Tab_left_over(i).Visible = False
        Tab_right_over(i).Visible = False
    Next
    For KL = 0 To ButMouse.UBound
        Button_left(KL).Visible = False
        Button_right(KL).Visible = False
        Button_center(KL).Visible = False
    Next
    For i = 0 To TBMouse.UBound
        RibbonTop_over(i).Visible = False
    Next
    ButtonRibbonover.Visible = False
    ButtonRibbonon.Visible = False
    Endon.Visible = False
    Maxon.Visible = False
    Minon.Visible = False
    'Freeze False
    Err.Clear
End Sub
Private Sub TabMouse_Click(Index As Integer)
    On Error Resume Next
    Dim i As Integer
    For i = 0 To TabMouse.UBound
        Tab_center_over(i).Visible = False
        Tab_left_over(i).Visible = False
        Tab_right_over(i).Visible = False
        Tab_center(i).Visible = False
        Tab_left(i).Visible = False
        Tab_right(i).Visible = False
        Tab_caption(i).ForeColor = TAB_NORMAL
    Next
    Tab_caption(Index).ForeColor = TAB_SELECTED
    Tab_center(Index).Visible = True
    Tab_left(Index).Visible = True
    Tab_right(Index).Visible = True
    TabSelected = TabButtons(Index).TabID
    CatsUpdate
    RaiseEvent TabClick(TabButtons(Index).TabID, TabButtons(Index).TabCaption)
    Tab_right(Index).ZOrder 0
    'Me.DisableUpdates False
    Err.Clear
End Sub
Public Sub TabSelect(Index As Integer)
    On Error Resume Next
    Dim i As Integer
    For i = 0 To TabMouse.UBound
        Tab_center_over(i).Visible = False
        Tab_left_over(i).Visible = False
        Tab_right_over(i).Visible = False
        Tab_center(i).Visible = False
        Tab_left(i).Visible = False
        Tab_right(i).Visible = False
        Tab_caption(i).ForeColor = TAB_NORMAL
    Next
    Tab_caption(Index).ForeColor = TAB_SELECTED
    Tab_center(Index).Visible = True
    Tab_left(Index).Visible = True
    Tab_right(Index).Visible = True
    TabSelected = TabButtons(Index).TabID
    Err.Clear
End Sub
Private Sub TabMouse_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    'Freeze
    Dim i As Integer
    Dim KL As Integer
    For i = 0 To TabMouse.UBound
        If i = Index Then
            If Tab_center(i).Visible = False Then
                Tab_center_over(Index).Visible = True
                Tab_left_over(Index).Visible = True
                Tab_right_over(Index).Visible = True
            End If
        Else
            Tab_center_over(i).Visible = False
            Tab_left_over(i).Visible = False
            Tab_right_over(i).Visible = False
        End If
    Next
    For i = 0 To CatMouse.UBound
        Cat_Center_on(i).Visible = False
        Cat_Left_on(i).Visible = False
        Cat_Right_on(i).Visible = False
        If Cat_Dlg(i).Visible = True Then
            Cat_Dlg_on(i).Visible = False
            Cat_Dlg_over(i).Visible = False
        End If
    Next
    For KL = 0 To ButMouse.UBound
        Button_left(KL).Visible = False
        Button_right(KL).Visible = False
        Button_center(KL).Visible = False
    Next
    For i = 0 To TBMouse.UBound
        RibbonTop_over(i).Visible = False
    Next
    RibbonTopCustom_over.Visible = False
    Endon.Visible = False
    Maxon.Visible = False
    Minon.Visible = False
    ButtonRibbonover.Visible = False
    ButtonRibbonon.Visible = False
    'Freeze False
    Err.Clear
End Sub
Private Sub TBMouse_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    'Freeze
    Dim i As Integer
    Dim KL As Integer
    For i = 0 To TBMouse.UBound
        RibbonTop_over(i).Visible = False
    Next
    RibbonTop_over(Index).Visible = True
    For i = 0 To TabMouse.UBound
        Tab_center_over(i).Visible = False
        Tab_left_over(i).Visible = False
        Tab_right_over(i).Visible = False
    Next
    For KL = 0 To ButMouse.UBound
        Button_left(KL).Visible = False
        Button_right(KL).Visible = False
        Button_center(KL).Visible = False
    Next
    For i = 0 To CatMouse.UBound
        Cat_Center_on(i).Visible = False
        Cat_Left_on(i).Visible = False
        Cat_Right_on(i).Visible = False
        If Cat_Dlg(i).Visible = True Then
            Cat_Dlg_on(i).Visible = False
            Cat_Dlg_over(i).Visible = False
        End If
    Next
    RibbonTopCustom_over.Visible = False
    Endon.Visible = False
    Maxon.Visible = False
    Minon.Visible = False
    ButtonRibbonover.Visible = False
    ButtonRibbonon.Visible = False
    'Freeze False
    Err.Clear
End Sub
Private Sub TabsUpdate()
    On Error Resume Next
    Dim i As Integer
    Dim tTabs As Integer
    tTabs = Tab_caption.Count - 1
    If tTabs >= 1 Then
        For i = 1 To tTabs
            Unload Tab_caption(i)
            Unload Tab_left(i)
            Unload Tab_center(i)
            Unload Tab_right(i)
            Unload Tab_left_over(i)
            Unload Tab_center_over(i)
            Unload Tab_right_over(i)
            Unload TabMouse(i)
        Next
    End If
    'DoEvents
    For i = 0 To (TotalTabs - 1)
        If i <> 0 Then
            Load Tab_caption(i)
            Load Tab_left(i)
            Load Tab_center(i)
            Load Tab_right(i)
            Load Tab_left_over(i)
            Load Tab_center_over(i)
            Load Tab_right_over(i)
            Load TabMouse(i)
            Tab_left(i).Left = Tab_right(i - 1).Left + Tab_right(i).Width
        Else
            Tab_left(0).Left = ButtonRibbon.Width
        End If
        TabMouse(i).Left = Tab_left(i).Left
        Tab_caption(i).Top = 395 + 60
        Tab_center(i).Top = 395
        Tab_left(i).Top = 395
        Tab_right(i).Top = 395
        Tab_center_over(i).Top = 395
        Tab_left_over(i).Top = 395
        Tab_right_over(i).Top = 395
        TabMouse(i).Top = 395
        Tab_caption(i) = TabButtons(i).TabCaption
        Tab_center(i).Width = Tab_caption(i).Width
        Tab_center(i).Left = Tab_left(i).Left + Tab_left(i).Width
        Tab_caption(i).Left = Tab_center(i).Left
        Tab_right(i).Left = Tab_center(i).Left + Tab_center(i).Width
        Tab_center_over(i).Width = Tab_center(i).Width
        Tab_center_over(i).Left = Tab_center(i).Left
        Tab_left_over(i).Left = Tab_left(i).Left
        Tab_right_over(i).Left = Tab_right(i).Left
        TabMouse(i).Width = Tab_left(i).Width + Tab_right(i).Width + Tab_center(i).Width
        Tab_caption(i).ForeColor = TAB_NORMAL
        Tab_caption(i).Visible = True
        If i = 0 Then
            Tab_center(i).Visible = True
            Tab_left(i).Visible = True
            Tab_right(i).Visible = True
            Tab_caption(i).ForeColor = TAB_SELECTED
        End If
        TabMouse(i).Visible = TabButtons(i).TabVisible
        Tab_center(i).ZOrder 0
        Tab_left(i).ZOrder 0
        Tab_right(i).ZOrder 0
        Tab_center_over(i).ZOrder 0
        Tab_left_over(i).ZOrder 0
        Tab_right_over(i).ZOrder 0
        Tab_caption(i).ZOrder 0
        TabMouse(i).ZOrder 0
    Next
    Err.Clear
End Sub
Private Sub CatsUpdate()
    On Error Resume Next
    Dim TotalCatsT As Integer
    Dim CatsIDT() As String
    Dim CatsCT() As String
    Dim CatsTT() As String
    Dim CatsDT() As Boolean
    Dim CatsToolT() As String
    Dim i As Integer
    Dim BUTSIZE As Integer
    Dim KL As Integer
    ReDim CatsIDT(TotalCats) As String
    ReDim CatsCT(TotalCats) As String
    ReDim CatsTT(TotalCats) As String
    ReDim CatsDT(TotalCats) As Boolean
    ReDim CatsToolT(TotalCats) As String
    TotalCatsT = 0
    For i = 0 To TotalCats - 1
        If CategoryButtons(i).CatsT = TabSelected And TabSelected <> """" And CategoryButtons(i).CatsT <> """" Then
            CatsIDT(TotalCatsT) = CategoryButtons(i).CatsID
            CatsTT(TotalCatsT) = CategoryButtons(i).CatsT
            CatsCT(TotalCatsT) = CategoryButtons(i).CatsC
            CatsDT(TotalCatsT) = CategoryButtons(i).CatsD
            CatsToolT(TotalCatsT) = CategoryButtons(i).CatsTool
            TotalCatsT = TotalCatsT + 1
        End If
    Next
    'DoEvents
    If CatMouse.UBound >= 1 Then
        For i = 1 To CatMouse.UBound
            Unload Cat_Left_off(i)
            Unload Cat_Left_on(i)
            Unload Cat_Right_off(i)
            Unload Cat_Right_on(i)
            Unload Cat_Center_off(i)
            Unload Cat_Center_on(i)
            Unload Cat_Caption(i)
            Unload CatMouse(i)
            Unload Cat_Dlg(i)
            Unload Cat_Dlg_on(i)
            Unload Cat_Dlg_over(i)
        Next
    End If
    'DoEvents
    If Button_center.UBound >= 1 Then
        For i = 1 To Button_center.UBound
            Unload Button_left(i)
            Unload Button_center(i)
            Unload Button_right(i)
            Unload Button_left_over(i)
            Unload Button_center_over(i)
            Unload Button_right_over(i)
            Unload Button_Caption(i)
            Unload Button_Icon(i)
            Unload Button_Text(i)
            Unload Glip_on(i)
            Unload Glip_off(i)
            Unload ButMouse(i)
            Unload txtBox(i)
            Unload cboBox(i)
            Unload datePick(i)
            Unload progBar(i)
        Next
    End If
    Button_left(0).Visible = False
    Button_center(0).Visible = False
    Button_right(0).Visible = False
    Button_Caption(0).Visible = False
    Button_Icon(0).Visible = False
    Button_Text(0).Visible = False
    txtBox(0).Visible = False
    cboBox(0).Visible = False
    datePick(0).Visible = False
    progBar(0).Visible = False
    ButMouse(0).Visible = False
    Cat_Left_off(0).Visible = False
    Cat_Left_on(0).Visible = False
    Cat_Right_off(0).Visible = False
    Cat_Right_on(0).Visible = False
    Cat_Center_off(0).Visible = False
    Cat_Center_on(0).Visible = False
    Cat_Caption(0).Visible = False
    CatMouse(0).Visible = False
    Cat_Dlg(0).Visible = False
    Cat_Dlg_on(0).Visible = False
    Cat_Dlg_over(0).Visible = False
    For i = 0 To (TotalCatsT - 1)
        If i <> 0 Then
            Load Cat_Left_off(i)
            Load Cat_Left_on(i)
            Load Cat_Right_off(i)
            Load Cat_Right_on(i)
            Load Cat_Center_off(i)
            Load Cat_Center_on(i)
            Load Cat_Caption(i)
            Load CatMouse(i)
            Load Cat_Dlg(i)
            Load Cat_Dlg_on(i)
            Load Cat_Dlg_over(i)
            Cat_Left_off(i).Left = Cat_Right_off(i - 1).Left + Cat_Right_off(i).Width
        Else
            Cat_Left_off(i).Left = 120
        End If
        CatMouse(i).Left = Cat_Left_off(i).Left
        Cat_Caption(i).Caption = CatsCT(i)
        Cat_Caption(i).Tag = CatsIDT(i)
        Cat_Center_off(i).Left = Cat_Left_off(i).Left + Cat_Left_off(i).Width
        BUTSIZE = ButtonsUpdate(CatsIDT(i), Cat_Center_off(i).Left)
        If CatsDT(i) = True Then
            Cat_Center_off(i).Width = Cat_Caption(i).Width + Cat_Dlg(i).Width
        Else
            Cat_Center_off(i).Width = Cat_Caption(i).Width
        End If
        If Cat_Center_off(i).Width < BUTSIZE Then
            Cat_Center_off(i).Width = BUTSIZE
            Cat_Caption(i).Left = Cat_Center_off(i).Left + ((Cat_Center_off(i).Width - Cat_Caption(i).Width) / 2)
        Else
            Cat_Caption(i).Left = Cat_Center_off(i).Left
        End If
        Cat_Right_off(i).Left = Cat_Center_off(i).Left + Cat_Center_off(i).Width
        Cat_Center_on(i).Width = Cat_Center_off(i).Width
        Cat_Center_on(i).Left = Cat_Center_off(i).Left
        Cat_Left_on(i).Left = Cat_Left_off(i).Left
        Cat_Right_on(i).Left = Cat_Right_off(i).Left
        CatMouse(i).Width = Cat_Left_off(i).Width + Cat_Right_off(i).Width + Cat_Center_off(i).Width
        Cat_Caption(i).Visible = True
        Cat_Center_off(i).Visible = True
        Cat_Left_off(i).Visible = True
        Cat_Right_off(i).Visible = True
        CatMouse(i).Visible = True
        Cat_Center_off(i).ZOrder 0
        Cat_Left_off(i).ZOrder 0
        Cat_Right_off(i).ZOrder 0
        Cat_Center_on(i).ZOrder 0
        Cat_Left_on(i).ZOrder 0
        Cat_Right_on(i).ZOrder 0
        Cat_Caption(i).ZOrder 0
        CatMouse(i).ZOrder 0
        Cat_Dlg(i).Left = (Cat_Right_off(i).Left - Cat_Dlg(i).Width) + 15
        Cat_Dlg(i).Top = (Cat_Right_off(i).Top + Cat_Right_off(i).Height) - (Cat_Dlg(i).Height + 60)
        Cat_Dlg_on(i).Left = Cat_Dlg(i).Left
        Cat_Dlg_over(i).Left = Cat_Dlg(i).Left
        Cat_Dlg_on(i).Top = Cat_Dlg(i).Top
        Cat_Dlg_over(i).Top = Cat_Dlg(i).Top
        Cat_Dlg_on(i).Visible = False
        Cat_Dlg_over(i).Visible = False
        If CatsDT(i) = True Then
            Cat_Dlg(i).Visible = True
        End If
        Cat_Dlg(i).ZOrder 0
        Cat_Dlg_on(i).ZOrder 0
        Cat_Dlg_over(i).ZOrder 0
        Cat_Dlg(i).ToolTipText = CatsToolT(i)
        Cat_Dlg_on(i).ToolTipText = CatsToolT(i)
        Cat_Dlg_over(i).ToolTipText = CatsToolT(i)
    Next
    'DoEvents
    For KL = 0 To ButMouse.UBound
        Button_left(KL).Visible = False
        Button_left(KL).ZOrder 0
        Button_right(KL).Visible = False
        Button_right(KL).ZOrder 0
        Button_center(KL).Visible = False
        Button_center(KL).ZOrder 0
        Button_left_over(KL).Visible = False
        Button_left_over(KL).ZOrder 0
        Button_right_over(KL).Visible = False
        Button_right_over(KL).ZOrder 0
        Button_center_over(KL).Visible = False
        Button_center_over(KL).ZOrder 0
        Button_Icon(KL).ZOrder 0
        Button_Text(KL).ZOrder 0
        txtBox(KL).ZOrder 0
        cboBox(KL).ZOrder 0
        datePick(KL).ZOrder 0
        progBar(KL).ZOrder 0
        Button_Caption(KL).ZOrder 0
        Glip_off(KL).ZOrder 0
        Glip_on(KL).ZOrder 0
        ButMouse(KL).ZOrder 0
    Next
    ComboBoxRefresh
    Err.Clear
End Sub
Public Sub Clear()
    On Error Resume Next
    'clear the ribbon
    TotalButton = 0
    TotalCats = 0
    TotalTabs = 0
    TotalTopButton = 0
    ImageList = Nothing
    txtBox(0).Text = ""
    cboBox(0).Clear
    progBar(0).Max = 100
    progBar(0).Min = 0
    progBar(0).Value = 0
    cboMaster.Clear
    cboMenus.Clear
    Dim i As Integer
    For i = 1 To RibbonTopImage.UBound
        Unload RibbonTopImage(i)
    Next
    For i = 1 To RibbonTop_over.UBound
        Unload RibbonTop_over(i)
    Next
    For i = 1 To TBMouse.UBound
        Unload TBMouse(i)
    Next
    For i = 1 To (TotalTabs - 1)
        Unload Tab_caption(i)
        Unload Tab_left(i)
        Unload Tab_center(i)
        Unload Tab_right(i)
        Unload Tab_left_over(i)
        Unload Tab_center_over(i)
        Unload Tab_right_over(i)
        Unload TabMouse(i)
    Next
    For i = 1 To CatMouse.UBound
        Unload Cat_Left_off(i)
        Unload Cat_Left_on(i)
        Unload Cat_Right_off(i)
        Unload Cat_Right_on(i)
        Unload Cat_Center_off(i)
        Unload Cat_Center_on(i)
        Unload Cat_Caption(i)
        Unload CatMouse(i)
        Unload Cat_Dlg(i)
        Unload Cat_Dlg_on(i)
        Unload Cat_Dlg_over(i)
    Next
    For i = 1 To Button_center.UBound
        Unload Button_left(i)
        Unload Button_center(i)
        Unload Button_right(i)
        Unload Button_left_over(i)
        Unload Button_center_over(i)
        Unload Button_right_over(i)
        Unload Button_Caption(i)
        Unload Button_Icon(i)
        Unload Button_Text(i)
        Unload Glip_on(i)
        Unload Glip_off(i)
        Unload ButMouse(i)
        Unload txtBox(i)
        Unload cboBox(i)
        Unload datePick(i)
        Unload progBar(i)
    Next
    Err.Clear
End Sub
Public Function TabSearch(ByVal zID As String) As Integer
    On Error Resume Next
    ' return the position of a tab
    Dim myID As String
    Dim myLocation As Integer
    TabSearch = -1
    For myLocation = 0 To TotalTabs - 1
        myID = TabButtons(myLocation).TabID
        If LCase$(myID) = LCase$(zID) Then
            TabSearch = myLocation
            Exit For
        End If
    Next
    Err.Clear
End Function
Public Sub EditButton(ByVal zID As String, ByVal zCaption As String, zPicture As Variant, Optional zMore As Boolean = False, Optional zToolTip As String = vbNullString, Optional SplitCaption As Boolean = False, Optional ByVal zText As String = vbNullString, Optional ByVal zNewID As String = vbNullString)
    On Error Resume Next
    ' edit the contents of a button
    Dim butCnt As Integer
    Dim myID As String
    Dim myLocation As Integer
    'FreezeWindow Me
    zCaption = ProperCase(zCaption)
    myID = vbNullString
    myLocation = -1
    For butCnt = 0 To TotalButton - 1
        myID = RibbonButtons(butCnt).TopBuID
        If LCase$(myID) = LCase$(zID) Then
            myLocation = butCnt
            Exit For
        End If
    Next
    If myLocation = -1 Then Exit Sub
    If SplitCaption = True Then zCaption = Replace$(zCaption, "", vbNewLine)
    RibbonButtons(myLocation).TopBuC = zCaption
    If Len(zNewID) > 0 Then RibbonButtons(myLocation).TopBuID = zNewID
    If Len(zToolTip) = 0 Or zToolTip = Null Then
        If InStr(zCaption, vbNewLine) Then
            zCaption = Replace$(zCaption, vbNewLine, "")
        End If
        RibbonButtons(myLocation).TopBuT = zCaption
    Else
        zToolTip = Replace$(zToolTip, vbNewLine, "")
        RibbonButtons(myLocation).TopBuT = zToolTip
    End If
    Set RibbonButtons(myLocation).TopBuI = Nothing
    If Len(zPicture) > 0 Then Set RibbonButtons(myLocation).TopBuI = zImg.ListImages.Item(GetIconIndex(zImg, zPicture)).Picture
    RibbonButtons(myLocation).TopBuG = zMore
    RibbonButtons(myLocation).TopBuX = zText
    CatsUpdate
    'FreezeWindow Me, False
    Err.Clear
End Sub
Public Sub AddButtonMenu(ByVal zID As String, ByVal zMenuID As String, ByVal zMenuCaption As String, Optional bHasSubMenus As Boolean = False)
    On Error Resume Next
    ' add a menu item to a button
    Dim butCnt As Integer
    Dim myID As String
    Dim myLocation As Integer
    Dim prevKey As String
    myLocation = -1
    For butCnt = 0 To TotalButton - 1
        myID = RibbonButtons(butCnt).TopBuID
        If LCase$(myID) = LCase$(zID) Then
            myLocation = butCnt
            Exit For
        End If
    Next
    If myLocation = -1 Then Exit Sub
    RibbonButtons(myLocation).menuName = zMenuID
    If Right$(zMenuID, 1) <> "\" Then zMenuID = zMenuID & "\"
    prevKey = MvFromMv(zMenuID, 1, -2, "\")
    'Debug.Print prevKey
    cboMenus.AddItem zID & "|" & zMenuID & "|" & Replace$(zMenuCaption, "&", "&&") & "|" & IIf(bHasSubMenus = True, "1", "0")
    Err.Clear
End Sub
Public Sub EditLabel(ByVal zID As String, ByVal zCaption As String, ByVal zText As String, ByVal zToolTip As String, Optional SplitCaption As Boolean = False)
    On Error Resume Next
    ' edit the contents of the label button
    Dim butCnt As Integer
    Dim myID As String
    Dim myLocation As Integer
    zCaption = ProperCase(zCaption)
    myLocation = -1
    For butCnt = 0 To TotalButton - 1
        myID = RibbonButtons(butCnt).TopBuID
        If LCase$(myID) = LCase$(zID) Then
            myLocation = butCnt
            Exit For
        End If
    Next
    If myLocation = -1 Then Exit Sub
    If SplitCaption = True Then zCaption = Replace$(zCaption, "", vbNewLine)
    RibbonButtons(myLocation).TopBuC = zCaption
    If Len(zToolTip) = 0 Or zToolTip = Null Then
        If InStr(zCaption, vbNewLine) Then
            zCaption = Replace$(zCaption, vbNewLine, "")
        End If
        RibbonButtons(myLocation).TopBuT = zCaption
    Else
        zToolTip = Replace$(zToolTip, vbNewLine, "")
        RibbonButtons(myLocation).TopBuT = zToolTip
    End If
    RibbonButtons(myLocation).TopBuX = zText
    CatsUpdate
    Err.Clear
End Sub
Private Function ButtonsUpdate(SubCat As String, PosIni As Integer) As Integer
    On Error Resume Next
    ' this has been changed to use non-fixed arrays which are redimensioned from time to time
    ' adrians code assumed that one wanted to add 90 buttons and limited everything to that
    Dim TotalButtonT As Integer
    Dim TopBuIDT() As String
    Dim TopBuST() As String
    Dim TopBuCT() As String
    Dim TopBuIT() As Picture
    Dim TopBuTT() As String
    Dim TopBuGT() As Boolean
    Dim TopTxtT() As String
    Dim TopTxtW() As Long
    Dim TopType() As String
    Dim TopFormat() As String
    Dim TopBuX() As String
    Dim TopMax() As Long
    Dim TopMin() As Long
    Dim TotalSize As Long
    Dim i As Integer
    Dim xt As Integer
    Dim posatu As Long
    Dim ESP As Long
    Dim esp2 As Long
    Dim area As Long
    TotalSize = 0
    TotalButtonT = 0
    ReDim TopBuIDT(TotalButton) As String
    ReDim TopBuST(TotalButton) As String
    ReDim TopBuCT(TotalButton) As String
    ReDim TopBuIT(TotalButton) As Picture
    ReDim TopBuTT(TotalButton) As String
    ReDim TopBuGT(TotalButton) As Boolean
    ReDim TopTxtT(TotalButton) As String
    ReDim TopTxtW(TotalButton) As Long
    ReDim TopType(TotalButton) As String
    ReDim TopFormat(TotalButton) As String
    ReDim TopMin(TotalButton) As Long
    ReDim TopMax(TotalButton) As Long
    ReDim TopBuX(TotalButton) As String
    For i = 0 To TotalButton - 1
        If RibbonButtons(i).TopBuS = SubCat Then
            TopBuIDT(TotalButtonT) = RibbonButtons(i).TopBuID
            TopBuST(TotalButtonT) = RibbonButtons(i).TopBuS
            TopBuCT(TotalButtonT) = RibbonButtons(i).TopBuC
            TopBuTT(TotalButtonT) = RibbonButtons(i).TopBuT
            Set TopBuIT(TotalButtonT) = RibbonButtons(i).TopBuI
            TopBuGT(TotalButtonT) = RibbonButtons(i).TopBuG
            TopTxtT(TotalButtonT) = RibbonButtons(i).TopTxt
            TopTxtW(TotalButtonT) = RibbonButtons(i).TopWdt
            TopType(TotalButtonT) = RibbonButtons(i).TopType
            TopFormat(TotalButtonT) = RibbonButtons(i).TopFormat
            TopMax(TotalButtonT) = RibbonButtons(i).TopMax
            TopMin(TotalButtonT) = RibbonButtons(i).TopMin
            TopBuX(TotalButtonT) = RibbonButtons(i).TopBuX
            TotalButtonT = TotalButtonT + 1
        End If
    Next
    Button_left(0).Visible = False
    Button_center(0).Visible = False
    Button_right(0).Visible = False
    Button_Caption(0).Visible = True
    Button_Icon(0).Visible = True
    Button_Text(0).Visible = True
    ButMouse(0).Visible = True
    txtBox(0).Visible = True
    cboBox(0).Visible = True
    datePick(0).Visible = True
    progBar(0).Visible = True
    xt = ButMouse.UBound + 1
    For i = xt To (TotalButtonT - 1) + xt
        If i <> 0 Then
            Load Button_left(i)
            Load Button_center(i)
            Load Button_right(i)
            Load Button_left_over(i)
            Load Button_center_over(i)
            Load Button_right_over(i)
            Load Button_Caption(i)
            Load Button_Icon(i)
            Load Glip_on(i)
            Load Glip_off(i)
            Load ButMouse(i)
            Load Button_Text(i)
            Load txtBox(i)
            txtBox(i).Text = ""
            Load cboBox(i)
            cboBox(i).Clear
            Load datePick(i)
            Load progBar(i)
        End If
        ButMouse(i).Tag = TopBuIDT(i - xt)
        ButMouse(i).Top = Cat_Left_off(0).Top + 60
        Button_left(i).Top = ButMouse(i).Top
        Button_center(i).Top = ButMouse(i).Top
        Button_right(i).Top = ButMouse(i).Top
        Button_left_over(i).Top = ButMouse(i).Top
        Button_center_over(i).Top = ButMouse(i).Top
        Button_right_over(i).Top = ButMouse(i).Top
        Button_Text(i).Top = ButMouse(i).Top
        Button_Text(i).Tag = TopBuIDT(i - xt)
        If i = xt Then
            posatu = PosIni
        Else
            posatu = ButMouse(i - 1).Left + ButMouse(i - 1).Width + 30
        End If
        ButMouse(i).Left = posatu
        Button_left(i).Left = ButMouse(i).Left
        Button_left_over(i).Left = Button_left(i).Left
        Button_center(i).Left = Button_left(i).Left + Button_left(i).Width
        Button_center_over(i).Left = Button_center(i).Left
        Button_Caption(i).Caption = TopBuCT(i - xt)
        Set Button_Icon(i) = TopBuIT(i - xt)
        Button_Text(i).Caption = vbNewLine & TopBuX(i - xt)
        Button_Text(i).Left = ButMouse(i).Left
        If Len(Button_Caption(i).Caption) > Len(Button_Text(i).Caption) Then Button_Text(i).Width = Button_Caption(i).Width
        If Len(Button_Text(i).Caption) > Len(Button_Caption(i).Caption) Then Button_Caption(i).Width = Button_Text(i).Width
        If ImageSize = Size160 Then
            Button_Icon(i).Width = 160
            Button_Icon(i).Height = 160
        ElseIf ImageSize = Size320 Then
            Button_Icon(i).Width = 320
            Button_Icon(i).Height = 320
        ElseIf ImageSize = Size240 Then
            Button_Icon(i).Width = 240
            Button_Icon(i).Height = 240
        End If
        Button_Text(i).BackStyle = 0
        If Len(TopTxtT(i - xt)) > 0 Then
            If TopType(i - xt) = "t" Then
                txtBox(i - xt).Visible = True
                txtBox(i - xt).Tag = TopTxtT(i - xt)
                txtBox(i - xt).ToolTipText = TopBuTT(i - xt)
                txtBox(i - xt).Width = TopTxtW(i - xt)
            ElseIf TopType(i - xt) = "c" Then
                cboBox(i - xt).Visible = True
                cboBox(i - xt).Tag = TopTxtT(i - xt)
                cboBox(i - xt).ToolTipText = TopBuTT(i - xt)
                cboBox(i - xt).Width = TopTxtW(i - xt)
            ElseIf TopType(i - xt) = "dp" Then
                datePick(i - xt).Visible = True
                datePick(i - xt).Tag = TopTxtT(i - xt)
                datePick(i - xt).ToolTipText = TopBuTT(i - xt)
                datePick(i - xt).Width = TopTxtW(i - xt)
                datePick(i - xt).Format = dtpCustom
                datePick(i - xt).CustomFormat = TopFormat(i - xt)
            ElseIf TopType(i - xt) = "prog" Then
                progBar(i - xt).Visible = True
                progBar(i - xt).Tag = TopTxtT(i - xt)
                progBar(i - xt).ToolTipText = TopBuTT(i - xt)
                progBar(i - xt).Width = TopTxtW(i - xt)
                progBar(i - xt).Max = TopMax(i - xt)
                progBar(i - xt).Min = TopMin(i - xt)
                progBar(i - xt).Value = 0
            End If
            Button_Icon(i).Width = TopTxtW(i - xt)
            Button_Icon(i).Height = 315
        Else
            txtBox(i - xt).Visible = False
            cboBox(i - xt).Visible = False
            datePick(i - xt).Visible = False
            progBar(i - xt).Visible = False
        End If
        Button_Icon(i).Stretch = True
        ESP = Button_center(i).Height - (Button_Icon(i).Height + Button_Caption(i).Height)
        If TopBuGT(i - xt) = True Then
            Button_Icon(i).Top = Button_center(i).Top + ((ESP - (Button_Caption(i).Height / 2)) / 2)
        Else
            Button_Icon(i).Top = Button_center(i).Top + ((ESP) / 2)
        End If
        txtBox(i - xt).Top = Button_Icon(i).Top
        cboBox(i - xt).Top = Button_Icon(i).Top
        datePick(i - xt).Top = Button_Icon(i).Top
        progBar(i - xt).Top = Button_Icon(i).Top
        Button_Caption(i).Top = Button_Icon(i).Top + Button_Icon(i).Height
        Glip_off(i).Top = Button_Caption(i).Top + Button_Caption(i).Height + ((Button_Caption(i).Height - Glip_off(i).Height) / 2)
        Glip_on(i).Top = Glip_off(i).Top
        If Button_Caption(i).Width > Button_Icon(i).Width Then
            Button_Caption(i).Left = Button_center(i).Left
            esp2 = (Button_Caption(i).Width - Button_Icon(i).Width) / 2
            Button_Icon(i).Left = Button_Caption(i).Left + esp2
            area = Button_Caption(i).Width
        Else
            Button_Icon(i).Left = Button_center(i).Left
            esp2 = (Button_Icon(i).Width - Button_Caption(i).Width) / 2
            Button_Caption(i).Left = Button_Icon(i).Left + esp2
            area = Button_Icon(i).Width
        End If
        Button_Text(i).Left = Button_Caption(i).Left
        txtBox(i - xt).Left = ButMouse(i).Left + 50
        cboBox(i - xt).Left = ButMouse(i).Left + 50
        datePick(i - xt).Left = ButMouse(i).Left + 50
        progBar(i - xt).Left = ButMouse(i).Left + 50
        Glip_off(i).Left = Button_Caption(i).Left + ((Button_Caption(i).Width - Glip_on(i).Width) / 2)
        Glip_on(i).Left = Glip_off(i).Left
        Button_center(i).Width = area
        Button_center_over(i).Width = Button_center(i).Width
        Button_right(i).Left = Button_center(i).Left + Button_center(i).Width
        Button_right_over(i).Left = Button_right(i).Left
        ButMouse(i).Width = (Button_right(i).Width + Button_right(i).Width) + Button_center(i).Width
        ButMouse(i).ToolTipText = TopBuTT(i - xt)
        Button_Icon(i).Visible = True
        Button_Caption(i).Visible = True
        ButMouse(i).Visible = True
        Button_Text(i).Visible = True
        If TopBuGT(i - xt) = True Then
            Glip_off(i).Visible = True
            Glip_off(i).ZOrder 0
            Glip_on(i).ZOrder 0
        End If
        TotalSize = TotalSize + ButMouse(i).Width + 30
    Next
    ButtonsUpdate = TotalSize - 30
    Err.Clear
End Function
Private Function ProperCase(ByVal StrString As String, Optional Delim As String = "\") As String
    On Error Resume Next
    'make a string propercase
    Dim spItems() As String
    Dim spSubs() As String
    Dim spTot As Long
    Dim spCnt As Long
    Dim spTott As Long
    Dim spCntt As Long
    Dim spSubss() As String
    StrString = Trim$(StrString)
    spTot = StrParse(spItems, StrString, Delim)
    For spCnt = 1 To spTot
        spItems(spCnt) = StrConv(spItems(spCnt), vbProperCase)
        spTott = StrParse(spSubss, spItems(spCnt), "|")
        For spCntt = 1 To spTott
            spSubss(spCntt) = StrConv(spSubss(spCntt), vbProperCase)
        Next
        spItems(spCnt) = MvFromArray(spSubss, "|")
    Next
    ProperCase = MvFromArray(spItems, Delim)
    Erase spItems
    Erase spSubs
    Err.Clear
End Function
Private Function StrParse(retarray() As String, ByVal strText As String, ByVal Delimiter As String, Optional RedimensionTo As Long = -1) As Long
    On Error Resume Next
    ' the VB split function clone, this starting at 1
    Dim varArray() As String
    Dim varCnt As Long
    Dim VarS As Long
    Dim VarE As Long
    Dim varA As Long
    varArray = Split(strText, Delimiter)
    VarS = LBound(varArray)
    VarE = UBound(varArray)
    varA = VarE + 1
    ReDim retarray(varA)
    For varCnt = VarS To VarE
        varA = varCnt + 1
        retarray(varA) = varArray(varCnt)
    Next
    If RedimensionTo <> -1 Then ReDim Preserve retarray(RedimensionTo)
    StrParse = UBound(retarray)
    Err.Clear
End Function
Private Function MvFromArray(vArray() As String, Optional ByVal Delim As String = "", Optional StartingAt As Long = 1, Optional TrimItem As Boolean = True, Optional Reverse As Boolean = False) As String
    On Error Resume Next
    ' build a delimited string from an array
    If Len(Delim) = 0 Then
        Delim = Chr$(253)
    End If
    Dim i As Long
    Dim BldStr As String
    Dim strL As String
    Dim totArray As Long
    If IsArray(vArray) = False Then Exit Function
    totArray = UBound(vArray)
    For i = StartingAt To totArray
        strL = vArray(i)
        If TrimItem = True Then
            strL = Trim$(strL)
        End If
        If i = totArray Then
            BldStr = BldStr & strL
        Else
            BldStr = BldStr & strL & Delim
        End If
    Next
    If Reverse = True Then
        For i = totArray To StartingAt Step -1
            strL = vArray(i)
            If TrimItem = True Then
                strL = Trim$(strL)
            End If
            If i = totArray Then
                BldStr = BldStr & strL
            Else
                BldStr = BldStr & strL & Delim
            End If
        Next
    End If
    MvFromArray = BldStr
    Err.Clear
End Function
Private Function MvSearch(ByVal strMv As String, ByVal StrSearch As String, Delimiter As String) As Long
    On Error Resume Next
    ' return the position of a delimited string for the searched one
    Dim xValues() As String
    Dim xPos As Long
    xValues = Split(strMv, Delimiter)
    xPos = ArraySearch(xValues, StrSearch)
    MvSearch = IIf((xPos = -1), 0, xPos + 1)
    Err.Clear
End Function
Private Function ArraySearch(varArray() As String, ByVal StrSearch As String) As Long
    On Error Resume Next
    ' return a position of the item in the array
    Dim ArrayTot As Long
    Dim arrayCnt As Long
    Dim strCur As String
    Dim arrayLow As Long
    ArrayTot = UBound(varArray)
    arrayLow = LBound(varArray)
    StrSearch = LCase$(Trim$(StrSearch))
    ArraySearch = -1
    For arrayCnt = arrayLow To ArrayTot
        strCur = LCase$(varArray(arrayCnt))
        Select Case strCur
        Case StrSearch
            ArraySearch = arrayCnt
            Exit For
        End Select
    Next
    Err.Clear
End Function
Private Function MvField(ByVal strData As String, ByVal fldPos As Long, ByVal Delim As String) As String
    On Error Resume Next
    ' returns a substring from a delimted string
    Dim spData() As String
    Dim spCnt As Long
    MvField = ""
    If Len(Delim) = 0 Then
        Delim = Chr$(253)
    End If
    If Len(strData) = 0 Then
        Err.Clear
        Exit Function
    End If
    Call StrParse(spData, strData, Delim)
    spCnt = UBound(spData)
    Select Case fldPos
    Case -1
        MvField = Trim$(spData(spCnt))
    Case -2
        MvField = Trim$(spData(spCnt - 1))
    Case Else
        If fldPos <= spCnt Then
            MvField = Trim$(spData(fldPos))
        End If
    End Select
    Err.Clear
End Function
Public Sub EditTopButton(zID As String, zNewID As String, zCaption As String, zPicture As Variant, zToolTip As String, Optional SplitCaption As Boolean = False)
    On Error Resume Next
    ' edit the top menu button details
    'Freeze
    Dim butCnt As Integer
    Dim myID As String
    Dim myLocation As Integer
    zCaption = ProperCase(zCaption)
    myID = """"
    myLocation = -1
    For butCnt = 0 To TotalTopButton - 1
        myID = TopButtons(butCnt).TopBID
        If LCase$(myID) = LCase$(zID) Then
            myLocation = butCnt
            Exit For
        End If
    Next
    If myLocation = -1 Then Exit Sub
    If SplitCaption = True Then zCaption = Replace$(zCaption, "", vbNewLine)
    TopButtons(myLocation).TopBID = zNewID
    TopButtons(myLocation).TopBC = zCaption
    If Len(zToolTip) = 0 Or zToolTip = Null Then
        If InStr(zCaption, vbNewLine) Then
            zCaption = Replace$(zCaption, vbNewLine, " ")
        End If
        TBMouse(myLocation).ToolTipText = zCaption
    Else
        zToolTip = Replace$(zToolTip, vbNewLine, " ")
        TBMouse(myLocation).ToolTipText = zToolTip
    End If
    Set RibbonTopImage(myLocation).Picture = zImg.ListImages.Item(GetIconIndex(zImg, zPicture)).Picture
    CatsUpdate
    'Freeze False
    Err.Clear
End Sub
Public Sub AddTextBox(zID As String, zSubCat As String, zCaption As String, zToolTip As String, ByVal txtName As String, ByVal txtWidth As Long)
    On Error Resume Next
    ' add a button with a textbox
    If UsePermissions = False Then GoTo AddIt
    Dim strPrefix As String
    Dim strSuffix As String
    strPrefix = MvField(zID, 1, "_")
    Select Case strPrefix
    Case "openportfolio"
        strSuffix = MvField(zID, -1, "-")
        If IsNumeric(strSuffix) = True Then
        Else
            strSuffix = MvField(zID, 3, "_")
            If MvSearch(Permissions, "openportfolio_" & strSuffix, ";") = 0 Then Exit Sub
        End If
    Case Else
        If MvSearch(Permissions, zID, ";") = 0 Then Exit Sub
    End Select
AddIt:
    zCaption = ProperCase(zCaption)
    TotalButton = TotalButton + 1
    ReDim Preserve RibbonButtons(TotalButton - 1)
    RibbonButtons(TotalButton - 1).TopBuID = zID
    RibbonButtons(TotalButton - 1).TopBuS = zSubCat
    RibbonButtons(TotalButton - 1).TopBuC = zCaption
    If Len(zToolTip) = 0 Or zToolTip = Null Then
        If InStr(zCaption, vbNewLine) Then
            zCaption = Replace$(zCaption, vbNewLine, " ")
        End If
        RibbonButtons(TotalButton - 1).TopBuT = zCaption
    Else
        zToolTip = Replace$(zToolTip, vbNewLine, " ")
        RibbonButtons(TotalButton - 1).TopBuT = zToolTip
    End If
    Set RibbonButtons(TotalButton - 1).TopBuI = Nothing
    RibbonButtons(TotalButton - 1).TopBuG = False
    RibbonButtons(TotalButton - 1).TopTxt = txtName
    RibbonButtons(TotalButton - 1).TopWdt = txtWidth
    RibbonButtons(TotalButton - 1).TopType = "t"
    Err.Clear
End Sub
Public Sub AddProgressBar(zID As String, zSubCat As String, zCaption As String, zToolTip As String, ByVal progName As String, progWidth As Long, ByVal minValue As Long, ByVal maxValue As Long)
    On Error Resume Next
    ' add a button with a progress bar
    If UsePermissions = False Then GoTo AddIt
    Dim strPrefix As String
    Dim strSuffix As String
    strPrefix = MvField(zID, 1, "_")
    Select Case strPrefix
    Case "openportfolio"
        strSuffix = MvField(zID, -1, "-")
        If IsNumeric(strSuffix) = True Then
        Else
            strSuffix = MvField(zID, 3, "_")
            If MvSearch(Permissions, "openportfolio_" & strSuffix, ";") = 0 Then Exit Sub
        End If
    Case Else
        If MvSearch(Permissions, zID, ";") = 0 Then Exit Sub
    End Select
AddIt:
    zCaption = ProperCase(zCaption)
    TotalButton = TotalButton + 1
    ReDim Preserve RibbonButtons(TotalButton - 1)
    RibbonButtons(TotalButton - 1).TopBuID = zID
    RibbonButtons(TotalButton - 1).TopBuS = zSubCat
    RibbonButtons(TotalButton - 1).TopBuC = zCaption
    If Len(zToolTip) = 0 Or zToolTip = Null Then
        If InStr(zCaption, vbNewLine) Then
            zCaption = Replace$(zCaption, vbNewLine, " ")
        End If
        RibbonButtons(TotalButton - 1).TopBuT = zCaption
    Else
        zToolTip = Replace$(zToolTip, vbNewLine, " ")
        RibbonButtons(TotalButton - 1).TopBuT = zToolTip
    End If
    Set RibbonButtons(TotalButton - 1).TopBuI = Nothing
    RibbonButtons(TotalButton - 1).TopBuG = False
    RibbonButtons(TotalButton - 1).TopTxt = progName
    RibbonButtons(TotalButton - 1).TopWdt = progWidth
    RibbonButtons(TotalButton - 1).TopType = "prog"
    RibbonButtons(TotalButton - 1).TopMin = minValue
    RibbonButtons(TotalButton - 1).TopMax = maxValue
    Err.Clear
End Sub
Public Sub AddComboBoxItem(ByVal ComboName As String, ByVal ComboItem As String)
    On Error Resume Next
    ' add an item to the combobox
    cboMaster.AddItem ComboName & "|" & ComboItem
    Err.Clear
End Sub
Private Function TextBoxSearch(TextBoxName As String) As Integer
    On Error Resume Next
    ' return the position of the combobox
    Dim rsCnt As Integer
    Dim rsTot As Integer
    Dim cboIdx As Integer
    Dim strTag As String
    cboIdx = -1
    rsTot = txtBox.Count - 1
    For rsCnt = 0 To rsTot
        strTag = txtBox(rsCnt).Tag
        If LCase$(strTag) = LCase$(TextBoxName) Then
            cboIdx = rsCnt
            Exit For
        End If
    Next
    TextBoxSearch = cboIdx
    Err.Clear
End Function
Public Sub TextBoxSetText(TextBoxName As String, TextBoxValue As String)
    On Error Resume Next
    Dim idxPos As Integer
    idxPos = TextBoxSearch(TextBoxName)
    If idxPos >= 0 Then
        txtBox(idxPos).Text = TextBoxValue
    End If
    Err.Clear
End Sub
Private Function ComboBoxSearch(ComboName As String) As Integer
    On Error Resume Next
    ' return the position of the combobox
    Dim rsCnt As Integer
    Dim rsTot As Integer
    Dim cboIdx As Integer
    Dim strTag As String
    cboIdx = -1
    rsTot = cboBox.Count - 1
    For rsCnt = 0 To rsTot
        strTag = cboBox(rsCnt).Tag
        If LCase$(strTag) = LCase$(ComboName) Then
            cboIdx = rsCnt
            Exit For
        End If
    Next
    ComboBoxSearch = cboIdx
    Err.Clear
End Function
Public Sub DatePickerSetDate(ByVal DatePickerName As String, ByVal newDate As String)
    On Error Resume Next
    Dim idxPos As Integer
    idxPos = DatePickerSearch(DatePickerName)
    If idxPos >= 0 Then datePick(idxPos).Value = Format$(newDate, datePick(idxPos).CustomFormat)
    Err.Clear
End Sub
Public Function DatePickerGetDate(ByVal DatePickerName As String) As String
    On Error Resume Next
    Dim idxPos As Integer
    idxPos = DatePickerSearch(DatePickerName)
    If idxPos >= 0 Then
        DatePickerGetDate = Format$(datePick(idxPos).Value, datePick(idxPos).CustomFormat)
    Else
        DatePickerGetDate = ""
    End If
    Err.Clear
End Function
Private Function DatePickerSearch(DatePickerName As String) As Integer
    On Error Resume Next
    ' return the position of the combobox
    Dim rsCnt As Integer
    Dim rsTot As Integer
    Dim cboIdx As Integer
    Dim strTag As String
    cboIdx = -1
    rsTot = datePick.Count - 1
    For rsCnt = 0 To rsTot
        strTag = datePick(rsCnt).Tag
        If LCase$(strTag) = LCase$(DatePickerName) Then
            cboIdx = rsCnt
            Exit For
        End If
    Next
    DatePickerSearch = cboIdx
    Err.Clear
End Function
Public Function ComboBoxGetText(ComboName As String) As String
    On Error Resume Next
    Dim idxPos As Integer
    idxPos = ComboBoxSearch(ComboName)
    If idxPos = -1 Then
        ComboBoxGetText = ""
    Else
        ComboBoxGetText = cboBox(idxPos).Text
    End If
    Err.Clear
End Function
Public Function TextBoxGetText(TextBoxName As String) As String
    On Error Resume Next
    Dim idxPos As Integer
    idxPos = TextBoxSearch(TextBoxName)
    If idxPos = -1 Then
        TextBoxGetText = ""
    Else
        TextBoxGetText = txtBox(idxPos).Text
    End If
    Err.Clear
End Function
Private Function ProgressBarSearch(ProgBarName As String) As Integer
    On Error Resume Next
    ' return the location of the progress bar
    Dim rsCnt As Integer
    Dim rsTot As Integer
    Dim cboIdx As Integer
    Dim strTag As String
    cboIdx = -1
    rsTot = progBar.Count - 1
    For rsCnt = 0 To rsTot
        strTag = progBar(rsCnt).Tag
        If LCase$(strTag) = LCase$(ProgBarName) Then
            cboIdx = rsCnt
            Exit For
        End If
    Next
    ProgressBarSearch = cboIdx
    Err.Clear
End Function
Private Function LabelSearch(LabelName As String) As Integer
    On Error Resume Next
    ' return the position of the label control
    Dim rsCnt As Integer
    Dim rsTot As Integer
    Dim cboIdx As Integer
    Dim strTag As String
    cboIdx = -1
    rsTot = Button_Text.Count - 1
    For rsCnt = 0 To rsTot
        strTag = Button_Text(rsCnt).Tag
        If LCase$(strTag) = LCase$(LabelName) Then
            cboIdx = rsCnt
            Exit For
        End If
    Next
    LabelSearch = cboIdx
    Err.Clear
End Function
Public Sub ProgressBarReset(ProgBarName As String, maxValue As Long, Optional minValue As Long = 0)
    On Error Resume Next
    ' reset the progress bar
    Dim idxPos As Integer
    idxPos = ProgressBarSearch(ProgBarName)
    If idxPos >= 0 Then
        progBar(idxPos).Value = 0
        progBar(idxPos).Max = maxValue
        progBar(idxPos).Min = minValue
    End If
    Err.Clear
End Sub
Public Sub ProgressBarUpdate(ProgBarName As String, curValue As Long)
    On Error Resume Next
    ' update the value of the progressbar button
    Dim idxPos As Integer
    idxPos = ProgressBarSearch(ProgBarName)
    If idxPos >= 0 Then
        If curValue <= progBar(idxPos).Max Then
            progBar(idxPos).Value = curValue
        Else
            progBar(idxPos).Value = progBar(idxPos).Min
        End If
    End If
    Err.Clear
End Sub
Public Sub LabelUpdate(LabelName As String, curValue As String)
    On Error Resume Next
    ' update the caption of the label button
    Dim idxPos As Integer
    idxPos = LabelSearch(LabelName)
    If idxPos >= 0 Then
        Button_Text(idxPos).Caption = vbNewLine & curValue
    End If
    Err.Clear
End Sub
Public Sub ComboBoxRefresh()
    On Error Resume Next
    ' load details of the combo boxes
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim strPrefix As String
    Dim strSuffix As String
    Dim strLine As String
    Dim idxLoc As Integer
    rsTot = cboBox.Count - 1
    For rsCnt = 0 To rsTot
        cboBox(rsCnt).Clear
    Next
    rsTot = cboMaster.ListCount - 1
    For rsCnt = 0 To rsTot
        strLine = cboMaster.List(rsCnt)
        strPrefix = MvField(strLine, 1, "|")
        strSuffix = MvField(strLine, 2, "|")
        idxLoc = ComboBoxSearch(strPrefix)
        If idxLoc >= 0 Then cboBox(idxLoc).AddItem strSuffix
    Next
    Err.Clear
End Sub
Public Sub AddDatePicker(zID As String, zSubCat As String, zCaption As String, zToolTip As String, ByVal DatePickerName As String, Optional ByVal DatePickerWidth As Long = 1355, Optional ByVal DatePickerFormat As String = "yyyy-mm-dd")
    On Error Resume Next
    ' add a button with a date picker control
    If UsePermissions = False Then GoTo AddIt
    Dim strPrefix As String
    Dim strSuffix As String
    strPrefix = MvField(zID, 1, "_")
    Select Case strPrefix
    Case "openportfolio"
        strSuffix = MvField(zID, -1, "-")
        If IsNumeric(strSuffix) = True Then
        Else
            strSuffix = MvField(zID, 3, "_")
            If MvSearch(Permissions, "openportfolio_" & strSuffix, ";") = 0 Then Exit Sub
        End If
    Case Else
        If MvSearch(Permissions, zID, ";") = 0 Then Exit Sub
    End Select
AddIt:
    zCaption = ProperCase(zCaption)
    TotalButton = TotalButton + 1
    ReDim Preserve RibbonButtons(TotalButton - 1)
    RibbonButtons(TotalButton - 1).TopBuID = zID
    RibbonButtons(TotalButton - 1).TopBuS = zSubCat
    RibbonButtons(TotalButton - 1).TopBuC = zCaption
    If Len(zToolTip) = 0 Or zToolTip = Null Then
        If InStr(zCaption, vbNewLine) Then
            zCaption = Replace$(zCaption, vbNewLine, " ")
        End If
        RibbonButtons(TotalButton - 1).TopBuT = zCaption
    Else
        zToolTip = Replace$(zToolTip, vbNewLine, " ")
        RibbonButtons(TotalButton - 1).TopBuT = zToolTip
    End If
    Set RibbonButtons(TotalButton - 1).TopBuI = Nothing
    RibbonButtons(TotalButton - 1).TopBuG = False
    RibbonButtons(TotalButton - 1).TopTxt = DatePickerName
    RibbonButtons(TotalButton - 1).TopWdt = DatePickerWidth
    RibbonButtons(TotalButton - 1).TopType = "dp"
    Err.Clear
End Sub
Public Sub ComboBoxClear(ByVal ComboName As String)
    On Error Resume Next
    ' clear the combobox
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim strPrefix As String
    Dim strSuffix As String
    Dim strLine As String
    Dim idxLoc As Integer
    rsTot = cboMaster.ListCount - 1
    For rsCnt = rsTot To 0 Step -1
        strLine = cboMaster.List(rsCnt)
        strPrefix = MvField(strLine, 1, "|")
        strSuffix = MvField(strLine, 2, "|")
        If LCase$(strPrefix) = LCase$(ComboName) Then cboMaster.RemoveItem rsCnt
    Next
    ComboBoxRefresh
    Err.Clear
End Sub
Public Sub AddLabel(zID As String, zSubCat As String, zCaption As String, zMore As Boolean, zToolTip As String, SplitCaption As Boolean, ByVal zLabel As String)
    On Error Resume Next
    ' add a button with a label
    If UsePermissions = False Then GoTo AddIt
    Dim strPrefix As String
    Dim strSuffix As String
    strPrefix = MvField(zID, 1, "_")
    Select Case strPrefix
    Case "openportfolio"
        strSuffix = MvField(zID, -1, "-")
        If IsNumeric(strSuffix) = True Then
        Else
            strSuffix = MvField(zID, 3, "_")
            If MvSearch(Permissions, "openportfolio_" & strSuffix, ";") = 0 Then Exit Sub
        End If
    Case Else
        If MvSearch(Permissions, zID, ";") = 0 Then Exit Sub
    End Select
AddIt:
    zCaption = ProperCase(zCaption)
    If SplitCaption = True Then zCaption = Replace$(zCaption, " ", vbNewLine)
    TotalButton = TotalButton + 1
    ReDim Preserve RibbonButtons(TotalButton - 1)
    RibbonButtons(TotalButton - 1).TopBuID = zID
    RibbonButtons(TotalButton - 1).TopBuS = zSubCat
    RibbonButtons(TotalButton - 1).TopBuC = zCaption
    If Len(zToolTip) = 0 Or zToolTip = Null Then
        If InStr(zCaption, vbNewLine) Then
            zCaption = Replace$(zCaption, vbNewLine, " ")
        End If
        RibbonButtons(TotalButton - 1).TopBuT = zCaption
    Else
        zToolTip = Replace$(zToolTip, vbNewLine, " ")
        RibbonButtons(TotalButton - 1).TopBuT = zToolTip
    End If
    Set RibbonButtons(TotalButton - 1).TopBuI = Nothing
    RibbonButtons(TotalButton - 1).TopBuG = zMore
    RibbonButtons(TotalButton - 1).TopTxt = ""
    RibbonButtons(TotalButton - 1).TopWdt = 0
    RibbonButtons(TotalButton - 1).TopType = ""
    RibbonButtons(TotalButton - 1).TopFormat = ""
    RibbonButtons(TotalButton - 1).TopBuX = zLabel
    Err.Clear
End Sub
Private Function ButtonSearch(ByVal zID As String) As Integer
    On Error Resume Next
    ' return the location of a button
    Dim butCnt As Integer
    Dim myID As String
    ButtonSearch = -1
    For butCnt = 0 To TotalButton - 1
        myID = RibbonButtons(butCnt).TopBuID
        If LCase$(myID) = LCase$(zID) Then
            ButtonSearch = butCnt
            Exit For
        End If
    Next
    Err.Clear
End Function
Private Function MvFromMv(ByVal strOriginalMv As String, ByVal startPos As Long, Optional ByVal NumOfItems As Long = -1, Optional ByVal Delim As String = vbNullString) As String
    On Error Resume Next
    ' extract a multi value string from another multi value string
    Dim sporiginal() As String
    Dim spTot As Long
    Dim spCnt As Long
    Dim sLine As String
    Dim endPos As Long
    sLine = vbNullString
    If Len(Delim) = 0 Then
        Delim = Chr$(253)
    End If
    StrParse sporiginal, strOriginalMv, Delim
    spTot = UBound(sporiginal)
    If NumOfItems = -1 Then
        endPos = spTot
    ElseIf NumOfItems = -2 Then
        endPos = spTot - 1
    Else
        endPos = (startPos + NumOfItems) - 1
    End If
    For spCnt = startPos To endPos
        If spCnt = endPos Then
            sLine = sLine & sporiginal(spCnt)
        Else
            sLine = sLine & sporiginal(spCnt) & Delim
        End If
    Next
    MvFromMv = sLine
    Err.Clear
End Function
Private Function RemDelim(ByVal Dataobj As String, ByVal Delimiter As String) As String
    On Error Resume Next
    ' remove a delimiter from a string
    Dim intDataSize As Long
    Dim intDelimSize As Long
    Dim strLast As String
    intDataSize = Len(Dataobj)
    intDelimSize = Len(Delimiter)
    strLast = Right$(Dataobj, intDelimSize)
    Select Case strLast
    Case Delimiter
        RemDelim = Left$(Dataobj, (intDataSize - intDelimSize))
    Case Else
        RemDelim = Dataobj
    End Select
    Err.Clear
End Function
Private Function MvCount(ByVal StringMv As String, Optional ByVal Delim As String = vbNullString) As Long
    On Error Resume Next
    ' count the number of delimited strings
    Dim xNew() As String
    If Len(Delim) = 0 Then
        Delim = Chr$(253)
    End If
    If Len(StringMv) = 0 Then
        MvCount = 0
        Err.Clear
        Exit Function
    End If
    xNew = Split(StringMv, Delim)
    MvCount = UBound(xNew) + 1
    Err.Clear
End Function
Private Function MenuSimilar(ByVal strMainMenu As String) As String
    On Error Resume Next
    ' this is used to group menus of a similar nature
    ' this does not perform a sort
    Dim menuTot As Long
    Dim menuCnt As Long
    Dim strLine As String
    Dim menu1 As String
    Dim strMenus As String
    Dim menuID As String
    strMenus = ""
    menuTot = cboMenus.ListCount - 1
    For menuCnt = 0 To menuTot
        strLine = cboMenus.List(menuCnt)
        menuID = MvField(strLine, 2, "|")
        menuID = RemDelim(menuID, "\")
        menu1 = MvField(menuID, 1, "\")
        If LCase$(menu1) = LCase$(strMainMenu) Then
            strMenus = strMenus & strLine & ";"
        End If
    Next
    MenuSimilar = RemDelim(strMenus, ";")
    Err.Clear
End Function
