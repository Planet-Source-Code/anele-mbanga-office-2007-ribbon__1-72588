VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_hMenu     As Long
Public Caption As String        ' If works as a sub-menu, this would be add as parent menu item
Private Declare Function SetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPos As Long) As Long
Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, lpNewItem As String) As Long
Private Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Private Const MF_STRING = &H0&
Private Const MF_SEPARATOR = &H800&
Private Const MF_MENUBARBREAK = &H20&
Private Const MF_POPUP = &H10&
Private Const MF_BYCOMMAND = &H0&
Private Const MF_GRAYED = &H1&
Private Const MF_DISABLED = &H2&
Private Const MF_CHECKED = &H8&
Private Const APIFALSE As Long = 0
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTL) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, ByVal lprc As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Const GW_CHILD = 5
Private Const GW_HWNDNEXT = 2
Private Type POINTL
    x As Long
    y As Long
End Type
Private Const TPM_RETURNCMD = &H100&
Private Const TPM_LEFTALIGN = &H0&
Private Const TPM_LEFTBUTTON = &H0&
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Function MenuCount() As Long
    On Error Resume Next
    MenuCount = menuCnt
    Err.Clear
End Function
Private Sub Class_Initialize()
    On Error Resume Next
    m_hMenu = CreatePopupMenu()
    Err.Clear
End Sub
Private Sub Class_Terminate()
    On Error Resume Next
    DestroyMenu m_hMenu
    Err.Clear
End Sub
Friend Property Get hMenu() As Long
    On Error Resume Next
    hMenu = m_hMenu
    Err.Clear
End Property
Public Sub Reset()
    On Error Resume Next
    menuCnt = 0
    ReDim menuKeys(menuCnt)
    Err.Clear
End Sub
Public Function MenuSearch(ByVal strMenuID As String) As Long
    On Error Resume Next
    Dim rsCnt As Long
    Dim rsTot As Long
    MenuSearch = -1
    rsTot = UBound(menuKeys)
    For rsCnt = 0 To rsTot
        If LCase$(menuKeys(rsCnt)) = LCase$(strMenuID) Then
            MenuSearch = rsCnt
            Exit For
        End If
    Next
    Err.Clear
End Function
Public Sub AddItem(ByVal nID As Long, varItem As Variant, Optional bDefault As Boolean = False, Optional bChecked As Boolean = False, Optional bDisabled As Boolean = False, Optional bGrayed As Boolean = False, Optional bNewColumn As Boolean = False)
    On Error Resume Next
    If TypeName(varItem) = "String" Then
        If varItem = "-" Then
            AppendMenu m_hMenu, MF_STRING Or MF_SEPARATOR, nID, ByVal vbNullString
        Else
            AppendMenu m_hMenu, MF_STRING Or IIf(bNewColumn, MF_MENUBARBREAK, 0) Or IIf(bChecked, MF_CHECKED, 0), nID, ByVal varItem
        End If
    ElseIf TypeOf varItem Is clsMenu Then
        Dim cSubMenu As clsMenu
        Set cSubMenu = varItem
        AppendMenu m_hMenu, MF_STRING Or MF_POPUP Or IIf(bNewColumn, MF_MENUBARBREAK, 0), cSubMenu.hMenu, ByVal cSubMenu.Caption
    End If
    If bDefault Then SetMenuDefaultItem m_hMenu, nID, APIFALSE
    If bGrayed Then EnableMenuItem m_hMenu, nID, MF_BYCOMMAND Or MF_GRAYED
    If bDisabled Then EnableMenuItem m_hMenu, nID, MF_BYCOMMAND Or MF_DISABLED
    Err.Clear
End Sub
Public Sub AddMenu(ByVal menuID As String, varItem As Variant, Optional bDefault As Boolean = False, Optional bChecked As Boolean = False, Optional bDisabled As Boolean = False, Optional bGrayed As Boolean = False, Optional bNewColumn As Boolean = False)
    On Error Resume Next
    Dim nID As Long
    menuCnt = menuCnt + 1
    nID = menuCnt
    ReDim Preserve menuKeys(nID)
    menuKeys(nID) = menuID
    If TypeName(varItem) = "String" Then
        If varItem = "-" Then
            AppendMenu m_hMenu, MF_STRING Or MF_SEPARATOR, nID, ByVal vbNullString
        Else
            AppendMenu m_hMenu, MF_STRING Or IIf(bNewColumn, MF_MENUBARBREAK, 0) Or IIf(bChecked, MF_CHECKED, 0), nID, ByVal varItem
        End If
    ElseIf TypeOf varItem Is clsMenu Then
        Dim cSubMenu As clsMenu
        Set cSubMenu = varItem
        AppendMenu m_hMenu, MF_STRING Or MF_POPUP Or IIf(bNewColumn, MF_MENUBARBREAK, 0), cSubMenu.hMenu, ByVal cSubMenu.Caption
    End If
    If bDefault Then SetMenuDefaultItem m_hMenu, nID, APIFALSE
    If bGrayed Then EnableMenuItem m_hMenu, nID, MF_BYCOMMAND Or MF_GRAYED
    If bDisabled Then EnableMenuItem m_hMenu, nID, MF_BYCOMMAND Or MF_DISABLED
    Err.Clear
End Sub
Public Function TrackPopup(Optional ByVal hwnd As Long = -1, Optional ByVal PopX As Long = -1, Optional ByVal PopY As Long = -1, Optional ByVal hWndOfBeneathControl As Long = -1) As Long
    On Error Resume Next
    Dim H As Long
    Dim x As Long
    Dim y As Long
    If hwnd = -1 Or hwnd = 0 Then
        ' Find the top window of current process
        Dim hDesktop As Long: hDesktop = GetDesktopWindow()
        Dim hChild As Long: hChild = GetWindow(hDesktop, GW_CHILD)
        Dim idCurrent As Long: idCurrent = GetCurrentProcessId()
        Do While hChild
            Dim idChild As Long
            GetWindowThreadProcessId hChild, idChild
            If idChild = idCurrent Then Exit Do
            hChild = GetWindow(hChild, GW_HWNDNEXT)
        Loop
        If hChild = 0 Then Err.Raise -1, "cMenu.TrackPopup", "Cannot find top window of current process!"
        H = hChild
    Else
        H = hwnd
    End If
    If hWndOfBeneathControl <> -1 Then
        Dim rt As RECT
        GetWindowRect hWndOfBeneathControl, rt
        x = rt.Left
        y = rt.Bottom
    Else
        Dim pt As POINTL
        GetCursorPos pt
        If PopX = -1 Then x = pt.x Else: x = PopX
        If PopY = -1 Then y = pt.y Else: y = PopY
    End If
    TrackPopup = TrackPopupMenu(m_hMenu, TPM_RETURNCMD + TPM_LEFTALIGN + TPM_LEFTBUTTON, x, y, 0, H, 0)
    Err.Clear
End Function
Public Function TrackMenu(Optional ByVal hwnd As Long = -1, Optional ByVal PopX As Long = -1, Optional ByVal PopY As Long = -1, Optional ByVal hWndOfBeneathControl As Long = -1) As String
    On Error Resume Next
    Dim H As Long
    Dim x As Long
    Dim y As Long
    Dim nID As Long
    If hwnd = -1 Or hwnd = 0 Then
        ' Find the top window of current process
        Dim hDesktop As Long: hDesktop = GetDesktopWindow()
        Dim hChild As Long: hChild = GetWindow(hDesktop, GW_CHILD)
        Dim idCurrent As Long: idCurrent = GetCurrentProcessId()
        Do While hChild
            Dim idChild As Long
            GetWindowThreadProcessId hChild, idChild
            If idChild = idCurrent Then Exit Do
            hChild = GetWindow(hChild, GW_HWNDNEXT)
        Loop
        If hChild = 0 Then Err.Raise -1, "cMenu.TrackPopup", "Cannot find top window of current process!"
        H = hChild
    Else
        H = hwnd
    End If
    If hWndOfBeneathControl <> -1 Then
        Dim rt As RECT
        GetWindowRect hWndOfBeneathControl, rt
        x = rt.Left
        y = rt.Bottom
    Else
        Dim pt As POINTL
        GetCursorPos pt
        If PopX = -1 Then x = pt.x Else: x = PopX
        If PopY = -1 Then y = pt.y Else: y = PopY
    End If
    nID = TrackPopupMenu(m_hMenu, TPM_RETURNCMD + TPM_LEFTALIGN + TPM_LEFTBUTTON, x, y, 0, H, 0)
    TrackMenu = menuKeys(nID)
    Err.Clear
End Function
Public Sub AddMenus(parentMenu As Variant, ByVal strMv As String, Optional ByVal strPrefix As String = "", Optional ByVal strDelim As String = ";")
    On Error Resume Next
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim spMenu() As String
    Dim strMenu As String
    spMenu = Split(strMv, strDelim)
    rsTot = UBound(spMenu)
    For rsCnt = 0 To rsTot
        strMenu = strPrefix & spMenu(rsCnt)
        parentMenu.AddMenu strMenu, Replace$(spMenu(rsCnt), "&", "&&")
    Next
    Err.Clear
End Sub
