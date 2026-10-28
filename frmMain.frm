VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "VB6 menu icons - open both menus to compare"
   ClientHeight    =   2160
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6300
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   6300
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Menu mnuClassic 
      Caption         =   "&1 - SetMenuItemBitmaps"
      Begin VB.Menu mnuClassicNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuClassicOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuClassicSave 
         Caption         =   "&Save"
      End
   End
   Begin VB.Menu mnuModern 
      Caption         =   "&2 - SetMenuItemInfo"
      Begin VB.Menu mnuModernNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuModernOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuModernSave 
         Caption         =   "&Save"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' This is a normal VB6 menu, created in Menu Editor.
' No owner-draw, subclassing or external OCX.
' Each menu contains the same three commands, in the same order.

Private Declare Function GetMenu Lib "user32" ( _
    ByVal hwnd As Long) As Long

Private Declare Function GetSubMenu Lib "user32" ( _
    ByVal hMenu As Long, ByVal nPos As Long) As Long

Private Declare Function SetMenuItemBitmaps Lib "user32" ( _
    ByVal hMenu As Long, ByVal nPosition As Long, _
    ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, _
    ByVal hBitmapChecked As Long) As Long

Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" ( _
    ByVal hMenu As Long, ByVal uItem As Long, _
    ByVal fByPosition As Long, ByRef lpmii As MENUITEMINFO) As Long

Private Declare Function DrawMenuBar Lib "user32" ( _
    ByVal hwnd As Long) As Long

Private Const MF_BYPOSITION As Long = &H400&
Private Const MIIM_BITMAP As Long = &H80&

' Win32 MENUITEMINFOA: 12 x 4-byte fields = 48 bytes in VB6.
' dwTypeData is a pointer, not a VB String.
Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As Long
    cch As Long
    hbmpItem As Long
End Type

' Keep pictures alive while Windows uses their HBITMAP handles.
Private mClassicIcons(0 To 2) As Object
Private mModernIcons(0 To 2) As Object

Private Sub Form_Load()
    On Error GoTo LoadError

    Dim hBar As Long
    Dim hClassic As Long
    Dim hModern As Long
    Dim i As Long
    Dim itemInfo As MENUITEMINFO

    ' Legacy version: monochrome images (white is transparent).
    Set mClassicIcons(0) = LoadPicture(App.Path & "\icons\mono_new.bmp")
    Set mClassicIcons(1) = LoadPicture(App.Path & "\icons\mono_open.bmp")
    Set mClassicIcons(2) = LoadPicture(App.Path & "\icons\mono_save.bmp")

    ' MIIM_BITMAP version: color images (white background).
    Set mModernIcons(0) = LoadPicture(App.Path & "\icons\color_new.bmp")
    Set mModernIcons(1) = LoadPicture(App.Path & "\icons\color_open.bmp")
    Set mModernIcons(2) = LoadPicture(App.Path & "\icons\color_save.bmp")

    hBar = GetMenu(Me.hwnd)
    If hBar = 0 Then Err.Raise vbObjectError + 101, , "GetMenu failed"

    ' The two top-level VB6 menus are at positions 0 and 1.
    hClassic = GetSubMenu(hBar, 0)
    hModern = GetSubMenu(hBar, 1)
    If hClassic = 0 Or hModern = 0 Then
        Err.Raise vbObjectError + 102, , "GetSubMenu failed"
    End If

    For i = 0 To 2
        ' Method 1: put a bitmap in the checkmark area.
        ' The last two parameters mean UNCHECKED and CHECKED,
        ' not normal and hovered. We use the same image for both.
        If SetMenuItemBitmaps(hClassic, i, MF_BYPOSITION, _
                mClassicIcons(i).Handle, mClassicIcons(i).Handle) = 0 Then
            Err.Raise vbObjectError + 103, , "SetMenuItemBitmaps failed, item " & CStr(i)
        End If

        ' Method 2: assign the main item bitmap via MENUITEMINFO.
        ' Only MIIM_BITMAP is requested: the caption remains intact.
        itemInfo.cbSize = LenB(itemInfo)
        itemInfo.fMask = MIIM_BITMAP
        itemInfo.hbmpItem = mModernIcons(i).Handle

        If SetMenuItemInfo(hModern, i, 1&, itemInfo) = 0 Then
            Err.Raise vbObjectError + 104, , "SetMenuItemInfo failed, item " & CStr(i)
        End If
    Next i

    Call DrawMenuBar(Me.hwnd)
    Exit Sub

LoadError:
    MsgBox "Could not initialize menu icons:" & vbCrLf & Err.Description, _
           vbCritical, "Menu icon demo"
    Unload Me
End Sub

Private Sub mnuClassicNew_Click()
    ShowSelection "SetMenuItemBitmaps", "New"
End Sub

Private Sub mnuClassicOpen_Click()
    ShowSelection "SetMenuItemBitmaps", "Open"
End Sub

Private Sub mnuClassicSave_Click()
    ShowSelection "SetMenuItemBitmaps", "Save"
End Sub

Private Sub mnuModernNew_Click()
    ShowSelection "SetMenuItemInfo", "New"
End Sub

Private Sub mnuModernOpen_Click()
    ShowSelection "SetMenuItemInfo", "Open"
End Sub

Private Sub mnuModernSave_Click()
    ShowSelection "SetMenuItemInfo", "Save"
End Sub

Private Sub ShowSelection(ByVal methodName As String, ByVal actionName As String)
    MsgBox methodName & " -> " & actionName, vbInformation, "Clicked"
End Sub
