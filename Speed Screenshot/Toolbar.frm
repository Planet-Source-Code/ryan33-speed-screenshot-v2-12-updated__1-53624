VERSION 5.00
Begin VB.Form frmToolbar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Speed Screenshot Toolbar"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5865
   Icon            =   "Toolbar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Toolbar.frx":2E7A
   ScaleHeight     =   1065
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image imgHelp 
      Height          =   810
      Left            =   4920
      MouseIcon       =   "Toolbar.frx":237E4
      MousePointer    =   99  'Custom
      Picture         =   "Toolbar.frx":23936
      ToolTipText     =   "Help"
      Top             =   120
      Width           =   825
   End
   Begin VB.Image imgPrefrences 
      Enabled         =   0   'False
      Height          =   825
      Left            =   3960
      MouseIcon       =   "Toolbar.frx":25CE8
      MousePointer    =   99  'Custom
      Picture         =   "Toolbar.frx":25E3A
      ToolTipText     =   "Prefrences"
      Top             =   120
      Width           =   885
   End
   Begin VB.Image imgResize 
      Enabled         =   0   'False
      Height          =   825
      Left            =   3000
      MouseIcon       =   "Toolbar.frx":28528
      MousePointer    =   99  'Custom
      Picture         =   "Toolbar.frx":2867A
      ToolTipText     =   "Resize"
      Top             =   120
      Width           =   870
   End
   Begin VB.Image imgRectangle 
      Enabled         =   0   'False
      Height          =   825
      Left            =   2040
      MouseIcon       =   "Toolbar.frx":2AC8C
      MousePointer    =   99  'Custom
      Picture         =   "Toolbar.frx":2ADDE
      ToolTipText     =   "Rectangle"
      Top             =   120
      Width           =   870
   End
   Begin VB.Image imgWindow 
      Height          =   825
      Left            =   1080
      MouseIcon       =   "Toolbar.frx":2D3F0
      MousePointer    =   99  'Custom
      Picture         =   "Toolbar.frx":2D542
      ToolTipText     =   "Window"
      Top             =   120
      Width           =   870
   End
   Begin VB.Image imgFullscreen 
      Height          =   780
      Left            =   120
      MouseIcon       =   "Toolbar.frx":2FB54
      MousePointer    =   99  'Custom
      Picture         =   "Toolbar.frx":2FCA6
      ToolTipText     =   "Full Screen"
      Top             =   120
      Width           =   870
   End
End
Attribute VB_Name = "frmToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Const VK_MENU = &H12
Private Const VK_SNAPSHOT = &H2C
Private Const KEYEVENTF_KEYUP = &H2

Private Sub CopyToClipboard(ByVal form_only As Boolean)

  Dim alt_scan_code As Long

    If form_only Then
        alt_scan_code = MapVirtualKey(VK_MENU, 0)
        keybd_event VK_MENU, alt_scan_code, 0, 0
        DoEvents
    End If
    keybd_event VK_SNAPSHOT, 0, 0, 0
    DoEvents
    If form_only Then
        keybd_event VK_MENU, alt_scan_code, KEYEVENTF_KEYUP, 0
        DoEvents
    End If

End Sub

Private Sub imgFullscreen_Click()
    ' Auto hide main form if user has auto hide checked
    ' Checked by default.
    If frmMain.mnuQuickOptionsAutoHide.Checked = True Then
       frmToolbar.Hide
       Me.Hide
    End If

    ' Decide if the picure will displayed in a picture box
    ' or a image box on the main form.
    If frmMain.picSS.Visible = True Then
       frmMain.picSS.Enabled = True
       CopyToClipboard (frmMain.chkWholeScreen.Value = vbUnchecked)
       frmMain.picSS.Picture = Clipboard.GetData
       frmMain.imgSS.Picture = Clipboard.GetData
         ' Save file automaticly if Auto-Save is checked
         If frmMain.mnuQuickOptionsAutoSave.Checked = True Then
            SavePicture frmMain.picSS.Picture, frmMain.txtSaveFiles.Text & "\" & frmMain.txtFileName.Text + frmMain.txtSSNum.Text + ".bmp"
         End If
       frmMain.txtSSNum.Text = frmMain.txtSSNum.Text + 1
       frmMain.cmdFullPreview.Enabled = True
    Else
       frmMain.imgSS.Visible = True
       frmMain.imgSS.Enabled = True
       CopyToClipboard (frmMain.chkWholeScreen.Value = vbUnchecked)
       frmMain.imgSS.Picture = Clipboard.GetData
       frmMain.picSS.Picture = Clipboard.GetData
         ' Save file automaticly if Auto-Save is checked
         If frmMain.mnuQuickOptionsAutoSave.Checked = True Then
            SavePicture frmMain.imgSS.Picture, frmMain.txtSaveFiles.Text & "\" & frmMain.txtFileName.Text + frmMain.txtSSNum.Text + ".bmp"
         End If
       frmMain.txtSSNum.Text = frmMain.txtSSNum.Text + 1
       frmMain.cmdFullPreview.Enabled = True
    End If
    
    ' After taking a screenshot reshow the toolbar
    frmToolbar.Show
    
End Sub

Private Sub imgHelp_Click()
    frmAbout.Show
End Sub

Private Sub imgPrefrences_Click()
    'frmPrefrences.Show
    MsgBox "Sorry this feature not added yet.", , "Stokes Speed Screenshot"
End Sub

Private Sub imgRectangle_Click()
    MsgBox "Sorry this feature not added yet.", , "Stokes Speed Screenshot"
End Sub

Private Sub imgResize_Click()
    MsgBox "Sorry this feature not added yet.", , "Stokes Speed Screenshot"
End Sub

Private Sub imgWindow_Click()
    frmMain.chkWholeScreen.Value = vbUnchecked
    frmMain.mnuQuickOptionsCapture.Checked = False
    Call imgFullscreen_Click
    frmToolbar.Show

End Sub
