VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stokes Speed Screenshot"
   ClientHeight    =   7575
   ClientLeft      =   1995
   ClientTop       =   660
   ClientWidth     =   7815
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   7815
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrShowHideTBnQO 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6600
      Top             =   720
   End
   Begin VB.Timer tmrHideMinimized 
      Left            =   7080
      Top             =   720
   End
   Begin VB.Timer tmrHKResetDefaults 
      Interval        =   1
      Left            =   2160
      Top             =   720
   End
   Begin VB.Timer tmrHKFullPreview 
      Interval        =   1
      Left            =   2160
      Top             =   240
   End
   Begin VB.Timer tmrHKSave 
      Interval        =   1
      Left            =   1680
      Top             =   240
   End
   Begin VB.Frame fraMainMenu 
      BackColor       =   &H00E0E0E0&
      Height          =   7455
      Left            =   120
      TabIndex        =   21
      Top             =   0
      Width           =   855
      Begin VB.CommandButton cmdClearMemory 
         BackColor       =   &H00E0E0E0&
         Height          =   615
         Left            =   120
         Picture         =   "Main.frx":2E7A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Clear Previous Capture from Memory"
         Top             =   3840
         Width           =   615
      End
      Begin VB.CommandButton cmdViewPrevSShot 
         BackColor       =   &H00E0E0E0&
         Height          =   615
         Left            =   120
         Picture         =   "Main.frx":3184
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "View Previous Capture"
         Top             =   3120
         Width           =   615
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00E0E0E0&
         Height          =   615
         Left            =   120
         Picture         =   "Main.frx":5926
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Clear Viewer"
         Top             =   4560
         Width           =   615
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00E0E0E0&
         Height          =   615
         Left            =   120
         Picture         =   "Main.frx":61F0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Save (Ctrl+S)"
         Top             =   2400
         Width           =   615
      End
      Begin VB.CommandButton cmdToolbar 
         BackColor       =   &H00E0E0E0&
         Height          =   615
         Left            =   120
         Picture         =   "Main.frx":7EEA
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Tool Bar (Ctrl+T)"
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton cmdGetScreen 
         BackColor       =   &H00E0E0E0&
         Default         =   -1  'True
         Height          =   615
         Left            =   120
         Picture         =   "Main.frx":A2BC
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Get Screen! (Ctrl+G)"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdHide 
         BackColor       =   &H00E0E0E0&
         Height          =   615
         Left            =   120
         Picture         =   "Main.frx":BFB6
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Hide Me (Ctrl+H)"
         Top             =   5280
         Width           =   615
      End
      Begin VB.CommandButton cmdFullPreview 
         BackColor       =   &H00E0E0E0&
         Height          =   615
         Left            =   120
         Picture         =   "Main.frx":DCB0
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Full Preview (Ctrl+P)"
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton cmdReset 
         BackColor       =   &H00E0E0E0&
         Height          =   615
         Left            =   120
         Picture         =   "Main.frx":10082
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Reset Defaults (Ctrl+D)"
         Top             =   6000
         Width           =   615
      End
      Begin VB.CommandButton cmdPrefs 
         BackColor       =   &H00E0E0E0&
         Height          =   615
         Left            =   120
         Picture         =   "Main.frx":13C34
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Preferences"
         Top             =   6720
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Frame fraQO 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Quick-Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   1935
      Left            =   1080
      TabIndex        =   17
      Top             =   5520
      Width           =   6615
      Begin VB.OptionButton opnAllon 
         BackColor       =   &H00E0E0E0&
         Caption         =   "All On"
         Height          =   255
         Left            =   5640
         TabIndex        =   24
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton opnAlloff 
         BackColor       =   &H00E0E0E0&
         Caption         =   "All Off"
         Height          =   255
         Left            =   5640
         TabIndex        =   23
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton cmdBrowse 
         BackColor       =   &H00C0C0C0&
         Caption         =   "..."
         Height          =   255
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Browse"
         Top             =   1440
         Width           =   375
      End
      Begin VB.CheckBox chkAutoSave 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Auto Save"
         Height          =   255
         Left            =   3240
         TabIndex        =   14
         Top             =   600
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkAutoHide 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Auto Hide"
         Height          =   255
         Left            =   3240
         TabIndex        =   15
         Top             =   840
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkWholeScreen 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Capture Whole Screen"
         Height          =   255
         Left            =   3240
         TabIndex        =   13
         Top             =   360
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.TextBox txtSSNum 
         Height          =   285
         Left            =   1800
         TabIndex        =   10
         Text            =   "0"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtFileName 
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Text            =   "screenshot"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtSaveFiles 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Top             =   1440
         Width           =   4695
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         X1              =   3000
         X2              =   3000
         Y1              =   120
         Y2              =   1200
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   6600
         X2              =   3000
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label lblSSPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Picture Number:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "File Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblSaveFiles 
         BackStyle       =   0  'Transparent
         Caption         =   "Save Files:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1440
         Width           =   975
      End
   End
   Begin VB.Timer tmrCheckToAutoSave 
      Interval        =   1
      Left            =   6120
      Top             =   240
   End
   Begin VB.Timer tmrDelayScreenshot 
      Left            =   6600
      Top             =   240
   End
   Begin VB.Timer tmrAutoHide 
      Interval        =   1
      Left            =   7080
      Top             =   240
   End
   Begin VB.Timer tmrHKToolbar 
      Interval        =   1
      Left            =   1680
      Top             =   720
   End
   Begin VB.Timer tmrHKHideme 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1200
      Top             =   720
   End
   Begin VB.Timer tmrHKScreenshot 
      Interval        =   1
      Left            =   1200
      Top             =   240
   End
   Begin Speed_Screenshot.TrayControl TrayControl1 
      Left            =   7080
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      ToolTipText     =   "Speed Screenshot"
   End
   Begin VB.PictureBox picSS 
      BackColor       =   &H00808080&
      Enabled         =   0   'False
      FillStyle       =   0  'Solid
      Height          =   5295
      Left            =   1080
      ScaleHeight     =   5235
      ScaleWidth      =   6555
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Zoom View"
      Top             =   120
      Width           =   6615
   End
   Begin VB.Image imgSS 
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   5295
      Left            =   1080
      MouseIcon       =   "Main.frx":16006
      Stretch         =   -1  'True
      ToolTipText     =   "Full View"
      Top             =   120
      Visible         =   0   'False
      Width           =   6615
   End
   Begin VB.Menu mnuIconMenu 
      Caption         =   "Icon Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuIconMenuGetScreen 
         Caption         =   "Get Screen"
      End
      Begin VB.Menu mnuIconMenuToolbar 
         Caption         =   "Toolbar"
      End
      Begin VB.Menu frmToolbarFullPreview 
         Caption         =   "Full Preview"
      End
      Begin VB.Menu mnuIconMenuClearViewer 
         Caption         =   "Clear Viewer"
      End
      Begin VB.Menu mnuIconMenuBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIconMenuPrefs 
         Caption         =   "Preferences"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuIconMenuResetDefaults 
         Caption         =   "Reset to Defaults"
      End
      Begin VB.Menu mnuIconMenuBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIconMenuShow 
         Caption         =   "Show"
      End
      Begin VB.Menu mnuIconMenuHide 
         Caption         =   "Hide"
      End
      Begin VB.Menu mnuIconMenuBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIconMenuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuIconMenuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileGetScreen 
         Caption         =   "&Get Screen"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileFullPreview 
         Caption         =   "&Full Preview"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileViewPreviousCapture 
         Caption         =   "&View Previous Capture"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClearMem 
         Caption         =   "Clear &Memory"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuFileClearViewer 
         Caption         =   "Clear &Viewer"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save Screenshot"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileBrowseSavePath 
         Caption         =   "Browse Save &Path"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileHide 
         Caption         =   "&Hide Me"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Visible         =   0   'False
      Begin VB.Menu mnuEditPrefs 
         Caption         =   "&Preferences"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewPicBox 
         Caption         =   "&Normal View"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewImgBox 
         Caption         =   "&Full View"
      End
      Begin VB.Menu mnuViewBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewSideToolbar 
         Caption         =   "&Side Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewQuickOptions 
         Caption         =   "&Quick-Options"
         Checked         =   -1  'True
         Visible         =   0   'False
      End
      Begin VB.Menu mnuViewTrayIcon 
         Caption         =   "Tray &Icon"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewToolBar 
         Caption         =   "&ToolBar"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuQuickOptions 
      Caption         =   "&Quick-Options"
      Begin VB.Menu mnuQuickOptionsCapture 
         Caption         =   "&Capture Whole Screen"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuQuickOptionsBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuickOptionsAutoSave 
         Caption         =   "Auto &Save"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuQuickOptionsAutoHide 
         Caption         =   "&Auto Hide"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuQuickOptionsBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuickOptionsHideMinimized 
         Caption         =   "&Hide when Minimized"
         Checked         =   -1  'True
         Visible         =   0   'False
      End
      Begin VB.Menu mnuQuickOptionsBar3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuQuickOptionsAllOn 
         Caption         =   "All O&n"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuQuickOptionsAllOff 
         Caption         =   "All O&ff"
      End
      Begin VB.Menu mnuQuickOptionsBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuickOptionsReset 
         Caption         =   "&Reset to Defaults"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
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

' For frmPrefrences
Dim CtrlCap
Dim CtrlHide
Dim CtrlToolBar
Dim CapLetter
Dim HideLetter
Dim ToolBarLetter

' For frmMain
Dim HideMain As Integer
Dim HideMinimize As Integer
Dim DelayScreenshot As Integer
Dim strSaveFileName As String
Dim strFilename As String
Dim BrowseSavePath As String
'Dim GetFileName As String

Private Sub chkAutoHide_Click()
    If chkAutoHide.Value = vbChecked Then
       mnuQuickOptionsAutoHide.Checked = True
    Else
       mnuQuickOptionsAutoHide.Checked = False
    End If
End Sub

Private Sub chkAutoSave_Click()
    If chkAutoSave.Value = vbChecked Then
       mnuQuickOptionsAutoSave.Checked = True
    Else
       mnuQuickOptionsAutoSave.Checked = False
    End If
End Sub

Private Sub cmdBrowse_Click()
    ' Browse for directory
    frmBrowse.Show
End Sub

Private Sub cmdClear_Click()

    If picSS.Picture = 0 Or imgSS.Picture = 0 Then
       MsgBox "Already Cleared!", , "Speed Screenshot Help"
       Exit Sub
    End If
    
    Dim AskClearPrev As Integer
    AskClearPrev = MsgBox("Are you sure you want to clear the preview boxs?", vbQuestion + vbYesNo, "Speed Speedshot")
    
    ' If user says yes
    If AskClearPrev = 6 Then
       picSS.Picture = LoadPicture("")
       imgSS.Picture = LoadPicture("")
       MsgBox "Preview boxs are now cleared.", , "Speed Screenshot"
    Else
       ' If user says no
       Exit Sub
    End If
    
End Sub

Private Sub cmdClearMemory_Click()
    
    Dim AskClearMemory As Integer
    AskClearMemory = MsgBox("Are you sure you want to clear memory?", vbQuestion + vbYesNo, "Speed Speedshot")
    
    ' If user says yes
    If AskClearMemory = 6 Then
       Clipboard.Clear
       picSS.Picture = LoadPicture("")
       imgSS.Picture = LoadPicture("")
       MsgBox "Memory is now cleared.", , "Speed Screenshot"
    Else
       ' If user says no
       Exit Sub
    End If
    
End Sub

Private Sub cmdPrefs_Click()
    frmPrefrences.Show
End Sub

Private Sub cmdSave_Click()
    Call mnuFileSave_Click
End Sub

Private Sub cmdToolbar_Click()
    frmToolbar.Show
End Sub

Private Sub cmdViewPrevSShot_Click()
    
    picSS.Picture = Clipboard.GetData
    imgSS.Picture = Clipboard.GetData
    
    If picSS.Picture = 0 Or imgSS.Picture = 0 Then
       MsgBox "Cant show the previous capture, memory is clear.", , "Speed Screenshot"
    End If
    
End Sub

Private Sub Form_Load()
    
    Me.Caption = "Stokes Speed Screenshot v" & App.Major & "." & App.Minor & App.Revision
    
    'mnuIconMenuShow.Enabled = False
    mnuFileFullPreview.Enabled = False
    mnuFileSave.Enabled = False
    
    ' Show the taskbar icon in the tray
    TrayControl1.Enabled = True
    
    ' Default save files path
    txtSaveFiles.Text = "" & App.Path
    
    ' Show Tool tip
    TrayControl1.ToolTipText = "Stokes Speed Screenshot v" & App.Major & "." & App.Minor & App.Revision
        
End Sub

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

Private Sub cmdGetScreen_Click()

    ' Auto hide main form if user has auto hide checked
    ' Checked by default.
    If mnuQuickOptionsAutoHide.Checked = True Then
       ' Hide me
       Me.Hide
       HideMain = 1
       ' Delay taking the screenshot so the form has time
       ' to hide first.
       tmrDelayScreenshot.Interval = DelayScreenshot
       DelayScreenshot = 500
    End If
    
    ' Get the screenshot
    Call GetScreenshot
    
End Sub

Private Sub GetScreenshot()

    ' Decide if the picure will displayed in a picture box
    ' or a image box on the main form.
    If picSS.Visible = True Then
       picSS.Enabled = True
       CopyToClipboard (chkWholeScreen.Value = vbUnchecked)
       picSS.Picture = Clipboard.GetData
       imgSS.Picture = Clipboard.GetData
         ' Save file automaticly if Auto-Save is checked
         If mnuQuickOptionsAutoSave.Checked = True Then
            SavePicture picSS.Picture, txtSaveFiles.Text & "\" & txtFileName.Text + txtSSNum.Text + ".bmp"
         End If
       txtSSNum.Text = txtSSNum.Text + 1
       cmdFullPreview.Enabled = True
    Else
       imgSS.Visible = True
       imgSS.Enabled = True
       CopyToClipboard (chkWholeScreen.Value = vbUnchecked)
       imgSS.Picture = Clipboard.GetData
       picSS.Picture = Clipboard.GetData
         ' Save file automaticly if Auto-Save is checked
         If mnuQuickOptionsAutoSave.Checked = True Then
            SavePicture picSS.Picture, txtSaveFiles.Text & "\" & txtFileName.Text + txtSSNum.Text + ".bmp"
         End If
       txtSSNum.Text = txtSSNum.Text + 1
       cmdFullPreview.Enabled = True
    End If

    ' Enable Full preview
    mnuFileFullPreview.Enabled = True
    
End Sub

Private Sub chkWholeScreen_Click()
    If chkWholeScreen.Value = vbChecked Then
       mnuQuickOptionsCapture.Checked = True
    Else
       mnuQuickOptionsCapture.Checked = False
    End If
End Sub

Private Sub cmdFullPreview_Click()

    If frmMain.imgSS.Picture = 0 Then
       MsgBox "No screenshot loaded!", , "Oops!"
       Exit Sub
    Else
        frmShowFullScreen.imgFS.Picture = frmMain.imgSS.Picture
        frmShowFullScreen.Show
    End If
    
End Sub

Private Sub cmdHide_Click()
    mnuIconMenuHide.Enabled = False
    mnuIconMenuShow.Enabled = True
    frmMain.Hide
    HideMain = 1
End Sub

Private Sub cmdReset_Click()
    
    Dim AskResetDefaults As Integer
    AskResetDefaults = MsgBox("Are you sure you want to reset everything to defaults?", vbQuestion + vbYesNo, "Speed Speedshot")
    
    ' If user says yes
    If AskResetDefaults = 6 Then
       ' Reset everything on the main form to defaults
       mnuViewPicBox.Checked = True
       chkWholeScreen.Value = vbChecked
       chkAutoHide.Value = vbChecked
       chkAutoSave.Value = vbChecked
       txtSSNum.Text = "0"
       txtFileName.Text = "screenshot"
       txtSaveFiles.Text = "" & App.Path
    Else
       ' If user says no
       Exit Sub
    End If
    
End Sub

Private Sub frmToolbarFullPreview_Click()
    Call cmdFullPreview_Click
End Sub

Private Sub imgSS_DblClick()
    Call cmdFullPreview_Click
End Sub

Private Sub mnuEditPrefs_Click()
    frmPrefrences.Show
End Sub

Private Sub mnuFileBrowseSavePath_Click()
    frmBrowse.Show
End Sub

Private Sub mnuFileClearMem_Click()
    Call cmdClearMemory_Click
End Sub

Private Sub mnuFileClearViewer_Click()
    Call cmdClear_Click
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
    End
End Sub

Private Sub mnuFileFullPreview_Click()
    Call cmdFullPreview_Click
End Sub

Private Sub mnuFileViewPreviousCapture_Click()
    Call cmdViewPrevSShot_Click
End Sub

Private Sub mnuFileGetScreen_Click()
    Call cmdGetScreen_Click
End Sub

Private Sub mnuFileHide_Click()
    Call cmdHide_Click
End Sub

Private Sub mnuFileSave_Click()
    
    ' If no screenshot taking yet alert user
    If picSS.Picture = 0 Then
       MsgBox "Nothing to save yet!", , "Speed Screenshot"
       Exit Sub
    End If
    
    ' Otherwise save it for user
    On Error Resume Next
    SavePicture picSS.Picture, txtSaveFiles.Text & "\" & txtFileName.Text + txtSSNum.Text + ".bmp"
    MsgBox "Picture Saved to: " & txtSaveFiles.Text & "\" & txtFileName.Text + txtSSNum.Text + ".bmp"

End Sub

Private Sub mnuFileSaveAs_Click()
    
    ' Display the File Save dialog box.
    frmSaveFile.dlgSaveFile.DialogTitle = "File Save"
    frmSaveFile.dlgSaveFile.Filter = "Bitmap Image (.bmp)|*.bmp"
    frmSaveFile.dlgSaveFile.FilterIndex = 1
    'frmSaveFile.dlgSaveFile.Flags = vbOFNSaveAs Or vbOFNFileMustExist
    frmSaveFile.dlgSaveFile.CancelError = True
    frmSaveFile.dlgSaveFile.FileName = ""
    On Error Resume Next
    frmSaveFile.dlgSaveFile.ShowSave
    
    ' Save the file with the new filename.
    frmSaveFile.dlgSaveFile.FileName = strSaveFileName
    strFilename = "untitled.bmp"
    ' Get the strFilename, and then call the save procedure, GetstrFilename.
    'strFilename = GetFileName(strFilename)
    'strSaveFileName = GetFileName("untitled.bmp")
    
    If strSaveFileName <> "" Then
       'SaveFileAs (strSaveFileName)
    End If
    
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuIconMenuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuIconMenuClearViewer_Click()
    Call cmdClear_Click
End Sub

Private Sub mnuIconMenuExit_Click()
    Unload Me
    End
End Sub

Private Sub mnuIconMenuGetScreen_Click()
    Call cmdGetScreen_Click
End Sub

Private Sub mnuIconMenuHide_Click()
    mnuIconMenuHide.Enabled = False
    mnuIconMenuShow.Enabled = True
    frmMain.Hide
    HideMain = 1
End Sub

Private Sub mnuIconMenuPrefs_Click()
    frmPrefrences.Show
End Sub

Private Sub mnuIconMenuResetDefaults_Click()
    Call cmdReset_Click
End Sub

Private Sub mnuIconMenuShow_Click()
    mnuIconMenuHide.Enabled = True
    mnuIconMenuShow.Enabled = False
    frmMain.Show
    HideMain = 0
End Sub

Private Sub mnuIconMenuToolbar_Click()
    frmToolbar.Show
End Sub

Private Sub mnuQuickOptionsAllOff_Click()
    
    mnuQuickOptionsAllOff.Checked = Not mnuQuickOptionsAllOff.Checked
    
    If mnuQuickOptionsAllOff.Checked = True Then
       opnAlloff.Value = True
       Call opnAlloff_Click
    Else
       Exit Sub
    End If
    
End Sub

Private Sub mnuQuickOptionsAllOn_Click()
    
    mnuQuickOptionsAllOn.Checked = Not mnuQuickOptionsAllOn.Checked
    
    If mnuQuickOptionsAllOn.Checked = True Then
       opnAllon.Value = True
       Call opnAllon_Click
    Else
       Exit Sub
    End If
    
End Sub

Private Sub mnuQuickOptionsAutoHide_Click()
    
    mnuQuickOptionsAutoHide.Checked = Not mnuQuickOptionsAutoHide.Checked
    
    If mnuQuickOptionsAutoHide.Checked = True Then
       chkAutoHide.Value = vbChecked
    Else
       chkAutoHide.Value = vbUnchecked
    End If
    
End Sub

Private Sub mnuQuickOptionsAutoSave_Click()
    
    mnuQuickOptionsAutoSave.Checked = Not mnuQuickOptionsAutoSave.Checked
    
    If mnuQuickOptionsAutoSave.Checked = True Then
       chkAutoSave.Value = vbChecked
    Else
       chkAutoSave.Value = vbUnchecked
    End If
    
End Sub

Private Sub mnuQuickOptionsCapture_Click()
    
    mnuQuickOptionsCapture.Checked = Not mnuQuickOptionsCapture.Checked
    
    If mnuQuickOptionsCapture.Checked = True Then
       chkWholeScreen.Value = vbChecked
    Else
       chkWholeScreen.Value = vbUnchecked
    End If
    
End Sub

Private Sub mnuQuickOptionsHideMinimized_Click()
    
    mnuQuickOptionsHideMinimized.Checked = Not mnuQuickOptionsHideMinimized.Checked
    
    ' Turn auto hide when minimized on and off
    If mnuQuickOptionsHideMinimized.Checked = True Then
       HideMinimize = 1
    Else
       HideMinimize = 0
    End If
       
End Sub

Private Sub mnuQuickOptionsReset_Click()
    Call cmdReset_Click
End Sub

Private Sub mnuViewImgBox_Click()
    
    ' Check and uncheck
    mnuViewImgBox.Checked = Not mnuViewImgBox.Checked
    
    ' Show img box
    If mnuViewImgBox.Checked = True Then
       picSS.Visible = False
       imgSS.Visible = True
       mnuViewPicBox.Checked = False
    Else
       ' Keep it checked
       mnuViewImgBox.Checked = True
    End If
    
End Sub

Private Sub mnuViewPicBox_Click()
    
    ' Check and uncheck
    mnuViewPicBox.Checked = Not mnuViewPicBox.Checked
    
    ' Show pic boxes
    If mnuViewPicBox.Checked = True Then
       picSS.Visible = True
       imgSS.Visible = False
       mnuViewImgBox.Checked = False
    Else
       ' Keep it checked
       mnuViewPicBox.Checked = True
    End If
    
End Sub

Private Sub mnuViewQuickOptions_Click()

    mnuViewQuickOptions.Checked = Not mnuViewQuickOptions.Checked
    
    If mnuViewQuickOptions.Checked = True Then
       fraQO.Visible = True
    Else
       fraQO.Visible = False
    End If
    
    ' Turn on the timer just long enough to run
    tmrShowHideTBnQO.Enabled = True

End Sub

Private Sub mnuViewSideToolbar_Click()
    
    mnuViewSideToolbar.Checked = Not mnuViewSideToolbar.Checked
    
    If mnuViewSideToolbar.Checked = True Then
       fraMainMenu.Visible = True
       picSS.Left = 1080
       imgSS.Left = 1080
       fraQO.Left = 1080
       frmMain.Width = 7905
    Else
       fraMainMenu.Visible = False
       picSS.Left = 120
       imgSS.Left = 120
       fraQO.Left = 120
       frmMain.Width = 6945
    End If
    
    ' Turn on the timer just long enough to run
    'tmrShowHideTBnQO.Enabled = True
    
End Sub

Private Sub mnuViewToolBar_Click()
    frmToolbar.Show
End Sub

Private Sub mnuViewTrayIcon_Click()
    
    mnuViewTrayIcon.Checked = Not mnuViewTrayIcon.Checked
    
    If mnuViewTrayIcon.Checked = True Then
       TrayControl1.Enabled = True
       mnuIconMenuHide.Enabled = True
       mnuFileHide.Enabled = True
       cmdHide.Enabled = True
       mnuQuickOptionsAutoHide.Enabled = True
       chkAutoHide.Enabled = True
    Else
       TrayControl1.Enabled = False
       mnuIconMenuHide.Enabled = False
       mnuFileHide.Enabled = False
       cmdHide.Enabled = False
       mnuQuickOptionsAutoHide.Enabled = False
       mnuQuickOptionsAutoHide.Checked = False
       chkAutoHide.Value = vbUnchecked
       chkAutoHide.Enabled = False
    End If
    
    If mnuFileHide.Enabled = False Then
       Exit Sub
    End If
    
    If opnAllon.Value = True Then
       chkAutoHide.Value = vbChecked
    Else
       chkAutoHide.Value = vbUnchecked
    End If
    
End Sub

Private Sub opnAlloff_Click()

    opnAllon.Value = False
    mnuQuickOptionsAllOff.Checked = True
    mnuQuickOptionsAllOn.Checked = False

    chkWholeScreen.Value = vbUnchecked
    chkAutoHide.Value = vbUnchecked
    chkAutoSave.Value = vbUnchecked
    
    Call chkWholeScreen_Click
    Call chkAutoHide_Click
    Call chkAutoSave_Click


End Sub

Private Sub opnAllon_Click()

    opnAlloff.Value = False
    mnuQuickOptionsAllOn.Checked = True
    mnuQuickOptionsAllOff.Checked = False

    chkWholeScreen.Value = vbChecked
    chkAutoHide.Value = vbChecked
    chkAutoSave.Value = vbChecked
    
    ' Determine if Auto-Hide is enabled or not
    If chkAutoHide.Enabled = False Then
       chkAutoHide.Value = vbUnchecked
    Else
       chkAutoHide.Value = vbChecked
    End If
    
    Call chkWholeScreen_Click
    Call chkAutoHide_Click
    Call chkAutoSave_Click
    
End Sub

Private Sub picSS_DblClick()
    Call cmdFullPreview_Click
End Sub

Private Sub tmrAutoHide_Timer()
    
    ' If HideMain = true then enable show in menu
    ' 1 = True
    ' 0 = False
    If HideMain = 1 Then
       frmMain.mnuIconMenuHide.Enabled = False
       frmMain.mnuIconMenuShow.Enabled = True
    Else
       frmMain.mnuIconMenuHide.Enabled = True
       frmMain.mnuIconMenuShow.Enabled = False
    End If
    
End Sub

Private Sub tmrCheckToAutoSave_Timer()
    ' Constantly check to see if save should be enabled
    ' or not.
    If mnuQuickOptionsAutoSave.Checked = True Then
       mnuFileSave.Enabled = False
       mnuFileSaveAs.Enabled = False
    Else
       mnuFileSave.Enabled = True
       mnuFileSaveAs.Enabled = True
    End If
End Sub

Private Sub tmrHideMinimized_Timer()
    
    'Call mnuQuickOptionsHideMinimized_Click
    
    ' If auto hide minimize = true
    If frmMain.WindowState = vbMinimized And HideMinimize = 1 Then
       mnuIconMenuHide.Enabled = False
       mnuIconMenuShow.Enabled = True
       frmMain.Hide
       Exit Sub
    Else
       mnuIconMenuHide.Enabled = True
       mnuIconMenuShow.Enabled = False
       Exit Sub
    End If
    
    ' If auto hide minimize = false
    If frmMain.WindowState = vbMinimized And HideMinimize = 0 Then
       Exit Sub
    End If

End Sub

Private Sub tmrHKFullPreview_Timer()
    ' If user presses Ctrl+P then show the toolbar
    If GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyP) Then
       Call cmdFullPreview_Click
    End If
End Sub

Private Sub tmrHKHideme_Timer()
    ' If user presses Ctrl+H then hide me
    If GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyH) Then
       Me.Hide
    End If
End Sub

Private Sub tmrHKResetDefaults_Timer()
    ' If user presses Ctrl+D then reset to defaults
    If GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyD) Then
       Call cmdReset_Click
    End If
End Sub

Private Sub tmrHKScreenshot_Timer()
    ' If user presses Ctrl+G then take a screenshot
    If GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyG) Then
       Call cmdGetScreen_Click
    End If
End Sub

Private Sub tmrHKToolbar_Timer()
    ' If user presses Ctrl+T then show the toolbar
    If GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyT) Then
       frmToolbar.Show
    End If
End Sub

Private Sub tmrHKSave_Timer()
    ' If user presses Ctrl+S then save the file
    If GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyS) Then
       Call mnuFileSave_Click
    End If
End Sub

Private Sub tmrShowHideTBnQO_Timer()
    
    ' Resize main form if Quick-Options frame is already invisiable
    If fraQO.Visible = True Then
       frmMain.Height = 8280
    Else
       frmMain.Height = 6240
    End If
    
    ' Resize main form is side toolbar already invisiable
    If fraMainMenu.Visible = True Then
       frmMain.Height = 8280
    Else
       frmMain.Height = 6240
    End If
    
    ' Turn it self off after its done above
    tmrShowHideTBnQO.Enabled = False
    
End Sub

Private Sub TrayControl1_DblClick()
    Call mnuIconMenuShow_Click
    frmMain.Show
End Sub

Private Sub TrayControl1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuIconMenu
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   Cancel = 1
   Dim QuitApp As Integer
   QuitApp = MsgBox("Do you really want to quit?", vbQuestion + vbYesNo, "Confirm Exit")
    
   ' If user says yes then unload and terminate program.
   If QuitApp = 6 Then
      TrayControl1.Enabled = False
      Unload Me
      End
   End If
    
End Sub
