VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPrefrences 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stokes Speedshot - Preferences"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6375
   ControlBox      =   0   'False
   Icon            =   "Prefrences.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4455
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   7858
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Hot Keys"
      TabPicture(0)   =   "Prefrences.frx":2E7A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraCapture"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraHide"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdDefaultHotKeys"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraToolBar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Program Options"
      TabPicture(1)   =   "Prefrences.frx":2E96
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdDefaultGeneral"
      Tab(1).Control(1)=   "fraGeneral"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Notifications"
      TabPicture(2)   =   "Prefrences.frx":2EB2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   1215
         Left            =   2520
         TabIndex        =   19
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Frame fraToolBar 
         Caption         =   "Tool Bar"
         Height          =   1215
         Left            =   360
         TabIndex        =   18
         Top             =   2040
         Width           =   1935
         Begin VB.CheckBox chkCtrlToolBar 
            Caption         =   "+"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   600
            Value           =   1  'Checked
            Width           =   495
         End
         Begin VB.ComboBox cboToolBar 
            Height          =   315
            ItemData        =   "Prefrences.frx":2ECE
            Left            =   720
            List            =   "Prefrences.frx":2F20
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   600
            Width           =   855
         End
         Begin VB.Label lblCtrlTBar 
            Caption         =   "CTRL"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.CommandButton cmdDefaultGeneral 
         Caption         =   "Defaults"
         Height          =   375
         Left            =   -70440
         TabIndex        =   17
         Top             =   1560
         Width           =   975
      End
      Begin VB.Frame fraGeneral 
         Caption         =   "General Options"
         Height          =   1575
         Left            =   -74640
         TabIndex        =   13
         Top             =   600
         Width           =   5415
         Begin VB.CheckBox chkUseTrayIcon 
            Caption         =   "Use tray icon"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox chkAutoSave 
            Caption         =   "Automatically Save Screenshots"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   1080
            Value           =   1  'Checked
            Width           =   2655
         End
         Begin VB.CheckBox chkHideMe 
            Caption         =   "Hide before capturing"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   720
            Value           =   1  'Checked
            Width           =   1935
         End
      End
      Begin VB.CommandButton cmdDefaultHotKeys 
         Caption         =   "Defaults"
         Height          =   375
         Left            =   4800
         TabIndex        =   12
         Top             =   2880
         Width           =   975
      End
      Begin VB.Frame fraHide 
         Caption         =   "Hide"
         Height          =   1215
         Left            =   2520
         TabIndex        =   8
         Top             =   600
         Width           =   1935
         Begin VB.CheckBox chkCtrlHide 
            Caption         =   "+"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   600
            Value           =   1  'Checked
            Width           =   495
         End
         Begin VB.ComboBox cboHide 
            Height          =   315
            ItemData        =   "Prefrences.frx":2F72
            Left            =   720
            List            =   "Prefrences.frx":2FC4
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   600
            Width           =   855
         End
         Begin VB.Label lblctrlhide 
            Caption         =   "CTRL"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame fraCapture 
         Caption         =   "Capture"
         Height          =   1215
         Left            =   360
         TabIndex        =   4
         Top             =   600
         Width           =   1935
         Begin VB.ComboBox cboCapture 
            Height          =   315
            ItemData        =   "Prefrences.frx":3016
            Left            =   720
            List            =   "Prefrences.frx":3068
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   600
            Width           =   855
         End
         Begin VB.CheckBox chkCtrlCap 
            Caption         =   "+"
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   600
            Value           =   1  'Checked
            Width           =   495
         End
         Begin VB.Label lblCtrlcap 
            Caption         =   "CTRL"
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   360
            Width           =   495
         End
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4680
      Width           =   1095
   End
End
Attribute VB_Name = "frmPrefrences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim CtrlCap
Dim CtrlHide
Dim CtrlToolBar
Dim CapLetter
Dim HideLetter
Dim ToolBarLetter

Private Sub cmdApply_Click()
    Call SetPrefs
End Sub

Private Sub SetPrefs()

' --------------- Hotkeys Tab ----------------------
    
    ' Assign variables to checks on prefs form
    CtrlCap = chkCtrlCap.Value
    CtrlHide = chkCtrlHide.Value
    CtrlToolBar = chkCtrlToolBar.Value
    CapLetter = cboCapture.Text
    HideLetter = cboHide.Text
    ToolBarLetter = cboToolBar.Text
    
    ' Assign Ctrl Capture key if checked
    If chkCtrlCap.Value = vbChecked Then
       CtrlCap = vbKeyControl
    Else
       CtrlCap = 0
    End If
    
    ' Assign Ctrl Hide key if checked
    If chkCtrlHide.Value = vbChecked Then
       CtrlHide = vbKeyControl
    Else
       CtrlHide = 0
    End If
    
    ' Assign Ctrl Toolbar key if checked
    If chkCtrlToolBar.Value = vbChecked Then
       CtrlToolBar = vbKeyControl
    Else
       CtrlToolBar = 0
    End If
    
    ' If both boxs equal same then tell user
    If cboCapture.Text = cboHide.Text Then
       MsgBox "Cant be same value", , "Screenshot Help"
       Exit Sub
    End If
    
' --------------- Prefrences Tab -------------------
    
    ' Set tray icon
    If chkUseTrayIcon.Value = vbChecked Then
       frmMain.TrayControl1.Enabled = True
    Else
       frmMain.TrayControl1.Enabled = False
    End If
    
    ' Set hide me
    If chkHideMe.Value = vbChecked Then
       frmMain.mnuQuickOptionsAutoHide.Checked = True
    Else
       frmMain.mnuQuickOptionsAutoHide.Checked = False
    End If
    
    ' Set auto save
    If chkAutoSave.Value = vbChecked Then
       frmMain.mnuQuickOptionsAutoSave.Checked = True
    Else
       frmMain.mnuQuickOptionsAutoSave.Checked = False
    End If
    
' --------------- Notifications Tab ----------------


' --------------------------------------------------

    'Add code here to save settings without closeing form
    cmdApply.Enabled = False
    
End Sub

Private Sub cboToolBar_Click()
    cmdApply.Enabled = True
End Sub

Private Sub chkAutoSave_Click()
    cmdApply.Enabled = True
End Sub

Private Sub chkCtrlCap_Click()
    cmdApply.Enabled = True
End Sub

Private Sub chkCtrlToolBar_Click()
    cmdApply.Enabled = True
End Sub

Private Sub chkHideMe_Click()
    cmdApply.Enabled = True
End Sub

Private Sub chkUseTrayIcon_Click()
    cmdApply.Enabled = True
End Sub

Private Sub cmdDefaultGeneral_Click()
    chkUseTrayIcon.Value = vbChecked
    chkHideMe.Value = vbChecked
    chkAutoSave.Value = vbChecked
End Sub

Private Sub Form_Load()
    ' Set Defaults on load
    cboCapture.Text = "G"
    cboHide.Text = "H"
    cboToolBar.Text = "T"
    cmdApply.Enabled = False
End Sub

Private Sub cboCapture_Click()
    cmdApply.Enabled = True
End Sub

Private Sub cboHide_Click()
    cmdApply.Enabled = True
End Sub

Private Sub chkCtrlHide_Click()
    cmdApply.Enabled = True
End Sub

Private Sub chkCtrlHotkey_Click()
    cmdApply.Enabled = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDefaultHotKeys_Click()
    cmdApply.Enabled = True
    cboCapture.Text = "G"
    cboHide.Text = "H"
End Sub

Private Sub cmdOK_Click()
    cmdApply.Enabled = False
    'Add code here to save settings and close form
    Call SetPrefs
    Unload Me
End Sub
