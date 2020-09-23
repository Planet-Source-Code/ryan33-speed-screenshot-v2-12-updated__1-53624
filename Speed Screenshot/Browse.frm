VERSION 5.00
Begin VB.Form frmBrowse 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Browse for Save Directory"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   3375
   Icon            =   "Browse.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   3375
   StartUpPosition =   2  'CenterScreen
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00FFFFFF&
      Height          =   2790
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   3135
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3135
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0C0C0&
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3360
      Width           =   1095
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   600
      Pattern         =   "*.mp3"
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtSaveSettings 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   3135
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private fso As New FileSystemObject
'Private strm As TextStream
Private strName As String
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

' Set directory to save files
File1.Path = Dir1.Path
'File1.FileName = "savefiles.ini"
   
' Display path on main form
frmMain.txtSaveFiles.Text = Dir1.Path & File1.FileName
Unload Me

' Save setting to a file
'SaveSettings frmMain.txtSaveFiles.Text
    
End Sub

Private Sub GetSettings(ByVal FileName As String)
    Set strm = fso.OpenTextFile(FileName, ForReading)
    With strm
        frmMain.txtSaveFiles.Text = .ReadLine
        frmMain.txtSaveFiles.Text = FileName
        .Close
    End With
End Sub

Private Sub SaveSettings(ByVal FileName As String)
    Set strm = fso.CreateTextFile(FileName, True)
    With strm
        .WriteLine frmMain.txtSaveFiles.Text
        .Close
    End With
End Sub

Private Sub Drive1_Change()

    ' If Drive not valid show megbox and exit sub
    If Dir1.Path = "" Then
        MsgBox "Directory not valid!"
        Exit Sub
    Else
        ' Otherwise set save path
        On Error Resume Next
        Dir1.Path = Drive1.Drive
    End If

End Sub

Private Sub Form_Load()
    'txtSaveSettings.Text = App.Path & "\savefiles.ini"
    'txtSaveSettings.Text = App.Path & "\Settings.BartNet"
End Sub
