VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00E0E0E0&
   Caption         =   "About Stokes Speed Screenshot"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   5535
   ControlBox      =   0   'False
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Credits"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1335
      Left            =   240
      TabIndex        =   6
      Top             =   3000
      Width           =   5055
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Virgil Mutz"
         Height          =   255
         Left            =   2760
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Testers:"
         Height          =   255
         Left            =   1560
         TabIndex        =   9
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Ryan E. Stokes"
         Height          =   255
         Left            =   2760
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Written and Designed by:"
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame fraDescription 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1815
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   5055
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"About.frx":2E7A
         Height          =   1215
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   4575
      End
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0C0C0&
      Caption         =   "OK"
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label lblCC 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 2003 - Stokes Software Co."
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   4800
      Width           =   2895
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "version appears here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   5055
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Stokes Speed Screenshot"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version: " & App.Major & "." & App.Minor & App.Revision
End Sub
