VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form welcome 
   Caption         =   "Form3"
   ClientHeight    =   4335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14280
   LinkTopic       =   "Form3"
   ScaleHeight     =   4335
   ScaleWidth      =   14280
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   360
      Top             =   1800
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   615
      Left            =   4440
      TabIndex        =   0
      Top             =   2640
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   1085
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      OLEDropMode     =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Caption         =   "loading......"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      TabIndex        =   1
      Top             =   1440
      Width           =   4335
   End
End
Attribute VB_Name = "welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
ProgressBar1.Value = 0
ProgressBar1.Max = 100
End Sub

Private Sub Timer1_Timer()
If ProgressBar1.Value < 100 Then
  ProgressBar1.Value = ProgressBar1.Value + 1
Else
  login.Show
  Unload Me
  End If
End Sub
