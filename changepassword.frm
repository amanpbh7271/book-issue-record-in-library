VERSION 5.00
Begin VB.Form changepassword 
   Caption         =   "Form3"
   ClientHeight    =   4080
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10500
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   4080
   ScaleWidth      =   10500
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "CHANGE PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Left            =   4560
      TabIndex        =   0
      Top             =   840
      Width           =   7695
      Begin VB.CommandButton Command3 
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1560
         TabIndex        =   11
         Top             =   6240
         Width           =   3615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1560
         TabIndex        =   10
         Top             =   5280
         Width           =   3615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "SAVE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1560
         TabIndex        =   9
         Top             =   4320
         Width           =   3615
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4200
         TabIndex        =   4
         Text            =   "Text4"
         Top             =   3120
         Width           =   2415
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4200
         TabIndex        =   3
         Text            =   "Text3"
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4200
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4200
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080C0FF&
         Caption         =   "CONFIRM NEW  PASSWORD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   8
         Top             =   3240
         Width           =   3255
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080C0FF&
         Caption         =   "NEW PASSWORD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   7
         Top             =   2280
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080C0FF&
         Caption         =   "OLD PASSWORD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   6
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080C0FF&
         Caption         =   "USER NAME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   2775
      End
   End
End
Attribute VB_Name = "changepassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim use As New ADODB.Recordset
Dim use1 As New ADODB.Recordset

Private Sub Command1_Click()
use.Open "select * from use", cn
If use.Fields(0) = Text1.Text And use.Fields(1) = Text2.Text Then
 use.Close
   If Text3.Text = Text4.Text Then
   use1.Open "update use set username='" & Text1.Text & "' ,  pass='" & Text4.Text & "'", cn
   
   MsgBox "password is change succussfully......."
    Set user1 = Nothing
Else

  MsgBox "please confirm your password"
  
  
  End If
Else
MsgBox "please write valid user name and valid password"
use.Close
End If



End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text1.Text <> "" Then
Text2.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text2.Text <> "" Then
Text3.SetFocus
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text3.Text <> "" Then
Text4.SetFocus
End If
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text4.Text <> "" Then
Command1.SetFocus
End If
End Sub

