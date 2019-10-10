VERSION 5.00
Begin VB.Form login 
   BackColor       =   &H80000002&
   Caption         =   "Form5"
   ClientHeight    =   8565
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18735
   LinkTopic       =   "Form5"
   ScaleHeight     =   8565
   ScaleWidth      =   18735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   4
      Top             =   4920
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   3
      Top             =   4080
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   7680
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   1
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "LOGIN FORM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   3360
      TabIndex        =   5
      Top             =   1920
      Width           =   8895
      Begin VB.Label Label3 
         Caption         =   "PASSWORD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "USER NAME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   6
         Top             =   720
         Width           =   2295
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   "BOOK ISSUE RECORD IN  LIBRARY"
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
      Left            =   2280
      TabIndex        =   0
      Top             =   360
      Width           =   11775
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim r As New ADODB.Recordset
Dim i As Integer

Private Sub Command1_Click()
r.Open "select * from use", cn
For i = 1 To r.RecordCount
If r.Fields(0) = Text1.Text And r.Fields(1) = Text2.Text Then
MsgBox "login successful"
MDIForm1.Show
Unload Me
Else
MsgBox "login failed"
End If
Next
r.Close


End Sub


Private Sub Command2_Click()

Text1.Text = ""
Text2.Text = ""
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 And Text1.Text <> "" Then
Text2.SetFocus
End If
Select Case KeyAscii
 Case 65 To 122
 Case 13, 8, 32
 Case Else
 KeyAscii = 0
 End Select
End Sub

Private Sub text1_validate(cancel As Boolean)
If Command2.Value = False Then
If Text1.Text = "" Then
MsgBox "plaese enter user name"
cancel = True
End If
End If
End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text2.Text <> "" Then
Command1.SetFocus
End If
Select Case KeyAscii
 Case 65 To 122
 Case 13, 8, 32
 Case Else
 KeyAscii = 0
 End Select
End Sub

Private Sub text2_validate(cancel As Boolean)
If Text2.Text = "" Then
MsgBox "plaese enter password"
cancel = True
End If
End Sub
