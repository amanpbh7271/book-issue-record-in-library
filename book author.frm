VERSION 5.00
Begin VB.Form b_author 
   BackColor       =   &H00C0C0FF&
   Caption         =   "author"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14610
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   14610
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   8
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   9120
      TabIndex        =   7
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   4680
      TabIndex        =   6
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   6840
      TabIndex        =   5
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
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
      Height          =   495
      Index           =   3
      Left            =   2640
      TabIndex        =   4
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NEW"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   4320
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   2
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "  AUTHOR ENTRY FORM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1560
      TabIndex        =   3
      Top             =   720
      Width           =   8775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "AUTHOR NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   2280
      Width           =   3975
   End
End
Attribute VB_Name = "b_author"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rst_author As ADODB.Recordset
Public u_d As Boolean
Public no As Integer
Public seid As Integer


Sub command_button(b1 As Boolean, b2 As Boolean, b3 As Boolean, b4 As Boolean, b5 As Boolean)
    Command1(0).Enabled = b1
    Command1(3).Enabled = b2
    Command1(5).Enabled = b3
    Command1(4).Enabled = b4
    Command1(1).Enabled = b5
End Sub

Private Sub Command2_Click()
seid = 0
u_d = True
frmbookauthor.Show 1
''MsgBox "record"
 If seid <> 0 Then
Text1.Enabled = True
 Dim rst_author As ADODB.Recordset
 Set rst_author = New ADODB.Recordset
 rst_author.Open "select * from author where author_no=" & Val(seid) & "", cn
 Text1.Text = rst_author.Fields(1)

 Set rst_author = Nothing
 Call command_button(False, True, True, True, False)
 End If

End Sub

Private Sub Form_Load()
Text1.Enabled = False
Call command_button(True, False, False, False, True)
End Sub

Sub reset()
Text1.Text = ""
Text1.Enabled = False
End Sub
Private Sub Command1_Click(Index As Integer)
 Dim rst_author As ADODB.Recordset
 Set rst_author = New ADODB.Recordset
 Select Case Index
  Case 0
        Call command_button(False, True, False, True, False)
        Text1.Text = ""
        Text1.Enabled = True
        no = next_no("select max(author_no) from author")
        u_d = False
        


        
  Case 3
       If u_d = True Then
       If Text1.Text = "" Then
       MsgBox "Enter All Details"
        Else
        rst_author.Open "update author set author_name=' " & Text1.Text & "' where author_no=" & seid & "", cn
         MsgBox "record updated successfully"
   
       ''rst_author.Open "insert into author values (" & no & ",' " & Text1.Text & "' )", cn
       
       ''MsgBox " New record stored successfully"
       Call reset
    
    
      Call command_button(True, False, False, False, True)
      
      End If
   Set rst_author = Nothing
    u_d = False
  Else
    If Text1.Text = "" Then
     MsgBox "Enter All Details"
    Else
     ''rst_author.Open "update author set author_name=' " & Text1.Text & "' where author_no=" & seid & "", cn
       ''  MsgBox "record updated successfully"
       rst_author.Open "insert into author values (" & no & ",' " & Text1.Text & "' )", cn
       
       MsgBox " New record stored successfully"
       
     Call reset
     Call command_button(True, False, False, False, True)
    End If
    Set rst_author = Nothing
  End If
   
 
      
  Case 5
      
      res = MsgBox("Do you want to delete the record ?", vbYesNo)
     If res = 7 Then
        Exit Sub
     End If
     If Text1.Text <> "" Then
      rst_author.Open "delete from author where author_no=" & seid & "", cn
      MsgBox "record deleted successfully"
      Call reset
     Else
      MsgBox "record changed! Deletion can not be done  ", vbCritical
     End If
      
      
  Case 4
  Call command_button(True, False, True, False, True)
    Call reset
  
  Case 1
  
    Unload Me
 End Select
  
 
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text1.Text <> "" Then
Command1(3).SetFocus
End If
Select Case KeyAscii
 Case 65 To 122
 Case 13, 8, 32
 Case Else
 KeyAscii = 0
 End Select
 

End Sub

