VERSION 5.00
Begin VB.Form publisher 
   BackColor       =   &H00C0C0FF&
   Caption         =   "publisher"
   ClientHeight    =   8010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15165
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8010
   ScaleWidth      =   15165
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   9240
      TabIndex        =   10
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
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
      Left            =   9480
      TabIndex        =   9
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   2760
      TabIndex        =   8
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CANCLE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   7080
      TabIndex        =   7
      Top             =   4560
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   4800
      TabIndex        =   6
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "NEW"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   4560
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   2
      Top             =   3000
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   1
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "PUBLISHER ADDRESS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   5
      Top             =   3000
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "PUBLISHER NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   4
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "PUBLISHER ENTRY FORM"
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
      Left            =   1800
      TabIndex        =   3
      Top             =   480
      Width           =   8895
   End
End
Attribute VB_Name = "publisher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rst_publisher As New ADODB.Recordset
Public u_d As Boolean
Public id As Integer
Public seid As String

Sub command_button(b1 As Boolean, b2 As Boolean, b3 As Boolean, b4 As Boolean, b5 As Boolean)
    Command1(0).Enabled = b1
    Command1(5).Enabled = b2
    Command1(2).Enabled = b3
    Command1(3).Enabled = b4
    Command1(1).Enabled = b5
End Sub
Private Sub Command1_Click(Index As Integer)
'' Dim rst_publisher As ADODB.Recordset
 ''Set rst_pubblisher = New ADODB.Recordset
 Select Case Index
  Case 0
        Call command_button(False, True, False, True, False)
        Text1.Text = ""
        Text2.Text = ""
        Text1.Enabled = True
        Text2.Enabled = True
        id = next_id("select max(pub_id) from publisher")
        u_d = False
        
  Case 5
  If u_d = False Then
       If Text1.Text = "" Or Text2.Text = "" Then
       MsgBox "Enter All Details"
       Else
       Set rst_publisher = New ADODB.Recordset
        rst_publisher.Open "insert into publisher values(" & id & ",'" & Text1.Text & "',' " & Text2.Text & "')", cn
           MsgBox " New record stored successfully"
       Call reset
    
    
      Call command_button(True, False, False, False, True)
      
      End If
    Set rst_publisher = Nothing
    u_d = True
  Else
    If Text1.Text = "" Or Text2.Text = "" Then
     MsgBox "Enter All Details"
    Else
     Set rst_publisher = New ADODB.Recordset
     rst_publisher.Open "update publisher set name=' " & Text1.Text & " ',address=' " & Text2.Text & "' where pub_id=" & seid & "", cn
     MsgBox "record updated successfully"
     Call reset
     Call command_button(True, False, False, False, True)
    End If
    Set rst_publisher = Nothing
  End If
   
 
      
  Case 2
      
      res = MsgBox("Do you want to delete the record ?", vbYesNo)
     If res = 7 Then
        Exit Sub
     End If
     If Text1.Text <> "" And Text2.Text <> "" Then
      rst_publisher.Open "delete from publisher where pub_id=" & seid & "", cn
      MsgBox "record deleted successfully"
      Call reset
      Set rst_publisher = Nothing
     Else
      MsgBox "record changed! Deletion can not be done  ", vbCritical
     End If
      
      
  Case 3
  Call command_button(True, False, True, False, True)
    Call reset
  
  Case 1
  
   Unload Me
 End Select
  
 
End Sub




Private Sub Command2_Click()
seid = ""
u_d = True
frmPublisher.Show 1
''MsgBox "record"
 Text1.Enabled = True
Text2.Enabled = True
 If seid <> "" Then
 Dim rst_publisher As ADODB.Recordset
 Set rst_publisher = New ADODB.Recordset
 rst_publisher.Open "select * from publisher where pub_id=" & Val(seid) & "", cn
'' MsgBox rst_publisher.RecordCount
 Text1.Text = rst_publisher.Fields(1)
 Text2.Text = rst_publisher.Fields(2)
 Set rst_publisher = Nothing
 Call command_button(False, True, True, True, False)
 End If
 
End Sub



Private Sub Form_Load()
u_d = True
Call command_button(True, False, False, False, True)
Text1.Enabled = False
Text2.Enabled = False
End Sub
Sub reset()
Text1.Text = ""
Text2.Text = ""
Text1.Enabled = False
Text2.Enabled = False
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
Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text2.Text <> "" Then
Command1(5).SetFocus
End If
Select Case KeyAscii
 Case 65 To 122
 Case 13, 8, 32
 Case Else
 KeyAscii = 0
 End Select
 

End Sub

