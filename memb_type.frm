VERSION 5.00
Begin VB.Form memb_type 
   Caption         =   "memb_type"
   ClientHeight    =   4515
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9090
   LinkTopic       =   "Form3"
   ScaleHeight     =   4515
   ScaleWidth      =   9090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   495
      Left            =   4920
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "exit"
      Height          =   495
      Index           =   4
      Left            =   9600
      TabIndex        =   6
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "cancel"
      Height          =   495
      Index           =   3
      Left            =   7440
      TabIndex        =   5
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "delete"
      Height          =   495
      Index           =   2
      Left            =   5280
      TabIndex        =   4
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "save"
      Height          =   495
      Index           =   1
      Left            =   2880
      TabIndex        =   3
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "new"
      Height          =   495
      Index           =   0
      Left            =   840
      TabIndex        =   2
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "member type"
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "memb_type"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim id As Long
Public seid As Integer
Public u_d As Boolean



Sub command(b1 As Boolean, b2 As Boolean, b3 As Boolean, b4 As Boolean, b5 As Boolean)
    Command1(0).Enabled = b1
    Command1(1).Enabled = b2
    Command1(2).Enabled = b3
    Command1(3).Enabled = b4
    Command1(4).Enabled = b5
End Sub

Private Sub Command2_Click()
seid = 0
u_d = True
frm_memb_type.Show 1
MsgBox "record"
 If seid <> 0 Then
 Dim rst_memb_type As ADODB.Recordset
 Set rst_memb_type = New ADODB.Recordset
 rst_memb_type.Open "select * from memb_type where memb_id=" & Val(seid) & "", cn
 MsgBox rst_memb_type.RecordCount
 Text1.Text = rst_memb_type.Fields(1)
 Set rst_memb_type = Nothing
 Call command(False, True, True, True, False)
 End If

End Sub

Private Sub Form_Load()
u_d = True
Call command(True, False, False, False, True)

End Sub
Private Sub Command1_Click(Index As Integer)
 Dim rst_memb_type As ADODB.Recordset
 Set rst_memb_type = New ADODB.Recordset
 Select Case Index
  Case 0
        Call command(False, True, False, True, False)
        
        Text1.Text = ""
        
        id = next_id("select max(memb_id) from memb_type")
        u_d = False
  Case 1
  If u_d = False Then
    If Text1.Text = "" Then
     MsgBox "Enter Some Detail"
    Else
     rst_memb_type.Open "insert into memb_type values (" & id & ",' " & Text1.Text & " ')", cn
     MsgBox " New record stored successfully"
     Call reset
     Call command(True, False, False, False, True)
    End If
    Set rst_memb_type = Nothing
    u_d = True
  Else
    If Text1.Text = "" Then
     MsgBox "Enter Some Detail"
    Else
     rst_memb_type.Open "update memb_type set m_type = '" & Text1.Text & "'   where memb_id = " & seid & "", cn
     MsgBox "record updated successfully"
     Call reset
     Call command(True, False, False, False, True)
    End If
    Set rst_memb_type = Nothing
  End If
  Case 2
     res = MsgBox("Do you want to delete the record ?", vbYesNo)
     If res = 7 Then
        Exit Sub
     End If
     If Text1.Text <> "" Then
      rst_memb_type.Open "delete from memb_type where memb_id=" & seid & "", cn
      MsgBox "record deleted successfully"
      Call reset
     Else
      MsgBox "record changed! Deletion can not be done  ", vbCritical
     End If
  Case 3
    Call command(True, False, False, False, True)
    Call reset
  Case 4
   Unload Me
 End Select
End Sub
  
  Sub reset()
   Text1.Text = ""
  End Sub
  
  
