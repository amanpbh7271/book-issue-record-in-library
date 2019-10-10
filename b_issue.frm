VERSION 5.00
Begin VB.Form b_issue 
   Caption         =   "b_issue"
   ClientHeight    =   6615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9420
   LinkTopic       =   "Form2"
   ScaleHeight     =   6615
   ScaleWidth      =   9420
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Index           =   4
      Left            =   7560
      TabIndex        =   12
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Index           =   3
      Left            =   5760
      TabIndex        =   11
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Index           =   2
      Left            =   3960
      TabIndex        =   10
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Index           =   1
      Left            =   1920
      TabIndex        =   9
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   8
      Top             =   4920
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   3000
      TabIndex        =   7
      Text            =   "Combo2"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3120
      TabIndex        =   6
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      Top             =   2400
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2880
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "return_date"
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "issue_date"
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "available"
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "member_name"
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "b_issue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call cmb_populate1(Combo1, "select * from member")
Call cmb_populate2(Combo2, "select * from copy")
End Sub
Sub cmb_populate1(cmb As ComboBox, qry As String)
               Dim rst As New ADODB.Recordset
               Dim i As Integer
               i = 1
               rst.Open qry, cn
               cmb.AddItem "---select---"
               cmb.ItemData(0) = -1
               ''cmb.ListIndex = 0
               If rst.RecordCount > 0 Then
                  rst.MoveFirst
                  While Not rst.EOF
                         cmb.AddItem rst.Fields(1)
                         cmb.ItemData(i) = rst.Fields(0)
                         ''MsgBox cmb.ItemData(i)
                         i = i + 1
                         rst.MoveNext
                  Wend
               End If
End Sub
Sub cmb_populate2(cmb As ComboBox, qry As String)
               Dim rst As New ADODB.Recordset
               Dim i As Integer
               i = 1
               rst.Open qry, cn
               cmb.AddItem "---select---"
               cmb.ItemData(0) = -1
               ''cmb.ListIndex = 0
               If rst.RecordCount > 0 Then
                  rst.MoveFirst
                  While Not rst.EOF
                         cmb.AddItem rst.Fields(2)
                         cmb.ItemData(i) = rst.Fields(1)
                         ''MsgBox cmb.ItemData(i)
                         i = i + 1
                         rst.MoveNext
                  Wend
               End If
End Sub

Private Sub command(b1 As Boolean, b2 As Boolean, b3 As Boolean, b4 As Boolean, b5 As Boolean)
Command1(0).Enabled = b1
Command1(1).Enabled = b2
Command1(2).Enabled = b3
Command1(3).Enabled = b4
Command1(4).Enabled = b5
End Sub


Private Sub command1_Click(Index As Integer)
 Dim rst_issue As ADODB.Recordset
 Set rst_issue = New ADODB.Recordset
 Select Case Index
  Case 0
        Call command(False, True, False, True, False)
        Combo1.Text = ""
        Combo2.Text = ""
        Text1.Text = ""
        Text2.Text = ""
        
        id = next_id("select max(issue_id) from issue")
        u_d = False
  Case 1
  If u_d = False Then
    If Combo1.Text = "" Or Combo2.Text = "" Or Text1.Text = "" Or Text2.Text = "" Then
     MsgBox "Enter Some Detail"
    Else
     rst_copy.Open "insert into copy values (" & Combo1.ItemData(Combo1.ListIndex) & ", " & no & " ,' " & Combo2.Text & " ')", cn
     MsgBox " New record stored successfully"
     Call reset
     Call command(True, False, False, False, True)
    End If
    Set rst_copy = Nothing
    u_d = True
  Else
    If Combo1.Text = "" Or Combo2.Text = "" Then
     MsgBox "Enter Some Detail"
    Else
     rst_copy.Open "update copy set book_id= " & Combo1.ItemData(Combo1.ListIndex) & " ,available=' " & Combo2.Text & " ',  where acc_no=" & seid & "", cn
     MsgBox "record updated successfully"
     Call reset
     Call command(True, False, False, False, True)
    End If
    Set rst_copy = Nothing
  End If
  Case 2
     res = MsgBox("Do you want to delete the record ?", vbYesNo)
     If res = 7 Then
        Exit Sub
     End If
     If Combo1.ItemData(Combo1.ListIndex) <> "" And Combo2.Text <> "" Then
      rst_copy.Open "delete from copy where acc_no=" & seid & "", cn
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
 Combo1.ListIndex = 0
 Combo2.ListIndex = 0
 Text1.Text = ""
 Text2.Text = ""
End Sub


