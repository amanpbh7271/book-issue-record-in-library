VERSION 5.00
Begin VB.Form member 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Member Entry "
   ClientHeight    =   7035
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12690
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   7035
   ScaleWidth      =   12690
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   14
      Top             =   3480
      Width           =   3015
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2520
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   1
      Top             =   1680
      Width           =   855
   End
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
      Index           =   4
      Left            =   9840
      TabIndex        =   11
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
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
      Index           =   3
      Left            =   7560
      TabIndex        =   10
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   5400
      TabIndex        =   9
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SAVE"
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
      Index           =   1
      Left            =   3240
      TabIndex        =   8
      Top             =   6360
      Width           =   1575
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
      Left            =   960
      TabIndex        =   0
      Top             =   6360
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   7
      Top             =   4680
      Width           =   3135
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
      Height          =   495
      Left            =   4200
      TabIndex        =   6
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "CLASS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   13
      Top             =   3480
      Width           =   3255
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "MEMBER ENTRY FORM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1320
      TabIndex        =   5
      Top             =   240
      Width           =   8415
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ADDRESS"
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
      Left            =   360
      TabIndex        =   4
      Top             =   4680
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "MEMBER TYPE"
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
      Left            =   360
      TabIndex        =   3
      Top             =   2520
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "MEMBER NAME"
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
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   3255
   End
End
Attribute VB_Name = "member"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rst_member As ADODB.Recordset
Public u_d As Boolean
Public id As Integer
Public seid As String


Sub command(b1 As Boolean, b2 As Boolean, b3 As Boolean, b4 As Boolean, b5 As Boolean)
    Command1(0).Enabled = b1
    Command1(1).Enabled = b2
    Command1(2).Enabled = b3
    Command1(3).Enabled = b4
    Command1(4).Enabled = b5
End Sub

Private Sub Combo1_Click()
If Combo1.Text = " teacher " Then
Text2.Enabled = False
Text2.Text = ""
Else
Text2.Enabled = True
End If
End Sub

Private Sub Command2_Click()
seid = ""
u_d = True
frmmember.Show 1
 If seid <> "" Then
        Text1.Enabled = True
        Text2.Enabled = True
         Combo1.Enabled = True
         Text4.Enabled = True
         
 Dim rst_member As ADODB.Recordset
 Set rst_member = New ADODB.Recordset
 rst_member.Open "select * from member where memb_id='" & seid & "'", cn
 Text1.Text = rst_member.Fields(1)
 Combo1.Enabled = True
'' Combo1.ListIndex = 0
Call cmb_populate(Combo1, "select * from memb_type")

 Combo1.ListIndex = cmb_search(Combo1, rst_member.Fields(2))
 If rst_member.Fields(4) <> "" Then
 Text2.Text = rst_member.Fields(4)
 End If
 
 
 
 
 
 Text4.Text = rst_member.Fields(3)
 Set rst_publisher = Nothing
 Call command(False, True, True, True, False)
 End If
 
End Sub



Private Sub Form_Load()
u_d = False

Call command(True, False, False, False, True)
'Call cmb_populate(Combo1, "select * from memb_type")
'Combo1.ListIndex = 0
Text1.Enabled = False
Text2.Enabled = False
Combo1.Enabled = False
Text4.Enabled = False
End Sub

Sub reset()
Text1.Text = ""
Text4.Text = ""
Text2.Text = ""
Combo1.Clear
Text1.Enabled = False
Text2.Enabled = False
Combo1.Enabled = False
Text4.Enabled = False

End Sub


Private Sub Command1_Click(Index As Integer)
 Dim rst_member As ADODB.Recordset
 Set rst_member = New ADODB.Recordset
 
 Select Case Index
  Case 0
        Call command(False, True, False, True, False)
        Text1.Text = ""
        Text4.Text = ""
        Text2.Text = ""
        Combo1.Clear
        Text1.Enabled = True
        Text2.Enabled = True
         Combo1.Enabled = True
         Text4.Enabled = True
        
        id = next_id("select max(memb_id) from member")
        u_d = False
        Call cmb_populate(Combo1, "select * from memb_type")
        Combo1.ListIndex = 0
  
  Case 1
  If u_d = False Then
  If Combo1.Text = "teacher" Then
    If Text1.Text = "" Or Combo1.Text = "" Or Text4.Text = "" Or Combo1.Text = "---select---" Then
     MsgBox "Enter All Details"
    Else
     rst_member.Open "insert into member values (" & id & ",'" & Text1.Text & "', " & Combo1.ItemData(Combo1.ListIndex) & " ,'" & Text4.Text & "','" & Text2.Text & "' )", cn
     MsgBox " New record stored successfully"
     Call reset
     Call command(True, False, False, False, True)
    End If
    Set rst_member = Nothing
    
    
    Else
     
     If Text1.Text = "" Or Combo1.Text = "" Or Text4.Text = "" Or Text2.Text = "" Or Combo1.Text = "---select---" Then
     MsgBox "Enter All Details"
    Else
     rst_member.Open "insert into member values (" & id & ",'" & Text1.Text & "', " & Combo1.ItemData(Combo1.ListIndex) & " ,'" & Text4.Text & "','" & Text2.Text & "' )", cn
     MsgBox " New record stored successfully"
     Call reset
     Call command(True, False, False, False, True)
    End If
    Set rst_member = Nothing
  
     
  End If
     
     
     
     
  Else
    If Combo1.Text = "teacher" Then
     If Text1.Text = "" Or Combo1.Text = "" Or Text4.Text = "" Then
     MsgBox "Enter Some Detail"
     Else
    '' Debug.Print "update member set name=' " & Text1.Text & " ',memb_type=' " & Text2.Text & " ',exp_date='" & Text3.Text & "',address='" & Text4.Text & "' where memb_id='" & seid & "'"
     rst_member.Open "update member set name=' " & Text1.Text & " ',membt_id= " & Combo1.ItemData(Combo1.ListIndex) & " ,  address=' " & Text4.Text & "',class='" & Text2.Text & "' where memb_id='" & seid & "'", cn
     MsgBox "record updated successfully"
     Call reset
     Call command(True, False, False, False, True)
    End If
    Set rst_member = Nothing
      Else
         
         
    If Text1.Text = "" Or Combo1.Text = "" Or Text4.Text = "" Or Text2.Text = "" Then
     MsgBox "Enter Some Detail"
     Else
    '' Debug.Print "update member set name=' " & Text1.Text & " ',memb_type=' " & Text2.Text & " ',exp_date='" & Text3.Text & "',address='" & Text4.Text & "' where memb_id='" & seid & "'"
     rst_member.Open "update member set name=' " & Text1.Text & " ',membt_id= " & Combo1.ItemData(Combo1.ListIndex) & " ,  address=' " & Text4.Text & "',class='" & Text2.Text & "' where memb_id='" & seid & "'", cn
     MsgBox "record updated successfully"
     Call reset
     Call command(True, False, False, False, True)
    End If
    Set rst_member = Nothing
  
     End If
  End If
  Case 2
     res = MsgBox("Do you want to delete the record ?", vbYesNo)
     If res = 7 Then
        Exit Sub
     End If
     If Text1.Text <> "" And Combo1.Text <> "" And Text2.Text <> "" And Text4.Text <> "" Then
      rst_member.Open "delete from member where memb_id='" & seid & "'", cn
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

Private Sub Text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 And Text1.Text <> "" Then
Combo1.SetFocus
End If
Select Case KeyAscii
 Case 65 To 122
 Case 13, 8, 32
 Case Else
 KeyAscii = 0
 End Select
 

End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text3.Text <> "" Then
Text4.SetFocus
End If
Select Case KeyAscii
 Case 47, 8
 Case 48 To 57
 Case Else
 KeyAscii = 0
 End Select
 

End Sub


Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text4.Text <> "" Then
Command1(1).SetFocus
End If
Select Case KeyAscii
 Case 65 To 122
 Case 13, 8, 32
 Case Else
 KeyAscii = 0
 End Select
 
End Sub

