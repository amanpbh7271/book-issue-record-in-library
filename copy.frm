VERSION 5.00
Begin VB.Form book_accession 
   BackColor       =   &H00C0E0FF&
   Caption         =   "book_accession"
   ClientHeight    =   7485
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15555
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   7485
   ScaleWidth      =   15555
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   10
      Top             =   2520
      Width           =   1815
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   3720
      TabIndex        =   9
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6600
      TabIndex        =   8
      Top             =   2520
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4440
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1560
      Width           =   1695
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
      Left            =   8280
      TabIndex        =   5
      Top             =   5520
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CANCEL"
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
      Left            =   5520
      TabIndex        =   4
      Top             =   5520
      Width           =   2535
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
      Index           =   1
      Left            =   3120
      TabIndex        =   3
      Top             =   5520
      Width           =   2175
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
      TabIndex        =   2
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "BOOK ACCESSION FORM "
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
      Left            =   1200
      TabIndex        =   6
      Top             =   360
      Width           =   8535
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Total  Copies"
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
      Left            =   1320
      TabIndex        =   1
      Top             =   2520
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   " Book Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   1560
      Width           =   2775
   End
End
Attribute VB_Name = "book_accession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public u_d As Boolean
Dim no As Long
Public seid As String
Dim rst_book As New ADODB.Recordset
Dim rst_copy As New ADODB.Recordset
Dim midd() As String
Dim a As Integer
Dim s As String


Private Sub command(b1 As Boolean, b2 As Boolean, b3 As Boolean, b4 As Boolean, b5 As Boolean)
Command1(0).Enabled = b1
Command1(1).Enabled = b2
''Command1(2).Enabled = b3
Command1(3).Enabled = b4
Command1(4).Enabled = b5
End Sub

Private Sub Command2_Click()
If Combo1.Text <> "---select---" And Combo1.Text <> "" Then
If List1.ListCount > 0 Then
For i = 1 To Val(Text1.Text)
 a = a + 1
 List1.AddItem a
 Next

Else
Dim rst_c As New ADODB.Recordset
rst_c.Open "select max(acc_no) from copy", cn
a = rst_c.Fields(0)
For i = 1 To Val(Text1.Text)
 a = a + 1
 List1.AddItem a
 Next

End If
Else
 MsgBox "plz select book"

End If
End Sub

Private Sub Form_Load()


'Call cmb_populate1(Combo1, "select * from book")
'Combo1.ListIndex = 0
u_d = True
Combo1.Enabled = False
Text1.Enabled = False
Command2.Enabled = False
Call command(True, False, False, False, True)


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
                         cmb.AddItem rst.Fields(3)
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
                         ''cmb.ItemData(i) = rst.Fields(0)
                         ''MsgBox cmb.ItemData(i)
                         i = i + 1
                         rst.MoveNext
                  Wend
               End If
End Sub


Private Sub Command1_Click(Index As Integer)
 s = True
 Select Case Index
  Case 0
        Call command(False, True, False, True, False)
        ''Combo1.Text = ""
        Text1.Text = ""
        List1.Clear
        
       '' no = next_id("select max(acc_no) from copy")
        
        Combo1.Enabled = True
       Text1.Enabled = True
       Command2.Enabled = True
        
        u_d = False
        
Call cmb_populate1(Combo1, "select * from book")
Combo1.ListIndex = 0
  Case 1
  If u_d = False Then
    If Combo1.Text = "" Or List1.List(0) = "" Then
     MsgBox "Enter All Detail"
    Else
     For i = 0 To List1.ListCount - 1
      rst_copy.Open "insert into copy values (" & Combo1.ItemData(Combo1.ListIndex) & "," & List1.List(i) & "," & s & ")", cn
      Set rst_copy = Nothing
     Next
     MsgBox " New record stored successfully"
     Call reset
     Call command(True, False, False, False, True)
    End If
    Set rst_copy = Nothing
    u_d = False
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
 Text1.Text = ""
 List1.Clear
 Combo1.Clear
Combo1.Enabled = False
Text1.Enabled = False
Command2.Enabled = False
End Sub




Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text1.Text <> "" Then
Command2.SetFocus
End If
Select Case KeyAscii
 Case 48 To 57
 Case 13, 8, 32
 Case Else
 KeyAscii = 0
 End Select
 

End Sub
Private Sub combo1_keypress(KeyAscii As Integer)
If KeyAscii = 13 And Combo1.Text <> "" Then
Text1.SetFocus
End If
Select Case KeyAscii
 Case 65 To 122
 Case 13, 8, 32
 Case Else
 KeyAscii = 0
 End Select
 

End Sub
