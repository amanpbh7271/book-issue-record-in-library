VERSION 5.00
Begin VB.Form books 
   BackColor       =   &H00C0C0FF&
   Caption         =   "BOOK ENTRY FORM"
   ClientHeight    =   8190
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13800
   LinkTopic       =   "Form6"
   MDIChild        =   -1  'True
   ScaleHeight     =   8190
   ScaleWidth      =   13800
   Begin VB.CommandButton Command3 
      Caption         =   "REMOVE"
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
      Left            =   10440
      TabIndex        =   16
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   15
      Top             =   1800
      Width           =   735
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      ItemData        =   "book.frx":0000
      Left            =   8160
      List            =   "book.frx":0002
      TabIndex        =   14
      Top             =   2520
      Width           =   1935
   End
   Begin VB.ComboBox combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5160
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   2760
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5160
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT"
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
      Index           =   4
      Left            =   8880
      TabIndex        =   10
      Top             =   6600
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
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
      Index           =   3
      Left            =   6840
      TabIndex        =   9
      Top             =   6600
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
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
      Index           =   2
      Left            =   4680
      TabIndex        =   8
      Top             =   6600
      Width           =   1815
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
      Height          =   495
      Index           =   1
      Left            =   2520
      TabIndex        =   7
      Top             =   6600
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NEW"
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
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   6600
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   5160
      TabIndex        =   5
      Top             =   3720
      Width           =   2295
   End
   Begin VB.TextBox text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   5160
      TabIndex        =   2
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label AUTHOR 
      BackColor       =   &H00C0C0FF&
      Caption         =   "AUTHOR"
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
      Left            =   1800
      TabIndex        =   12
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "BOOK ENTRY FORM"
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
      Left            =   3120
      TabIndex        =   6
      Top             =   480
      Width           =   6975
   End
   Begin VB.Label lbltitle 
      BackColor       =   &H00C0C0FF&
      Caption         =   "BOOK NAME"
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
      Left            =   1800
      TabIndex        =   4
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Label lblprice 
      BackColor       =   &H00C0C0FF&
      Caption         =   "PRICE"
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
      Left            =   1800
      TabIndex        =   3
      Top             =   3840
      Width           =   2775
   End
   Begin VB.Label lblname 
      BackColor       =   &H00C0C0FF&
      Caption         =   "PUBLISHER NAME"
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
      Left            =   1800
      TabIndex        =   1
      Top             =   5040
      Width           =   3015
   End
End
Attribute VB_Name = "books"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst_author As New ADODB.Recordset
Dim rst_book_author As New ADODB.Recordset
Dim aud As Integer
Dim b_id As Integer
Dim bid As String
Public seid As Integer
Dim rst_book As New ADODB.Recordset
Dim u_d As Boolean
Dim re As Integer
Dim book_a As New ADODB.Recordset





Sub clear_combo(cmb As ComboBox)
    Dim X As Long
    i = 0
    X = cmb.ListCount
    While i < X
        cmb.RemoveItem i
        X = cmb.ListCount
        
    Wend
End Sub




Private Sub Combo2_Change()
Call clear_combo(Combo2)
 If Combo2.Text <> "" Then
 Call cmb_populate(Combo2, "select * from author")
  End If
End Sub

Private Sub Combo2_Click()
 Dim i As Integer
 If Combo2.Text <> "---select---" Then
  For i = 0 To List1.ListCount - 1
  If List1.ItemData(i) = Combo2.ItemData(Combo2.ListIndex) Then
  MsgBox "this author name is already added"
  Exit Sub
  End If
  Next
 
 
 List1.AddItem Combo2.List(Combo2.ListIndex)
 i = List1.ListCount - 1
 List1.ItemData(i) = Combo2.ItemData(Combo2.ListIndex)

 End If

 
End Sub

Private Sub Command2_Click()
seid = 0
Combo1.Clear
Combo2.Clear
List1.Clear
u_d = True
frmbook.Show 1
''MsgBox "record"
 If seid <> 0 Then
 Text2.Enabled = True
Text3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
 Dim rst_book As ADODB.Recordset
 Set rst_book = New ADODB.Recordset
 rst_book.Open "select * from book where book_id=" & seid & "", cn
'' MsgBox rst_book.RecordCount
''Combo1.ListIndex = 0
Call cmb_populate(Combo1, "select * from publisher")
 Combo1.ListIndex = cmb_search(Combo1, rst_book.Fields(1))
 Call cmb_populate(Combo2, "select * from author")
 Combo2.ListIndex = 0
 Text2.Text = rst_book.Fields(2)
 Text3.Text = rst_book.Fields(3)
''Combo2.ListIndex = cmb_search(Combo2, rst_author.Fields(1))
 Dim i As Integer
 i = 0
 Set rst_book = Nothing
 Set rst_book_author = New ADODB.Recordset
 rst_book_author.Open "select author.author_no,author.author_name from book,book_author ,author where book.book_id=" & seid & " and book_author.book_id= book.book_id and book_author.author_no=author.author_no", cn
  If rst_book_author.RecordCount > 0 Then
   For i = 0 To rst_book_author.RecordCount - 1
     List1.AddItem rst_book_author.Fields(1)
     List1.ItemData(i) = rst_book_author.Fields(0)
     rst_book_author.MoveNext
   Next
   Set rst_book_author = Nothing
  End If
 Set rst_book_author = Nothing
  u_d = True
 Call command_button(False, True, True, True, False)
 End If

End Sub

Private Sub Command3_Click()
If List1.ListCount > 0 Then
If List1.Selected(re) = True Then
List1.RemoveItem (re)
re = 0
End If
End If
End Sub

Private Sub Form_Load()

'Call cmb_populate(Combo1, "select * from publisher")
'Combo1.ListIndex = 0
' Call cmb_populate(Combo2, "select * from author")
' Combo2.ListIndex = 0
i = 0
u_d = True
Text2.Enabled = False
Text3.Enabled = False
Combo1.Enabled = False
Combo2.Enabled = False
Call command_button(True, False, False, False, True)

End Sub


Sub command_button(b1 As Boolean, b2 As Boolean, b3 As Boolean, b4 As Boolean, b5 As Boolean)
    Command1(0).Enabled = b1
    Command1(1).Enabled = b2
    Command1(2).Enabled = b3
    Command1(3).Enabled = b4
    Command1(4).Enabled = b5
End Sub
Private Sub Command1_Click(Index As Integer)
 
 Select Case Index
  Case 0
        Call command_button(False, True, False, True, False)
        Text3.Text = ""
       List1.Clear
      ' Combo2.Text = ""
       Text2.Text = ""
      ' Combo1.Text = ""
      Text2.Enabled = True
Text3.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True

        b_id = next_id("select max(book_id)from book")

        no = next_no("select max (author_no)from book_author")
        Call cmb_populate(Combo1, "select * from publisher")
        Combo1.ListIndex = 0
       Call cmb_populate(Combo2, "select * from author")
       Combo2.ListIndex = 0

        u_d = False
        
  Case 1
  If u_d = False Then
       If Combo1.Text = "" Or Combo1.Text = "---select---" Or Text2.Text = "" Or Text3.Text = "" Or List1.ListCount = 0 Then
       MsgBox "Enter All Details"
       Else
       Set rst_book = New ADODB.Recordset
       Set rst_book_author = New ADODB.Recordset
       rst_book.Open "insert into book values(" & b_id & "," & Combo1.ItemData(Combo1.ListIndex) & " ,'" & Text2.Text & "',' " & Text3.Text & "')", cn
         
        For j = 0 To List1.ListCount - 1
               rst_book_author.Open "insert into book_author values(" & b_id & ", " & List1.ItemData(j) & ")", cn
        
        Next
          
          MsgBox " New record stored successfully"
          
       Call reset
    
    
      Call command_button(True, False, False, False, True)
      
      End If
    Set rst_book = Nothing
    Set rst_book_author = Nothing
    u_d = True
  Else
    
    If Combo1.Text = "" Or Combo1.Text = "---select---" Or Text2.Text = "" Or Text3.Text = "" Or List1.ListCount = 0 Then
     MsgBox "Enter All Details"
    Else
     Set rst_book = New ADODB.Recordset
     Set rst_book_author = New ADODB.Recordset
     rst_book.Open "update book set pub_id=" & Combo1.ItemData(Combo1.ListIndex) & " ,price=' " & Text2.Text & "',title='" & Text3.Text & "' where book_id=" & seid & "", cn
     
      book_a.Open "delete from  book_author where book_id=" & seid & "", cn
      Set book_a = Nothing
      
      
     For j = 0 To List1.ListCount - 1
     rst_book_author.Open "insert into book_author values(" & seid & ", " & List1.ItemData(j) & ")", cn
          
     Next

     
     
     
     
     MsgBox "record updated successfully"
     Call reset
     Call command_button(True, False, False, False, True)
    End If
    Set rst_book = Nothing
    Set rst_book_author = Nothing
  End If
   
 
      
  Case 2
      
      res = MsgBox("Do you want to delete the record ?", vbYesNo)
     If res = 7 Then
        Exit Sub
     End If
     If Combo1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" And List1.ListCount <> 0 Then
      rst_book_author.Open "delete from book_author where book_id=" & seid & "", cn
      Set rst_book_author = Nothing
      rst_book_author.Open "delete from copy where book_id=" & seid & "", cn
      Set rst_book_author = Nothing
      rst_book.Open "delete from book where book_id=" & seid & "", cn
      Set rst_book = Nothing
      MsgBox "record deleted successfully"
      Call reset
      Set rst_book = Nothing
      Set rst_book_author = Nothing
     Else
      MsgBox "record changed! Deletion can not be done  ", vbCritical
     End If
      
      
  Case 3
  Call command_button(True, False, False, False, True)
    Call reset
  
  Case 4
  
   Unload Me
 End Select
  
 
End Sub




Private Sub List1_Click()
re = List1.ListIndex

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Text3.Text <> "" Then
Combo2.SetFocus
End If
Select Case KeyAscii
 Case 65 To 122
 Case 48 To 57
 Case 13, 8, 32
 Case Else
 KeyAscii = 0
 End Select
 

End Sub
Private Sub combo2_keypress(KeyAscii As Integer)
If KeyAscii = 13 And Combo2.Text <> "" Then
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
Combo1.SetFocus
End If
Select Case KeyAscii
Case 36
Case 48 To 57
Case 13, 8, 32
 Case Else
 KeyAscii = 0
 End Select
 

End Sub
Private Sub combo1_keypress(KeyAscii As Integer)
If KeyAscii = 13 And Combo1.Text <> "" Then
Command1(1).SetFocus
End If
Select Case KeyAscii
 Case 65 To 122
 Case 13, 8, 32
 Case Else
 KeyAscii = 0
 End Select
 

End Sub

Sub reset()
Text2.Text = ""
Combo1.Clear
Text3.Text = ""
Combo2.Clear
List1.Clear
Text2.Enabled = False
Text3.Enabled = False
Combo1.Enabled = False
Combo2.Enabled = False
End Sub

Private Sub Timer1_Timer()
If Label1.Visible = True Then
Label1.Visible = False
Else
Label1.Visible = True
End If

End Sub

