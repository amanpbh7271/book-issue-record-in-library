VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form issue_book 
   BackColor       =   &H00C0C0FF&
   Caption         =   "issue_book"
   ClientHeight    =   8730
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   14220
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   8730
   ScaleWidth      =   14220
   Begin VB.CommandButton Command6 
      Caption         =   "EXIT"
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
      Left            =   2280
      TabIndex        =   22
      Top             =   5160
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2175
      Left            =   4800
      TabIndex        =   21
      Top             =   5160
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   3836
      _Version        =   393216
      Cols            =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Remove"
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
      Left            =   12120
      TabIndex        =   20
      Top             =   5160
      Width           =   1695
   End
   Begin VB.ComboBox Combo3 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "issue_book.frx":0000
      Left            =   3120
      List            =   "issue_book.frx":0002
      Locked          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   17
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ISSUE"
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
      Left            =   600
      TabIndex        =   16
      Top             =   5160
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3120
      Locked          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12360
      TabIndex        =   14
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12240
      TabIndex        =   13
      Top             =   960
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "issue_book.frx":0004
      Left            =   10560
      List            =   "issue_book.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   3000
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   10560
      TabIndex        =   11
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
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
      Left            =   10560
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   840
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "BOOKS DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   8040
      TabIndex        =   6
      Top             =   360
      Width           =   5775
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Title"
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
         TabIndex        =   9
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Accession No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   8
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Author Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   7
         Top             =   1320
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   5
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
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
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "MEMBER DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   5895
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Memb ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   18
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label label1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Member Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Member type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   2
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Class"
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
         TabIndex        =   1
         Top             =   2880
         Width           =   1215
      End
   End
End
Attribute VB_Name = "issue_book"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim rst_member As New ADODB.Recordset
 Dim rst_issue As New ADODB.Recordset
 Dim rst_copy As New ADODB.Recordset
Public d As Date
Public r As String
 Dim id As Long
Public seid As String
Public i_id As Integer
Dim i As Integer
Dim s As Integer
 Const strChecked = "þ"
 Const strUnChecked = "q"




Private Sub Command1_Click()
Text3.Text = ""
Text2.Text = ""
Call cmb_populate(Combo2, "select * from memb_type")
Call cmb_populate(Combo3, "select * from member")

searchmemb1.Show 1
 If r <> "" Then
 Dim rst_member As ADODB.Recordset
 Set rst_member = New ADODB.Recordset
 rst_member.Open "select * from member where memb_id= '" & r & "' ", cn

 Combo3.ListIndex = cmb_search(Combo3, rst_member.Fields(0))
   Text3.Text = rst_member.Fields(0)
 If Not IsNull(rst_member.Fields(4)) Then
 
 Text2.Text = rst_member.Fields(4)
 End If
 Combo2.ListIndex = cmb_search(Combo2, rst_member.Fields(2))
  Set rst_member = Nothing
End If
If Combo3.Text <> "" Then
  Command2.Enabled = True
  End If
End Sub

Private Sub Command2_Click()
Command3.Enabled = True
Text4.Text = ""
List1.Clear
Combo1.Clear
seid = ""
u_d = True
searchbook.Show 1
 If seid <> "" Then
 Dim rst_book As ADODB.Recordset
 Set rst_book = New ADODB.Recordset
 Dim rst_copy As ADODB.Recordset
 Set rst_copy = New ADODB.Recordset
 
 rst_book.Open "select * from book where book_id=" & seid & "", cn
 Text4.Text = rst_book.Fields(3)
Call cmb_populate1(Combo1, "select acc_no from copy where copy.book_id=" & seid & "and available=true")
Combo1.ListIndex = 0
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

 
 Set rst_book = Nothing
 End If

End Sub

Private Sub Command3_Click()
If Combo1.Text <> "" And Combo1.Text <> "---select---" Then
If MSFlexGrid1.Rows > 1 Then
    For ii = 1 To MSFlexGrid1.Rows - 1
        If Combo1.Text = MSFlexGrid1.TextMatrix(ii, 1) Then
            MsgBox ("Äccession Number already exists....")
            Exit Sub
        End If
    Next
End If


MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1


MSFlexGrid1.TextMatrix(i, 0) = s

MSFlexGrid1.TextMatrix(i, 1) = Combo1.Text
s = s + 1
MSFlexGrid1.Row = i
  MSFlexGrid1.Col = 2
 MSFlexGrid1.Text = strUnChecked
                   MSFlexGrid1.CellFontName = "Wingdings"
                    MSFlexGrid1.CellFontSize = 14
                    MSFlexGrid1.CellAlignment = flexAlignCenterCenter

i = i + 1

End If
   
If MSFlexGrid1.Rows > 1 Then
MSFlexGrid1.Enabled = True
Else
 MSFlexGrid1.Enabled = False
 End If



End Sub



Private Sub Command4_Click()
If MSFlexGrid1.Rows > 1 Then
If Text3.Text <> "" Then
f = False
i_id = next_id("select max(issue_id)from issue")
b = 202

e = Null
d = Date$
Set rst_issue = New ADODB.Recordset
If MSFlexGrid1.Rows > 1 Then
For i = 1 To MSFlexGrid1.Rows - 1
 ''Debug.Print "insert into issue values(" & i_id & ",'" & Combo3.ItemData(Combo3.ListIndex) & "'," & MSFlexGrid1.TextMatrix(i, 1) & ",'" & d & ",null)"
 rst_issue.Open "insert into issue values(" & i_id & ",'" & Combo3.ItemData(Combo3.ListIndex) & "'," & MSFlexGrid1.TextMatrix(i, 1) & ",'" & d & "',null)", cn
 
Next
End If

If MSFlexGrid1.Rows > 1 Then
For i = 1 To MSFlexGrid1.Rows - 1
''Debug.Print "update copy set available = '" & f & "' where acc_no = " & CInt(MSFlexGrid1.TextMatrix(i, 1)) & ""
rst_copy.Open "update copy set available = " & f & " where acc_no = " & CInt(MSFlexGrid1.TextMatrix(i, 1)) & "", cn



Next
End If
MsgBox " book susseccfully issued ......"
MSFlexGrid1.Rows = 1
Combo2.Clear
Text2.Text = ""
Text3.Text = ""
Combo3.Clear
Text4.Text = ""
Combo1.Clear
List1.Clear
Command2.Enabled = False
Command3.Enabled = False
i = 1
End If
End If
If MSFlexGrid1.Rows > 1 Then
MSFlexGrid1.Enabled = True
Else
MSFlexGrid1.Enabled = False
End If

End Sub

Private Sub Command5_Click()
a = 1
b = 2
j = MSFlexGrid1.Rows - 1
While iii <= j
 If MSFlexGrid1.Rows > 1 Then
 If MSFlexGrid1.TextMatrix(iii, 2) = strChecked Then
   If MSFlexGrid1.Rows > 1 Then
       If MSFlexGrid1.Rows - MSFlexGrid1.FixedRows = 1 Then
         MSFlexGrid1.Rows = MSFlexGrid1.Rows - 1
         iii = iii - 1
        Else
         MSFlexGrid1.RemoveItem (iii)
         iii = iii - 1
       End If
      i = 1
   End If
   
 End If
 End If
iii = iii + 1
j = MSFlexGrid1.Rows - 1
Wend

If MSFlexGrid1.Rows > 0 Then
 
i = MSFlexGrid1.Rows + 1 - MSFlexGrid1.FixedRows
s = i
 
End If

For iii = 1 To MSFlexGrid1.Rows - 1
  MSFlexGrid1.TextMatrix(iii, 0) = iii
Next
If MSFlexGrid1.Rows > 1 Then
MSFlexGrid1.Enabled = True
Else
MSFlexGrid1.Enabled = False
End If
End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Form_Load()
Command2.Enabled = False
Command3.Enabled = False
i = 1

s = 1
Call cmb_populate(Combo2, "select * from memb_type")
MSFlexGrid1.TextMatrix(0, 0) = "Sr. No"
MSFlexGrid1.TextMatrix(0, 1) = "acc_no"
MSFlexGrid1.Rows = 1
Call cmb_populate(Combo3, "select * from member")
MSFlexGrid1.Enabled = False
End Sub


Sub cmb_populate1(cmb As ComboBox, qry As String)
               Dim rst As New ADODB.Recordset
               Dim i As Integer
               i = 1
               rst.Open qry, cn
               cmb.Clear
               cmb.AddItem "---select---"
               cmb.ItemData(0) = -1
               ''cmb.ListIndex = 0
               If rst.RecordCount > 0 Then
                  rst.MoveFirst
                  While Not rst.EOF
                         
                        If rst.Fields.Count = 2 Then
                        cmb.AddItem rst.Fields(1)
                         cmb.ItemData(i) = rst.Fields(0)
                        Else
                         cmb.AddItem rst.Fields(0)
                         cmb.ItemData(i) = rst.Fields(0)
                          
                        End If
                        ''MsgBox cmb.ItemData(i)
                         i = i + 1
                         rst.MoveNext
                  Wend
               End If
End Sub

Private Sub TriggerCheckbox(iRow As Integer, icol As Integer)
  
        If icol <> 2 Then
        Exit Sub
        End If
        With MSFlexGrid1
                       
                        If .TextMatrix(iRow, icol) = strUnChecked Then
                .TextMatrix(iRow, icol) = strChecked
            Else
                .TextMatrix(iRow, icol) = strUnChecked
         End If
            
        End With
        
End Sub
 
Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 32 Then 'Enter/Space
        With MSFlexGrid1
            Call TriggerCheckbox(.Row, .Col)
        End With
    End If
End Sub
 
Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        With MSFlexGrid1
            If .MouseRow <> 0 Then
                Call TriggerCheckbox(.MouseRow, .MouseCol)
            End If
        End With
    End If
End Sub
 

