VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form returnbook 
   BackColor       =   &H00C0C0FF&
   Caption         =   "returnbbok"
   ClientHeight    =   6645
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   17910
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   6645
   ScaleWidth      =   17910
   Begin VB.CommandButton Command3 
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
      Height          =   615
      Left            =   3600
      TabIndex        =   11
      Top             =   4200
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
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
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1200
      Width           =   2295
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3735
      Left            =   10080
      TabIndex        =   8
      Top             =   600
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   6588
      _Version        =   393216
      Rows            =   1
      Cols            =   4
      PictureType     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6480
      TabIndex        =   7
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
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
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
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
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2160
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "RETURN"
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
      Left            =   600
      TabIndex        =   2
      Top             =   4200
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
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
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "MEMB ID"
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
      Left            =   480
      TabIndex        =   9
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "CLASS"
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
      Left            =   480
      TabIndex        =   5
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "MEMBER  TYPE"
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
      Left            =   480
      TabIndex        =   3
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "MEMBER NAME"
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
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   2895
   End
End
Attribute VB_Name = "returnbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim rst_member As New ADODB.Recordset
 Dim rst_memb_type As New ADODB.Recordset
  Dim rst_memb As New ADODB.Recordset
  Dim rst_issue As New ADODB.Recordset
  Dim rst_copy As New ADODB.Recordset
  Dim rst_bissue As New ADODB.Recordset
  Public aid As Integer
  Dim id As Long
 Public seid As String
 Dim i As Integer
 Dim j As Integer
 Const strChecked = "þ"
 Const strUnChecked = "q"
 Public d As Date
 Public r As String



Private Sub Command1_Click()
d = Date$
a = 1
b = 3
c = 0
l = 0
If MSFlexGrid1.Rows <= 1 Then

MsgBox "there have no book"
 Else
   For m = 1 To MSFlexGrid1.Rows - 1
   If MSFlexGrid1.TextMatrix(m, 3) = strChecked Then
   l = l + 1
    End If
   Next

   For iii = 1 To MSFlexGrid1.Rows - 1
   If MSFlexGrid1.TextMatrix(iii, 3) = strChecked Then
   rst_copy.Open "update copy set available =  " & True & " where acc_no = " & MSFlexGrid1.TextMatrix(iii, 0) & " ", cn
   rst_bissue.Open "update issue set return_date= '" & d & "' where acc_no = " & MSFlexGrid1.TextMatrix(iii, 0) & " ", cn
   a = a + 1
   k = -1
   End If
    Next
     If l > 0 Then
     MsgBox "book is returned succesfully................."
     Else
     MsgBox "plz select book..."
     End If
     Set rst_copy = Nothing
     Set rst_bissue = Nothing
 End If
 
j = MSFlexGrid1.Rows - 1
While im <= j
 If MSFlexGrid1.Rows > 1 Then
 If MSFlexGrid1.TextMatrix(im, 3) = strChecked Then
   If MSFlexGrid1.Rows > 1 Then
       If MSFlexGrid1.Rows - MSFlexGrid1.FixedRows = 1 Then
         MSFlexGrid1.Rows = MSFlexGrid1.Rows - 1
         im = im - 1
        Else
         MSFlexGrid1.RemoveItem (im)
         im = im - 1
       End If
      i = 1
   End If
   
 End If
 End If
im = im + 1
j = MSFlexGrid1.Rows - 1
Wend
 
 
 
 
 
  If MSFlexGrid1.Rows > 1 Then
  MSFlexGrid1.Enabled = True
 Else
 MSFlexGrid1.Enabled = False
 End If


End Sub

  Private Sub Command2_Click()
  MSFlexGrid1.Enabled = True
  i = 1
  seid = ""
  MSFlexGrid1.Clear
  Text2.Text = ""
  Text1.Text = ""
  Text4.Text = ""
  Text5.Text = ""
i = 1
j = 0
MSFlexGrid1.TextMatrix(0, 0) = "Acc No"
MSFlexGrid1.TextMatrix(0, 1) = "Book Name"
MSFlexGrid1.TextMatrix(0, 2) = "Issue Date"
MSFlexGrid1.TextMatrix(0, 3) = " Select "
MSFlexGrid1.Rows = 1
  
  searchmemb1.Show 1
  If r <> "" Then
  Dim rst_member As ADODB.Recordset
  Set rst_member = New ADODB.Recordset
  Set rst_memb_type = New ADODB.Recordset
  rst_member.Open "select * from member where memb_id= '" & r & "' ", cn
  ''Debug.Print " select * from member, memb_type where member.membt_id=memb_type.membt_id and memb_id='" & seid & "' "
  rst_memb.Open " select * from member, memb_type where member.membt_id=memb_type.membt_id and memb_id='" & r & "' ", cn
  Text1.Text = rst_memb(1)
  Text2.Text = rst_memb(0)
  Text4.Text = rst_memb(6)
  
  
  If Not IsNull(rst_memb.Fields(4)) Then

  Text5.Text = rst_memb(4)
  End If
''  Debug.Print "select * from book,issue,copy where copy.acc_no=issue.acc_no and book.book_id=copy.book_id and issue.memb_id='" & seid & "' "
 
 
  rst_issue.Open "select * from book,issue,copy where copy.acc_no=issue.acc_no and book.book_id=copy.book_id and issue.memb_id='" & r & "' ", cn

  
  If MSFlexGrid1.Rows > 0 Then
  For ii = 0 To rst_issue.RecordCount - 1
  If rst_issue(11) = False Then
  If IsNull(rst_issue(8)) Then
  MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
  MSFlexGrid1.TextMatrix(i, j) = rst_issue(10)
  j = j + 1
  MSFlexGrid1.TextMatrix(i, j) = rst_issue(3)
  j = j + 1
  MSFlexGrid1.TextMatrix(i, j) = rst_issue(7)
  j = 0
  i = i + 1

  
  End If
  End If
  rst_issue.MoveNext
  Next
   
  For X = 1 To MSFlexGrid1.Rows - 1
  For Y = 3 To MSFlexGrid1.Cols - 1
  MSFlexGrid1.Row = X
  MSFlexGrid1.Col = Y
  
                   MSFlexGrid1.Text = strUnChecked
                   MSFlexGrid1.CellFontName = "Wingdings"
                    MSFlexGrid1.CellFontSize = 14
                    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
Next
Next
   
 End If
   
Set rst_member = Nothing
 Set rst_issue = Nothing
  Set rst_memb = Nothing
 
  r = ""
        
 
 End If
 If r <> "" Then
 If MSFlexGrid1.Rows <= 1 Then
 MsgBox "this member have no issue books"
 End If
 End If
 If MSFlexGrid1.Rows > 1 Then
  MSFlexGrid1.Enabled = True
 Else
 MSFlexGrid1.Enabled = False
 End If
 End Sub


Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
i = 1
j = 0
MSFlexGrid1.TextMatrix(0, 0) = "Acc No"
MSFlexGrid1.TextMatrix(0, 1) = "Book Name"
MSFlexGrid1.TextMatrix(0, 2) = "Issue Date"
MSFlexGrid1.TextMatrix(0, 3) = " Select "
MSFlexGrid1.Rows = 1
MSFlexGrid1.Enabled = False
End Sub

Private Sub TriggerCheckbox(iRow As Integer, icol As Integer)
        If icol <> 3 Then
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
 


