VERSION 5.00
Begin VB.Form searchmemb 
   Caption         =   "searchmemb"
   ClientHeight    =   3900
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form3"
   ScaleHeight     =   3900
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   960
      TabIndex        =   1
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "MEMBER NAME"
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
      Left            =   960
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "searchmemb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim eid() As String
Public r_member As New ADODB.Recordset



Private Sub list1_Click()
issue_book.seid = eid(List1.ListIndex)
returnbook.seid = eid(List1.ListIndex)
Unload Me
End Sub

Private Sub text1_Change()
List1.Clear
 If Text1.Text <> "" Then
 r_member.Open "select memb_id, name from member where name like '%" & Text1 & "%' ", cn
  If r_member.RecordCount > 0 Then
 ReDim eid(r_member.RecordCount - 1)
 For i = 1 To r_member.RecordCount
  List1.AddItem r_member.Fields(1)
  eid(i - 1) = r_member.Fields(0)
  r_member.MoveNext
 Next
   Set r_member = Nothing
  End If
Set r_member = Nothing
End If
End Sub



