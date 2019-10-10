VERSION 5.00
Begin VB.Form frmcopy 
   Caption         =   "frmcopy"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   1440
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "frmcopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim eid() As String
Public r_copy As New ADODB.Recordset
Public r_book As New ADODB.Recordset


Private Sub list1_Click()
Copy.seid = eid(List1.ListIndex)
Unload Me
End Sub

Private Sub text1_Change()
 List1.Clear
 If Text1.Text <> "" Then
 r_book.Open "select book_id, title,author name from book,author where title like '%" & Text1 & "%' ", cn
  If r_copy.RecordCount > 0 Then
 ReDim eid(r_copy.RecordCount - 1)
 For i = 1 To r_copy.RecordCount
  List1.AddItem r_copy.Fields(1)
  eid(i - 1) = r_copy.Fields(0)
  r_copy.MoveNext
 Next
   Set r_copy = Nothing
  End If
Set r_copy = Nothing
End If
End Sub


