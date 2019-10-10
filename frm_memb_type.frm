VERSION 5.00
Begin VB.Form frm_memb_type 
   Caption         =   "frm_memb_type"
   ClientHeight    =   3960
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7590
   LinkTopic       =   "Form3"
   ScaleHeight     =   3960
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   1560
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   1560
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "frm_memb_type"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim eid() As String
Public r_memb_type As New ADODB.Recordset




Private Sub List1_Click()
memb_type.seid = eid(List1.ListIndex)
Unload Me
End Sub

Private Sub Text1_Change()
 List1.Clear
 If Text1.Text <> "" Then
 r_memb_type.Open "select memb_id, m_type from memb_type where m_type like '%" & Text1 & "%' ", cn
  If r_memb_type.RecordCount > 0 Then
 ReDim eid(r_memb_type.RecordCount - 1)
 For i = 1 To r_memb_type.RecordCount
  List1.AddItem r_memb_type.Fields(1)
  eid(i - 1) = r_memb_type.Fields(0)
  r_memb_type.MoveNext
 Next
   Set r_memb_type = Nothing
  End If
Set r_memb_type = Nothing
End If
End Sub


