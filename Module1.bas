Attribute VB_Name = "Module1"
Public cn As ADODB.Connection
Sub main()
Set cn = New ADODB.Connection
cn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=F:\aman\library\aman.accdb;Persist Security Info=False"
cn.CursorLocation = adUseClient
cn.Open
If cn.State Then
cn.BeginTrans
''MsgBox "connection established successfully"
cn.CommitTrans
End If
''login.Show
''Form1.Show
''publisher.Show
''member.Show
''b_author.Show
'' book_accession.Show
books.Show
''Form2.Show
''searchbook.Show
''MDIForm1.Show
 '' memb_type.Show
''issue_book.Show
''returnbook.Show
''changepassword.Show
  ''Form2.Show
''report.Show
  ''welcome.Show
  End Sub
Function next_id(qry As String) As Long
               Dim rst As New ADODB.Recordset
               rst.Open qry, cn
               If IsNull(rst.Fields(0)) Then
                   next_id = 1
               Else
                   next_id = rst.Fields(0) + 1
               End If
End Function

Function next_no(qry As String) As Long
               Dim rst As New ADODB.Recordset
               rst.Open qry, cn
               If IsNull(rst.Fields(0)) Then
                   next_no = 1
               Else
                   next_no = rst.Fields(0) + 1
               End If
End Function

Sub cmb_populate(cmb As ComboBox, qry As String)
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
                          
                        ''here
                        ''MsgBox cmb.ItemData(i)
                         i = i + 1
                         rst.MoveNext
                  Wend
               End If
End Sub

Function cmb_search(cmb As ComboBox, id As Long) As Long
 For i = 1 To cmb.ListCount - 1
   If cmb.ItemData(i) = id Then
    cmb_search = i
    Exit Function
   End If
 Next
End Function

