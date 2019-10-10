VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmbook 
   BackColor       =   &H00C0E0FF&
   Caption         =   "frmbook"
   ClientHeight    =   6030
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6570
   LinkTopic       =   "Form2"
   ScaleHeight     =   6030
   ScaleWidth      =   6570
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1335
      Left            =   1320
      TabIndex        =   2
      Top             =   2400
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   2355
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1440
      TabIndex        =   0
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   " BOOK NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   600
      Width           =   3375
   End
End
Attribute VB_Name = "frmbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public r_book As New ADODB.Recordset
Dim bid() As String

Private Sub DataGrid1_Click()
If DataGrid1.Row >= 0 Then
Y = DataGrid1.Columns(0)
books.seid = Y
End If
Unload Me

End Sub

Private Sub text1_Change()

 If Text1.Text <> "" Then
'' Debug.Print " select author_no, author_name from author where author_name like '%" & Text1 & "%'"
 r_book.Open "select book_id as Book_Id, title as Book_Name from book where title like '%" & Text1 & "%' ", cn
Set DataGrid1.DataSource = r_book
Set r_book = Nothing
End If
End Sub



