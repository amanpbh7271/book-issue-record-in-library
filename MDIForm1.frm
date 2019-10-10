VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00C0C0FF&
   Caption         =   "MDIForm1"
   ClientHeight    =   6840
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   17160
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   6000
      Top             =   2040
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   17100
      TabIndex        =   0
      Top             =   0
      Width           =   17160
      Begin VB.CommandButton Command1 
         Caption         =   "log out"
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
         Left            =   16080
         TabIndex        =   1
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   22.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   735
         Left            =   6600
         TabIndex        =   2
         Top             =   0
         Width           =   8175
      End
   End
   Begin VB.Menu pub 
      Caption         =   "PUBLISHER RECORD"
   End
   Begin VB.Menu memb 
      Caption         =   "MEMBER RECORD"
   End
   Begin VB.Menu BOOK 
      Caption         =   "BOOK  RECORD"
      Begin VB.Menu bb_author 
         Caption         =   "BOOK AUTHOR RECORD"
      End
      Begin VB.Menu a_acc 
         Caption         =   "BOOK ACCESSION RECORD "
      End
      Begin VB.Menu bok 
         Caption         =   "BOOK RECORD"
      End
   End
   Begin VB.Menu issue_return 
      Caption         =   "ISSUE RECORD"
      Begin VB.Menu issue 
         Caption         =   "ISSUE BOOK RECORD"
      End
      Begin VB.Menu return 
         Caption         =   "RETURN BOOK RECORD"
      End
   End
   Begin VB.Menu rep 
      Caption         =   "REPORT"
   End
   Begin VB.Menu user_account 
      Caption         =   "USER ACCOUNT"
      Begin VB.Menu change 
         Caption         =   "CHANGE PASSWORD"
      End
      Begin VB.Menu logout 
         Caption         =   "LOGOUT"
      End
      Begin VB.Menu exit 
         Caption         =   "EXIT"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub a_acc_Click()
book_accession.Show
End Sub
Private Sub bb_author_Click()
b_author.Show
End Sub

Private Sub bok_Click()
books.Show
End Sub

Private Sub change_Click()
changepassword.Show
End Sub
Private Sub Command1_Click()
Unload Me
login.Show
End Sub
Private Sub exit_Click()
Unload Me
End Sub
Private Sub issue_Click()
issue_book.Show
End Sub

Private Sub Label1_Click()

End Sub

Private Sub logout_Click()
login.Show
Unload Me
End Sub
Private Sub MDIForm_Load()
Label1.Caption = "BOOK ISSUE RECORD IN LIBRARY"
End Sub
Private Sub memb_Click()
member.Show
End Sub
Private Sub pub_Click()
publisher.Show
End Sub

Private Sub report_Click()

End Sub

Private Sub rep_Click()
report.Show
End Sub

Private Sub return_Click()
returnbook.Show
End Sub

Private Sub Timer1_Timer()
Label1.Left = Label1.Left - 150
If Label1.Left + Label1.Width <= 0 Then
Label1.Left = Picture1.Width
End If
End Sub
