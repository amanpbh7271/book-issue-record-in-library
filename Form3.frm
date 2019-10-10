VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form report 
   Caption         =   "REPORT"
   ClientHeight    =   8355
   ClientLeft      =   0
   ClientTop       =   345
   ClientWidth     =   18810
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8355
   ScaleWidth      =   18810
   Begin VB.Frame Frame11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Frame8"
      Height          =   5655
      Left            =   5520
      TabIndex        =   50
      Top             =   5160
      Width           =   12855
      Begin MSDataGridLib.DataGrid DataGrid8 
         Height          =   3735
         Left            =   2880
         TabIndex        =   52
         Top             =   1680
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   6588
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
      Begin VB.CommandButton Command10 
         Caption         =   "SHOW"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   30
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3120
         TabIndex        =   51
         Top             =   720
         Width           =   3375
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      TabIndex        =   49
      Top             =   6840
      Width           =   2655
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Frame10"
      Height          =   4575
      Left            =   5520
      TabIndex        =   45
      Top             =   5280
      Width           =   7935
      Begin MSDataGridLib.DataGrid DataGrid7 
         Height          =   3135
         Left            =   960
         TabIndex        =   47
         Top             =   1200
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   5530
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   19
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
            Size            =   9.75
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
      Begin VB.CommandButton Command9 
         Caption         =   "show"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   30
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1200
         TabIndex        =   46
         Top             =   360
         Width           =   3735
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Frame9"
      ForeColor       =   &H80000017&
      Height          =   5055
      Left            =   7440
      TabIndex        =   37
      Top             =   3960
      Width           =   9015
      Begin MSDataGridLib.DataGrid DataGrid6 
         Height          =   3615
         Left            =   2760
         TabIndex        =   39
         Top             =   1320
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   6376
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
      Begin VB.CommandButton Command8 
         Caption         =   "SHOW"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2760
         TabIndex        =   38
         Top             =   360
         Width           =   4095
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Frame7"
      Height          =   4575
      Left            =   6480
      TabIndex        =   27
      Top             =   1680
      Width           =   10335
      Begin MSDataGridLib.DataGrid DataGrid5 
         Height          =   2175
         Left            =   960
         TabIndex        =   30
         Top             =   1680
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   3836
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
      Begin VB.CommandButton Command7 
         Caption         =   "SHOW"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6360
         TabIndex        =   29
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   3240
         TabIndex        =   28
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0E0FF&
         Caption         =   "ACC NO"
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
         Left            =   840
         TabIndex        =   36
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Frame6"
      Height          =   3375
      Left            =   7200
      TabIndex        =   23
      Top             =   1440
      Width           =   7095
      Begin VB.CommandButton Command6 
         Caption         =   "SHOW"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4680
         TabIndex        =   25
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Height          =   735
         Left            =   2760
         TabIndex        =   24
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0FFFF&
         Caption         =   "ACC NO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   22.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   40
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label5 
         Height          =   855
         Left            =   1800
         TabIndex        =   26
         Top             =   2040
         Width           =   3735
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Frame5"
      Height          =   4215
      Left            =   6000
      TabIndex        =   17
      Top             =   1200
      Width           =   12615
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
         Height          =   375
         Left            =   8760
         TabIndex        =   22
         Top             =   360
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTPicker5 
         Height          =   375
         Left            =   5640
         TabIndex        =   21
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   102825985
         CurrentDate     =   42847
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   375
         Left            =   2160
         TabIndex        =   20
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   102825985
         CurrentDate     =   42847
      End
      Begin MSDataGridLib.DataGrid DataGrid4 
         Height          =   2535
         Left            =   720
         TabIndex        =   19
         Top             =   1320
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   4471
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
      Begin VB.CommandButton Command5 
         Caption         =   "SHOW"
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
         Left            =   10680
         TabIndex        =   18
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0E0FF&
         Caption         =   "ACC_NO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7200
         TabIndex        =   42
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0E0FF&
         Caption         =   "ENDING DATE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   35
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0E0FF&
         Caption         =   "STARTING DATE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   34
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Frame4"
      Height          =   5055
      Left            =   6480
      TabIndex        =   14
      Top             =   1080
      Width           =   10695
      Begin MSDataGridLib.DataGrid DataGrid3 
         Height          =   2535
         Left            =   960
         TabIndex        =   16
         Top             =   1440
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   4471
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
      Begin VB.CommandButton Command4 
         BackColor       =   &H8000000D&
         Caption         =   "SHOW"
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
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Frame3"
      Height          =   4455
      Left            =   5760
      TabIndex        =   7
      Top             =   840
      Width           =   12375
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   2775
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   4895
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
      Begin VB.CommandButton Command3 
         Caption         =   "SHOW"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8280
         TabIndex        =   10
         Top             =   360
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   5640
         TabIndex        =   9
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   102825985
         CurrentDate     =   42842
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   2040
         TabIndex        =   8
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   102825985
         CurrentDate     =   42842
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0E0FF&
         Caption         =   "ENDING DATE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   33
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0E0FF&
         Caption         =   "STARTING DATE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Frame2"
      Height          =   5175
      Left            =   5400
      TabIndex        =   1
      Top             =   480
      Width           =   13335
      Begin VB.CommandButton command2 
         Caption         =   "SHOW"
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
         Left            =   6840
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   4080
         TabIndex        =   5
         Top             =   360
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         Format          =   102825985
         CurrentDate     =   42842
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3375
         Left            =   600
         TabIndex        =   3
         Top             =   960
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   5953
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
      Begin VB.Label Label6 
         BackColor       =   &H00FFC0FF&
         Caption         =   "SELECT DATE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   360
         Width           =   3615
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Frame1"
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   4815
      Begin VB.OptionButton Option9 
         BackColor       =   &H00C0C0FF&
         Caption         =   "TITLE WISE BOOK AVAILABILTY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   53
         Top             =   4440
         Width           =   3735
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H00C0C0FF&
         Caption         =   "TITLE WISE BOOK QUANTITY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   48
         Top             =   3960
         Width           =   3495
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00C0C0FF&
         Caption         =   "ISSUE COUNT OF BOOKS "
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
         Left            =   0
         TabIndex        =   44
         Top             =   3480
         Width           =   4335
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C0C0FF&
         Caption         =   "BOOKS RETURN AFTER SEVEN DAYS"
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
         Left            =   0
         TabIndex        =   43
         Top             =   3000
         Width           =   4215
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0C0FF&
         Caption         =   "NO OF TIMES BOOK HAS BEEN ISSUED"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   41
         Top             =   2640
         Width           =   4215
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0C0FF&
         Caption         =   "ISSUE DETAILS OF A BOOK"
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
         Left            =   0
         TabIndex        =   13
         Top             =   2160
         Width           =   4575
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0C0FF&
         Caption         =   "ISSUE BOOK RECORDS"
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
         Left            =   0
         TabIndex        =   12
         Top             =   1680
         Width           =   3855
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "ISSUE DETAILS BETWEEN TWO DAYS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   1200
         Width           =   4575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "ISSUE DETAILS ON A DAY "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   720
         Width           =   2895
      End
   End
End
Attribute VB_Name = "report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim issue As New ADODB.Recordset
''Dim c As Date



Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command10_Click()
issue.Open "select * from query11", cn
Set DataGrid8.DataSource = issue
Set issue = Nothing
End Sub

Private Sub Command2_Click()
k = DTPicker1.Value
''MsgBox k
''Debug.Print "select * from issue where return_date=" & c & " "
If k <> "" Then
  issue.Open "SELECT issue.memb_id, name ,title,issue.acc_no,issue_date,return_date FROM book,member,copy ,issue where member.memb_id=issue.memb_id and copy.acc_no=issue.acc_no and  copy.book_id=book.book_id and issue_date=#" & k & "# ", cn
End If
Set DataGrid1.DataSource = issue
 Set issue = Nothing

End Sub

Private Sub Command3_Click()
a = DTPicker2.Value
b = DTPicker3.Value
If a <> "" And b <> "" Then

''Debug.Print " issue.memb_id, name ,book.book_id,title,issue.acc_no,issue_date,return_date FROM book,member,copy ,issue where member.memb_id=issue.memb_id and copy.acc_no=issue.acc_no and  copy.book_id=book.book_id and issue_date>=#" & b & "#  and issue_date<= #" & a & "#"
issue.Open "SELECT issue.memb_id, name ,title,issue.acc_no,issue_date,return_date FROM book,member,copy ,issue where member.memb_id=issue.memb_id and copy.acc_no=issue.acc_no and  copy.book_id=book.book_id and issue_date>=#" & a & "#  and issue_date<= #" & b & "# ", cn
Set DataGrid2.DataSource = issue
Set issue = Nothing
End If


End Sub

Private Sub Command4_Click()

If k = "" Then
  issue.Open "SELECT issue.memb_id, name ,title,issue.acc_no,issue_date FROM book,member,copy ,issue where member.memb_id=issue.memb_id and copy.acc_no=issue.acc_no and  copy.book_id=book.book_id and  IsNull(return_date) ", cn
Set DataGrid3.DataSource = issue
Set issue = Nothing
End If

End Sub

Private Sub Command5_Click()
If Text1.Text <> "" Then
issue.Open "SELECT name, title ,issue_date,return_date From member, book, issue, Copy WHERE issue.memb_id=member.memb_id and copy.book_id=book.book_id  and   issue.acc_no=copy.acc_no and issue.acc_no= " & Val(Text1.Text) & " and  issue_date>= #" & DTPicker4.Value & "# and issue_date<=#" & DTPicker5.Value & "# ", cn
Set DataGrid4.DataSource = issue
Set issue = Nothing
End If
End Sub

Private Sub Command6_Click()
If Text2.Text <> "" Then
issue.Open "SELECT count(title) From member, book, issue, Copy WHERE issue.memb_id=member.memb_id and copy.book_id=book.book_id  and   issue.acc_no=copy.acc_no and issue.acc_no=" & Val(Text2.Text) & " ", cn
Label5.Caption = issue.Fields(0)
Label5.FontSize = 22
Label5.FontBold = True

Set issue = Nothing
End If
End Sub

Private Sub Command7_Click()
If Text3.Text <> "" Then
issue.Open "SELECT issue.memb_id, name ,title,issue.acc_no,issue_date,return_date  From member, book, issue, Copy WHERE issue.memb_id=member.memb_id and copy.book_id=book.book_id  and   issue.acc_no=copy.acc_no and issue.acc_no= " & Val(Text3.Text) & " and  return_date - issue_date >7 ", cn
Set DataGrid5.DataSource = issue
Set issue = Nothing
End If
End Sub

Private Sub Command8_Click()
issue.Open " SELECT * from query9 ", cn
Set DataGrid6.DataSource = issue
Set issue = Nothing
End Sub

Private Sub Command9_Click()
issue.Open "SELECT title, count(acc_no) AS total_book From Copy right join BOOK on Copy.book_id = BOOK.book_id GROUP BY title ORDER BY title ", cn
Set DataGrid7.DataSource = issue
Set issue = Nothing

End Sub

Private Sub Form_Load()
Frame2.Top = 480
Frame3.Top = 480
Frame4.Top = 480
Frame5.Top = 480
Frame6.Top = 480
Frame7.Top = 480
Frame9.Top = 480
Frame10.Top = 480
Frame11.Top = 480
Frame2.Left = 5400
Frame3.Left = 5400
Frame4.Left = 5400
Frame5.Left = 5400
Frame6.Left = 5400
Frame7.Left = 5400
Frame9.Left = 5400
Frame10.Left = 5400
Frame11.Left = 5400





Frame1.Caption = ""
Frame2.Caption = ""
Frame3.Caption = ""
Frame4.Caption = ""
Frame5.Caption = ""
Frame6.Caption = ""
Frame7.Caption = ""
Frame7.Caption = ""
Frame9.Caption = ""
Frame10.Caption = ""
Frame11.Caption = ""
Frame2.Visible = False

Frame3.Visible = False

Frame4.Visible = False
Frame5.Visible = False

Frame6.Visible = False
Frame7.Visible = False
Frame9.Visible = False
Frame10.Visible = False
Frame11.Visible = False
End Sub

Private Sub Option1_Click()
Frame2.Visible = True
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = False
Frame7.Visible = False
Frame9.Visible = False
Frame10.Visible = False
Frame11.Visible = False
End Sub

Private Sub Option2_Click()
Frame2.Visible = False
Frame3.Visible = True
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = False
Frame7.Visible = False
Frame9.Visible = False
Frame10.Visible = False
Frame11.Visible = False
End Sub

Private Sub Option3_Click()
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = True
Frame5.Visible = False
Frame6.Visible = False
Frame7.Visible = False
Frame9.Visible = False
Frame10.Visible = False
Frame11.Visible = False
End Sub

Private Sub Option4_Click()
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = True
Frame6.Visible = False
Frame7.Visible = False
Frame9.Visible = False
Frame10.Visible = False
Frame11.Visible = False
End Sub

Private Sub Option5_Click()
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = True
Frame7.Visible = False
Frame9.Visible = False
Frame10.Visible = False
Frame11.Visible = False
End Sub

Private Sub Option6_Click()
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = False
Frame7.Visible = True
Frame9.Visible = False
Frame10.Visible = False
Frame11.Visible = False
End Sub

Private Sub Option7_Click()
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = False
Frame7.Visible = False
Frame9.Visible = True
Frame10.Visible = False
Frame11.Visible = False
End Sub

Private Sub Option8_Click()
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = False
Frame7.Visible = False
Frame9.Visible = False
Frame10.Visible = True
Frame11.Visible = False
End Sub

Private Sub Option9_Click()
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = False
Frame7.Visible = False
Frame9.Visible = False
Frame10.Visible = False
Frame11.Visible = True

End Sub
