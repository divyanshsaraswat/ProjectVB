VERSION 5.00
Begin VB.Form Form3 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Form3"
   ClientHeight    =   8340
   ClientLeft      =   -60
   ClientTop       =   435
   ClientWidth     =   11880
   FillStyle       =   0  'Solid
   FontTransparent =   0   'False
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   8340
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.TextBox dno 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6480
      TabIndex        =   92
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox comment 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   91
      Top             =   5880
      Width           =   2895
   End
   Begin VB.CommandButton load 
      Caption         =   "&Load"
      Height          =   375
      Left            =   3480
      TabIndex        =   89
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton delete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   6720
      TabIndex        =   88
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton addnew 
      Caption         =   "&AddNew"
      Height          =   375
      Left            =   2280
      TabIndex        =   87
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton update 
      Caption         =   "&Update"
      Height          =   375
      Left            =   5640
      TabIndex        =   86
      Top             =   6720
      Width           =   975
   End
   Begin VB.TextBox scode 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10680
      TabIndex        =   84
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox duedate 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   83
      Top             =   5400
      Width           =   1095
   End
   Begin VB.TextBox sdate 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   82
      Top             =   4800
      Width           =   1095
   End
   Begin VB.TextBox through 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   81
      Top             =   3960
      Width           =   2175
   End
   Begin VB.TextBox partyname 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   80
      Top             =   4440
      Width           =   2175
   End
   Begin VB.TextBox pcode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   79
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox sowner 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   78
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5880
      TabIndex        =   77
      Top             =   5400
      Width           =   1095
   End
   Begin VB.TextBox com 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5880
      TabIndex        =   74
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox igw 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      DataField       =   "igw"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox gpw 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      DataField       =   "igw"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Invoicedate 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      DataField       =   "invoicedate"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd MMM yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2880
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox category 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   73
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox code 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8880
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton Next1 
      Caption         =   "Next"
      Height          =   375
      Left            =   5640
      TabIndex        =   72
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Prev 
      Caption         =   "Previous"
      Height          =   375
      Left            =   4320
      TabIndex        =   71
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox size 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   67
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox size1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   66
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox size2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   65
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox color 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   64
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox color1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   63
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox color2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   62
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox qul 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      TabIndex        =   61
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox qul1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      TabIndex        =   60
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox qul2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      TabIndex        =   59
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox gpur 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   52
      Top             =   3480
      Width           =   615
   End
   Begin VB.TextBox PC1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   51
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox PC2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   50
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox PC3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   49
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox makingcharges 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox pno 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10680
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox content3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox content2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox content4 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   28
      Text            =   "GOLD"
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox weight1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4080
      TabIndex        =   27
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox weight2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4080
      TabIndex        =   26
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox weight3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4080
      TabIndex        =   25
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox weight4 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox rate1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4920
      TabIndex        =   23
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox rate2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4920
      TabIndex        =   22
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox rate3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4920
      TabIndex        =   21
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox rate4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5040
      TabIndex        =   20
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox amt1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6120
      TabIndex        =   19
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox amt2 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6120
      TabIndex        =   18
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox amt3 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6120
      TabIndex        =   17
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox amt4 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6120
      TabIndex        =   16
      Top             =   3480
      Width           =   855
   End
   Begin VB.TextBox amt 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5880
      TabIndex        =   15
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox mcharges 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5880
      TabIndex        =   14
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox content1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "Dia."
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox mrate1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "0"
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox mrate2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "0"
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox mrate3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "0"
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox mrate4 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "0"
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox totamt 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton clear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   9960
      TabIndex        =   32
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton print1 
      Caption         =   "&Print"
      Height          =   375
      Left            =   7800
      TabIndex        =   31
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton save 
      Caption         =   "&Save"
      Height          =   375
      Left            =   4560
      TabIndex        =   12
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton close 
      Caption         =   "C&lose"
      Height          =   375
      Left            =   8880
      TabIndex        =   11
      Top             =   6720
      Width           =   975
   End
   Begin VB.Label Label17 
      BackColor       =   &H8000000E&
      Caption         =   "Comments:"
      Height          =   255
      Left            =   120
      TabIndex        =   90
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label Label16 
      BackColor       =   &H8000000E&
      Caption         =   "S.Code:"
      Height          =   255
      Left            =   9360
      TabIndex        =   85
      Top             =   480
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   2535
      Left            =   8640
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label Label12 
      BackColor       =   &H8000000E&
      Caption         =   "Net Amount:"
      Height          =   255
      Left            =   4080
      TabIndex        =   76
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Comm@(%)"
      Height          =   255
      Left            =   4080
      TabIndex        =   75
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Sz."
      Height          =   255
      Left            =   2640
      TabIndex        =   70
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label34 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Col."
      Height          =   255
      Left            =   1920
      TabIndex        =   69
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Qt."
      Height          =   255
      Left            =   1320
      TabIndex        =   68
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Weight(Ct.)"
      Height          =   375
      Left            =   3960
      TabIndex        =   58
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label32 
      BackColor       =   &H8000000E&
      Caption         =   "S.NO."
      Height          =   375
      Left            =   0
      TabIndex        =   57
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "2."
      Height          =   255
      Left            =   -120
      TabIndex        =   56
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "3."
      Height          =   255
      Left            =   -120
      TabIndex        =   55
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "4."
      Height          =   255
      Left            =   -120
      TabIndex        =   54
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "1."
      Height          =   255
      Left            =   -120
      TabIndex        =   53
      Top             =   2040
      Width           =   495
   End
   Begin VB.Line Line6 
      X1              =   -120
      X2              =   8040
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Pc."
      Height          =   375
      Left            =   3360
      TabIndex        =   48
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label13 
      BackColor       =   &H8000000E&
      Caption         =   "Price No.:"
      Height          =   255
      Left            =   9720
      TabIndex        =   47
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label11 
      BackColor       =   &H8000000E&
      Caption         =   "Total Amount:"
      Height          =   255
      Left            =   4080
      TabIndex        =   46
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label10 
      BackColor       =   &H8000000E&
      Caption         =   "Making Charges:"
      Height          =   375
      Left            =   4080
      TabIndex        =   45
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Line Line2 
      X1              =   -120
      X2              =   8040
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Amount"
      Height          =   255
      Left            =   6000
      TabIndex        =   44
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "SRt."
      Height          =   255
      Left            =   5160
      TabIndex        =   43
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Content"
      Height          =   255
      Left            =   360
      TabIndex        =   42
      Top             =   1680
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   12000
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "I Code:"
      Height          =   255
      Left            =   4080
      TabIndex        =   39
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Min.Rate"
      Height          =   255
      Left            =   7080
      TabIndex        =   38
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   12000
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label22 
      BackColor       =   &H8000000E&
      Caption         =   "Through:"
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label23 
      BackColor       =   &H8000000E&
      Caption         =   "Party Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label Label24 
      BackColor       =   &H8000000E&
      Caption         =   "Due Date:"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label25 
      BackColor       =   &H8000000E&
      Caption         =   "Date:"
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "DAKSH"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   720
      TabIndex        =   37
      Top             =   120
      Width           =   9375
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Caption         =   "Rec Date:"
      Height          =   255
      Left            =   1800
      TabIndex        =   40
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "IGW(Cts.)"
      Height          =   255
      Left            =   0
      TabIndex        =   41
      Top             =   960
      Width           =   975
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim res As New ADODB.Recordset
Dim res1 As New ADODB.Recordset
Dim picturepath As String
Dim flag As Boolean

Private Sub Addnew_Click()
On Error Resume Next
Clear_Click
Image1.Visible = True
Image1.Picture = LoadPicture()

str1 = "select min(nscode) from tblestsalevalue where chstatus= 'NA'"
res.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
If (IsNull(res.Fields(0))) Then
    res.close
    res.Open "select max(nscode) from tblestsalevalue ", MDIForm1.con1, adOpenDynamic, adLockOptimistic
    
    If (IsNull(res.Fields(0))) Then
        scode.Text = 1
        save.Enabled = True
    Else
        scode = res.Fields(0) + 1
        save.Enabled = True
    End If
Else
    scode = res.Fields(0)
    update.Enabled = True
End If
res.close
flag = True
Addnew.Enabled = False
load.Enabled = False
Code.SetFocus
End Sub

Private Sub Clear_Click()
sowner.Text = ""
Category.Text = ""
Form3.Code = ""
Form3.amt1 = ""
Form3.amt2 = ""
Form3.amt3 = ""
Form3.amt4 = ""
Form3.content1 = ""
Form3.content2 = ""
Form3.content3 = ""
Form3.content4 = ""
Form3.igw = 0
Form3.gpw = 0
Form3.Invoicedate = ""
Form3.pno = ""
'Form3.sno = ""
Form3.makingcharges = 0
Form3.mrate1 = ""
Form3.mrate2 = ""
Form3.mrate3 = ""
Form3.mrate4 = ""
Form3.rate1 = 0
Form3.rate2 = 0
Form3.rate3 = 0
Form3.rate4 = 0
Form3.mcharges = 0
Form3.weight1 = 0
Form3.weight2 = 0
Form3.weight3 = 0
Form3.weight4 = 0
Form3.PC1 = 0
Form3.PC2 = 0
Form3.PC3 = 0
Form3.gpur = 0
Form3.amt = 0
Form3.sowner.Text = ""
Form3.through.Text = ""
Form3.duedate.Text = ""
Form3.sdate.Text = ""
pcode.Text = ""
scode.Text = ""
comment = ""
partyname = ""
Addnew.Enabled = True
load.Enabled = True
save.Enabled = False
update.Enabled = False
Image1.Visible = False
End Sub

Private Sub close_Click()
Unload Me

End Sub

Private Sub code_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    If (Val(Code.Text) <> Empty) Then
     loading
     calculate
     rate1.SetFocus
    Else
     MsgBox ("Enter The Valid Code For Nevigation")
    End If
End If
End Sub

Private Sub code_LostFocus()
If (Val(Code.Text) <> Empty) Then
     loading
    calculate
Else
MsgBox ("Enter The Valid Code For Nevigation")
End If
End Sub
Public Sub calculate()
amt11 = Val(mrate1) * Val(weight1)
amt21 = Val(mrate2) * Val(weight2)
amt31 = Val(mrate3) * Val(weight3)
'a = Val(igw) - (Val(weight1) + Val(weight2) + Val(weight3)) + (Val(igw) - (Val(weight1) + Val(weight2) + Val(weight3))) * Val(gpw) / 100
'weight4 = a
amt14 = Val(mrate4) * Val(weight4)
totamt = Round(Val(amt11) + Val(amt21) + Val(amt31) + Val(amt14) + Val(makingcharges))
End Sub



Private Sub comment_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    If (save.Enabled = True) Then
        save.SetFocus
    Else
        update.SetFocus
    End If
End If
End Sub

Private Sub delete_Click()
On Error Resume Next
If (scode.Text = Empty) Then
MsgBox ("Please Enter The S.Code")
Exit Sub
End If
If (Code.Text = Empty) Then
MsgBox ("Please Enter Valid Item No.")
Exit Sub
End If

res.Open "tblestsalevalue", MDIForm1.con2, adOpenDynamic, adLockOptimistic
res.Addnew
'////////////
res!ncode = Val(Code.Text)
res!dinvoicedate = CDate(Invoicedate)
res!chowner = sowner.Text
res!chcategory = Category.Text
res!chpcode = pcode.Text

If (sdate.Text <> Empty) Then
res!dsdate = CDate(sdate.Text)
Else
res!dsdate = CDate(date)
End If

If (through.Text <> Empty) Then
res!chthrough = through.Text
Else
res!chthrough = ""
End If

If (picturepath <> Empty) Then
res!chpicture = picturepath
picturepath = ""
Else
res!chpicture = ""
End If

If (partyname.Text <> Empty) Then
res!chpartyname = partyname.Text
Else
res!chpartyname = ""
End If

res!chcontent2 = content2
res!chcontent3 = content3

res!nweight1 = Val(weight1)
res!nweight2 = Val(weight2)
res!nweight3 = Val(weight3)
res!nweight4 = Val(weight4)

res!nsrate1 = Val(rate1)
res!nsrate2 = Val(rate2)
res!nsrate3 = Val(rate3)
res!nsrate4 = Val(rate4)

res!nsmaking = Val(mcharges)
res!nscode = Val(scode.Text)
res!ndno = Val(dno.Text)
res!dfpaydate = CDate(date)
res!chcomment = comment
res!npno = pno
res.update
res.close

str1 = "select * from tblestsalevalue where nscode=" & Val(scode.Text)
res.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
res!ncode = 0
res!chstatus = "NA"
res!chItemname = ""
res!chpcode = ""
res!chowner = ""
res!chthrough = ""
res!chpartyname = ""
res!chpicture = ""
'res!dinvoicedate = ""
res!chcontent2 = ""
res!chcontent3 = ""
res!nweight1 = 0
res!nweight2 = 0
res!nweight3 = 0
res!nweight4 = 0
res!nsrate1 = 0
res!nsrate2 = 0
res!nsrate3 = 0
res!nsrate4 = 0
res!nigw = 0
res!nsmaking = 0
res!npayrec1 = 0
res!npayrec2 = 0
res!npayrec3 = 0
res!ncom = 0
res!chcomment = ""
res.update
res.close
'con1.close
MsgBox ("Record Deleted")
Clear_Click
save.Enabled = False
Addnew.Enabled = True
update.Enabled = False
load.Enabled = True
delete.Enabled = False
End Sub

Private Sub duedate_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
   comment.SetFocus
End If
End Sub

Private Sub Form_Load()
    igw.BorderStyle = vbFixedSingle
    gpw.BorderStyle = vbFixedSingle
    Invoicedate.BorderStyle = vbFixedSingle
  '  maker.BorderStyle = 1
     pcode.BorderStyle = 1
    qul.BorderStyle = vbFixedSingle
    color.BorderStyle = 1
    size.BorderStyle = 1
    PC1.BorderStyle = 1
    weight1.BorderStyle = 1
   ' rate1.BorderStyle = 1
   ' amt1.BorderStyle = 1
    mrate1.BorderStyle = 1
    content2.BorderStyle = 1
   
    PC2.BorderStyle = 1
    weight2.BorderStyle = 1
    'rate2.BorderStyle = 1
    'amt2.BorderStyle = 1
    mrate2.BorderStyle = 1
    content3.BorderStyle = 1
    qul1.BorderStyle = 1
    color1.BorderStyle = 1
    size1.BorderStyle = 1
    
    PC3.BorderStyle = 1
    weight3.BorderStyle = 1
    'rate3.BorderStyle = 1
    'amt3.BorderStyle = 1
    mrate3.BorderStyle = 1
    content4.BorderStyle = 1
    qul2.BorderStyle = 1
    color2.BorderStyle = 1
    size2.BorderStyle = 1
    gpur.BorderStyle = 1
    weight4.BorderStyle = 1
   ' rate4.BorderStyle = 1
   ' amt4.BorderStyle = 1
    mrate4.BorderStyle = 1
    'mcharges.BorderStyle = 1
    'makingcharges.BorderStyle = 1
  '  amt.BorderStyle = 1
   ' totamt.BorderStyle = 1
    'Text1.BorderStyle = 1
   ' Text2.BorderStyle = 1
    
'con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\lahar\vb\db2.mdb;Persist Security Info=False"
'con.Open
'res.Open "select ncode from tblcosting", con
'While (res.EOF() = False)
'code.AddItem (res.Fields(0))
'res.MoveNext
'Wend
'res.close
save.Enabled = False
update.Enabled = False
delete.Enabled = False
End Sub
Public Sub printing()
    result = MsgBox("Do You want Picture In PrintOut!", vbYesNoCancel, "Confirmation")
    If (Val(result) = 7) Then
        Image1.Visible = False
     ElseIf (Val(result) = 2) Then
        Exit Sub
    End If
    Form3.igw.BorderStyle = vbBSNone
    Form3.gpw.BorderStyle = vbBSNone
    Form3.Invoicedate.BorderStyle = vbBSNone
    qul.BorderStyle = 0
    pcode.BorderStyle = 0
    color.BorderStyle = 0
    size.BorderStyle = 0
    PC1.BorderStyle = 0
    weight1.BorderStyle = 0
   ' rate1.BorderStyle = 0
   ' amt1.BorderStyle = 0
    mrate1.BorderStyle = 0
    content2.BorderStyle = 0
 
    PC2.BorderStyle = 0
    weight2.BorderStyle = 0
   ' rate2.BorderStyle = 0
   ' amt2.BorderStyle = 0
    mrate2.BorderStyle = 0
    content3.BorderStyle = 0
    qul1.BorderStyle = 0
    color1.BorderStyle = 0
    size1.BorderStyle = 0
    
    PC3.BorderStyle = 0
    weight3.BorderStyle = 0
    'rate3.BorderStyle = 0
    'amt3.BorderStyle = 0
    mrate3.BorderStyle = 0
    content4.BorderStyle = 0
    qul2.BorderStyle = 0
    color2.BorderStyle = 0
    size2.BorderStyle = 0
    
    gpur.BorderStyle = 0
    weight4.BorderStyle = 0
   ' rate4.BorderStyle = 0
   ' amt4.BorderStyle = 0
    mrate4.BorderStyle = 0
    'mcharges.BorderStyle = 0
    'makingcharges.BorderStyle = 0
    'amt.BorderStyle = 0
    'totamt.BorderStyle = 0
   ' Text1.BorderStyle = 0
   ' Text2.BorderStyle = 0
    Form3.save.Visible = False
    Form3.save.Visible = False
    Form3.Clear.Visible = False
    Form3.close.Visible = False
    Form3.print1.Visible = False
    load.Visible = False
    Addnew.Visible = False
    update.Visible = False
    delete.Visible = False
    comment.Visible = False
    Label17.Visible = False
        Form3.Refresh
     'MsgBox ("changing")
    Form3.PrintForm
    Image1.Visible = True
    igw.BorderStyle = vbFixedSingle
    gpw.BorderStyle = vbFixedSingle
    Invoicedate.BorderStyle = vbFixedSingle
    pcode.BorderStyle = 1
    qul.BorderStyle = vbFixedSingle
    color.BorderStyle = 1
    size.BorderStyle = 1
    PC1.BorderStyle = 1
    weight1.BorderStyle = 1
    rate1.BorderStyle = 1
    amt1.BorderStyle = 1
    mrate1.BorderStyle = 1
    content2.BorderStyle = 1
   
    PC2.BorderStyle = 1
    weight2.BorderStyle = 1
    rate2.BorderStyle = 1
    amt2.BorderStyle = 1
    mrate2.BorderStyle = 1
    content3.BorderStyle = 1
    qul1.BorderStyle = 1
    color1.BorderStyle = 1
    size1.BorderStyle = 1
    
    PC3.BorderStyle = 1
    weight3.BorderStyle = 1
    rate3.BorderStyle = 1
    amt3.BorderStyle = 1
    mrate3.BorderStyle = 1
    content4.BorderStyle = 1
    qul2.BorderStyle = 1
    color2.BorderStyle = 1
    size2.BorderStyle = 1
    gpur.BorderStyle = 1
    weight4.BorderStyle = 1
    rate4.BorderStyle = 1
    amt4.BorderStyle = 1
    mrate4.BorderStyle = 1
   ' mcharges.BorderStyle = 1
   ' makingcharges.BorderStyle = 1
    'amt.BorderStyle = 1
   ' totamt.BorderStyle = 1
   ' Text1.BorderStyle = 1
   ' Text2.BorderStyle = 1
    
    Form3.save.Visible = True
    Form3.Clear.Visible = True
    Form3.close.Visible = True
    Form3.print1.Visible = True
    load.Visible = True
    Addnew.Visible = True
    update.Visible = True
    delete.Visible = True
    comment.Visible = True
    Label17.Visible = True
End Sub
Public Sub loading()
On Error Resume Next
     str1 = "select * from tblcosting where ncode=" & Val(Code.Text)
    res.Open str1, MDIForm1.con1
    If (res.EOF = True) Then
    MsgBox ("Recod Not Found")
   ' res.close
    Else
        Code.Text = res!ncode
        Category = res!chcategory
        sowner = res!chowner
        If (IsNull(res!chpcode) = False) Then
        pcode = res!chpcode
        End If
        Invoicedate = Format(res!dinvoicedate, "dd-mmm-yy")
        content1 = res!chcontent1
        content2 = res!chcontent2
        strcontent2 = res!chcontent2
        
        If (strcontent2 = Empty) Then
        content2.Text = ""
        Else
        content2.Text = strcontent2
        End If
        
        strcontent3 = res!chcontent3
        
        If (strcontent3 = Empty) Then
        content3.Text = ""
        Else
        content3.Text = strcontent3
        End If
        
        If (res!chcontent4 = Empty) Then
        content4.Text = ""
        Else
        content4.Text = res!chcontent4
        End If
        
        igw = res!nigw
        weight1 = res!nweight1
        weight2 = res!nweight2
        weight3 = res!nweight3
        If (Val(res!pcs1) <> 0) Then
        PC1 = res!pcs1
        End If
        
        If (Val(res!pcs2) <> 0) Then
        PC2 = res!pcs2
        End If
        If (Val(res!pcs3) <> 0) Then
        PC3 = res!pcs3
        End If
                        
        If (Val(res!minrate1) <> 0) Then
        mrate1 = res!minrate1
        End If
        
        If (Val(res!minrate2) <> 0) Then
        mrate2 = res!minrate2
        End If
        If (Val(res!minrate3) <> 0) Then
        mrate3 = res!minrate3
        End If
        
        If (Val(res!minrate4) <> 0) Then
        mrate4 = res!minrate4
        End If
        
        
        gpur = res!gpur
        gpur = gpur + "K"
       
        gpw = res!ngpw
        makingcharges = res!nmaking1
        weight4 = Round(res!nweight4, 2)
        'a = (Val(igw) - (Val(weight1) + Val(weight2) + Val(weight3)) + (Val(igw) - (Val(weight1) + Val(weight2) + Val(weight3))) * Val(gpw) / 100) / 5
        'weight4 = a
        Sno = res!chcategory & "-" & Code.Text
        
        qul = res!chquality1
        qul1 = res!chquality2
        qul2 = res!chquality3
        If (IsNull(res!chcolor1) = False) Then
        color = res!chcolor1
        End If
        If (IsNull(res!chcolor2) = False) Then
        color1 = res!chcolor2
        End If
        If (IsNull(res!chcolor3) = False) Then
        color2 = res!chcolor3
        End If
        
        If (Val(res!chsize1) <> 0) Then
        size = res!chsize1
        End If
        
        If (Val(res!chsize2) <> 0) Then
        size1 = res!chsize2
        End If
        
        If (Val(res!chsize3) <> 0) Then
        size2 = res!chsize3
        End If
        
        If (IsNull(res!ndno) = False) Then
        dno = res!ndno
        Else
        dno = 0
        End If
        
        If (res!opicture <> Empty) Then
        Dim picturename As String
        pos = InStrRev(res!opicture, "\")
        If (pos <> 0) Then
        picturename = Mid(res!opicture, pos + 1)
        Else
        picturename = res!opicture
        End If
        Image1.Visible = True
        Image1.Picture = LoadPicture(MDIForm1.picturepath & picturename)
        picturepath = picturename
        Else
        Image1.Visible = False
        picturepath = ""
        End If
       End If
        res.close
        End Sub

Private Sub Form_Unload(Cancel As Integer)
'con.close
End Sub


Private Sub load_Click()
 flag = False
On Error Resume Next
If (scode.Text = Empty) Then
MsgBox ("Please Enter The S.Code")
Exit Sub
End If
On Error GoTo errorhandler
str1 = "select * from tblestsalevalue where nscode=" & scode.Text
'MsgBox (str1)
res.Open str1, MDIForm1.con1

If (res.EOF = True) Then
MsgBox ("No Record Found")
res.close
Exit Sub
End If
'////////////
Code.Text = Val(res!ncode)
Invoicedate = res!dinvoicedate
sowner.Text = res!chowner
Category.Text = res!chcategory

If (IsNull(res!npno) = False) Then
pno = Val(res!npno)
End If

If (IsNull(res!chpcode) = False) Then
pcode.Text = res!chpcode
End If
If (res!dsdate <> Empty) Then
 sdate = res!dsdate
Else
sdate = CDate(date)
End If

If (res!dduedate <> Empty) Then
duedate.Text = res!dduedate
Else
duedate.Text = CDate(date)
End If


If (res!chthrough <> Empty) Then
through.Text = res!chthrough
Else
through = ""
End If

 If (res!chpicture <> Empty) Then
        Dim picturename As String
        pos = InStrRev(res!chpicture, "\")
        If (pos <> 0) Then
        picturename = Mid(res!chpicture, pos + 1)
        Else
        picturename = res!chpicture
        End If
        Image1.Visible = True
        Image1.Picture = LoadPicture(MDIForm1.picturepath & picturename)
        picturepath = picturename
        Else
        Image1.Visible = False
        picturepath = ""
        End If
        
'If (res!chpicture <> Empty) Then
'picturepath = res!chpicture
'Else
'picturepath = ""
'End If

If (res!chpartyname <> Empty) Then
partyname.Text = res!chpartyname
Else
partyname = ""
End If

content2 = res!chcontent2
content3 = res!chcontent3
weight1 = Val(res!nweight1)
weight2 = Val(res!nweight2)
weight3 = Val(res!nweight3)
weight4 = Val(res!nweight4)

rate1 = Val(res!nsrate1)
rate2 = Val(res!nsrate2)
rate3 = Val(res!nsrate3)
rate4 = Val(res!nsrate4)
mcharges = Val(res!nsmaking)
If (IsNull(res!chcomment) = False) Then
comment = res!chcomment
Else
comment = ""
End If

If (IsNull(res!ndno) = False) Then
dno = res!ndno
Else
dno = 0
End If

com = Val(res!ncom)
res.close
calculate
update.Enabled = True
delete.Enabled = True
errorhandler:
If (Err.Number = 53) Then
MsgBox ("NO IMAGE FOUND")
Exit Sub
ElseIf (Err.Number = 94) Then
Resume Next
Else
Exit Sub
End If

End Sub

Private Sub makingcharges_Change()
totamt = (Val(mrate1) * Val(weight1)) + (Val(mrate2) * Val(weight2)) + (Val(mrate3) * Val(weight3)) + (Val(mrate4) * Val(weight4)) + Val(makingcharges)
End Sub

Private Sub mcharges_Change()
amt = (Val(rate1) * Val(weight1)) + (Val(rate2) * Val(weight2)) + (Val(rate3) * Val(weight3)) + (Val(rate4) * Val(weight4)) + Val(mcharges)
End Sub

Private Sub mcharges_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    through.SetFocus
    End If
End Sub

Private Sub mrate1_Change()
totamt = (Val(mrate1) * Val(weight1)) + (Val(mrate2) * Val(weight2)) + (Val(mrate3) * Val(weight3)) + (Val(mrate4) * Val(weight4)) + Val(makingcharges)
End Sub

Private Sub mrate2_Change()
totamt = (Val(mrate1) * Val(weight1)) + (Val(mrate2) * Val(weight2)) + (Val(mrate3) * Val(weight3)) + (Val(mrate4) * Val(weight4)) + Val(makingcharges)
End Sub

Private Sub mrate3_Change()
totamt = (Val(mrate1) * Val(weight1)) + (Val(mrate2) * Val(weight2)) + (Val(mrate3) * Val(weight3)) + (Val(mrate4) * Val(weight4)) + Val(makingcharges)
End Sub

Private Sub mrate4_Change()
totamt = (Val(mrate1) * Val(weight1)) + (Val(mrate2) * Val(weight2)) + (Val(mrate3) * Val(weight3)) + (Val(mrate4) * Val(weight4)) + Val(makingcharges)
End Sub

Private Sub print_Click()
End Sub

Private Sub Next1_Click()
If (Code.Text <> Empty) Then
    loading
    calculate
Else
MsgBox " Enter the valid code"
End If
End Sub







Private Sub partyname_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    sdate.SetFocus
    End If
End Sub



Private Sub Prev_Click()
If (Code.Text <> Empty) Then
loading
calculate
Else
MsgBox ("Enter The Valid Code For Nevigation")
End If
End Sub

Private Sub print1_Click()
printing
If (Val(ncode) <> 0) Then
str1 = "selet pflag from tblcosting where ncode=" & Val(Code.Text)
res.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
res!pflag = "Printed."
res.update
res.close
End If
End Sub

Private Sub rate1_Change()
amt1 = Val(rate1) * Val(weight1)
amt = (Val(rate1) * Val(weight1)) + (Val(rate2) * Val(weight2)) + (Val(rate3) * Val(weight3)) + (Val(rate4) * Val(weight4)) + Val(mcharges)
End Sub

Private Sub rate1_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    'load_Click
    rate2.SetFocus
    'rate1.SelText (rate1.Text)
    End If
End Sub

Private Sub rate2_Change()
amt2 = Val(rate2) * Val(weight2)
amt = (Val(rate1) * Val(weight1)) + (Val(rate2) * Val(weight2)) + (Val(rate3) * Val(weight3)) + (Val(rate4) * Val(weight4)) + Val(mcharges)
End Sub

Private Sub rate2_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    'load_Click
    rate3.SetFocus
    'rate1.SelText (rate1.Text)
    End If
End Sub

Private Sub rate3_Change()
amt3 = Val(rate3) * Val(weight3)
amt = (Val(rate1) * Val(weight1)) + (Val(rate2) * Val(weight2)) + (Val(rate3) * Val(weight3)) + (Val(rate4) * Val(weight4)) + Val(mcharges)
End Sub

Private Sub rate3_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    rate4.SetFocus
    'rate1.SelText (rate1.Text)
    End If
End Sub

Private Sub rate4_Change()
amt4 = Val(rate4) * Val(weight4)
amt = (Val(rate1) * Val(weight1)) + (Val(rate2) * Val(weight2)) + (Val(rate3) * Val(weight3)) + (Val(rate4) * Val(weight4)) + Val(mcharges)
End Sub

Private Sub rate4_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
   ' load_Click
    mcharges.SetFocus
    'rate1.SelText (rate1.Text)
    End If
End Sub

Private Sub save_Click()
'Dim con1 As New ADODB.Connection
'con1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\lahar\vb\db2.mdb;Persist Security Info=False"
'con.Open
On Error Resume Next
If (scode.Text = Empty) Then
MsgBox ("Please Enter The S.Code")
Exit Sub
End If
If (Code.Text = Empty) Then
MsgBox ("Please Enter Valid Item No.")
Exit Sub
End If
 
    str1 = "Select chissue,nappno from tblappmaster where nappno="
    str1 = str1 & " ( Select  nappno from tblappdetail where ncode=" & Val(Code.Text) & ") "
    res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
    name1 = ""

   If (res1.EOF = False) Then
     result = MsgBox("This Item Found in the Approval No " & res1!nappno & " And Issue to " & res1!chissue & ". Do You Want To Remove This? ", vbYesNo + vbDefaultButton2, "Confirmation")

    If (Val(result) = 7) Then
        res1.close
       ' code.Text = ""
        Code.SetFocus
        Exit Sub
    Else
        str1 = " Select  * from tblappdetail where ncode=" & Val(Code.Text)
        res1.close
        res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
        res1.delete
        res1.close
    End If
  Else
    res1.close
End If

res.Open "tblestsalevalue", MDIForm1.con1, adOpenDynamic, adLockOptimistic
res.Addnew
'////////////
res!ncode = Val(Code.Text)
res!dinvoicedate = CDate(Invoicedate)
res!chowner = sowner.Text
res!chcategory = Category.Text
res!chpcode = pcode.Text
If (sdate.Text <> Empty) Then
res!dsdate = CDate(sdate.Text)
Else
res!dsdate = CDate(date)
End If

If (duedate.Text <> Empty) Then
res!dduedate = CDate(duedate.Text)
Else
res!dduedate = CDate(date)
End If

If (through.Text <> Empty) Then
res!chthrough = through.Text
Else
res!chthrough = ""
End If

If (picturepath <> Empty) Then
res!chpicture = picturepath
picturepath = ""
Else
res!chpicture = ""
End If

If (partyname.Text <> Empty) Then
res!chpartyname = partyname.Text
Else
res!chpartyname = ""
End If
res!npno = Val(pno)
res!chcontent2 = content2
res!chcontent3 = content3
res!nweight1 = Val(weight1)
res!nweight2 = Val(weight2)
res!nweight3 = Val(weight3)
res!nweight4 = Val(weight4)
res!nsrate1 = Val(rate1)
res!nsrate2 = Val(rate2)
res!nsrate3 = Val(rate3)
res!nsrate4 = Val(rate4)
res!nsmaking = Val(mcharges)
res!nscode = Val(scode.Text)
res!ndno = dno.Text
res!ncom = Val(com)
res!chstatus = "A"
res!chcomment = comment
res.update
res.close

str1 = "select * from tblcosting where ncode=" & Val(Code.Text)
res.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
res!ncode = Val(Code)
res!chItemname = ""
res!chpcode = ""
'res!dinvoicedate = ""
res!chcontent1 = "Dia."
res!chcontent2 = ""
res!chcontent3 = ""
res!chcontent4 = "Gold"
res!nweight1 = 0
res!nweight2 = 0
res!nweight3 = 0
res!nweight4 = 0
res!nrate1 = 0
res!nrate2 = 0
res!nrate3 = 0
res!nrate4 = 0
res!nigw = 0
res!nmaking = 0
res!nmaking1 = 0
res!chmaker = ""
res!chpcode = ""
'res!drecdate = ""
res!chissueto = ""
res!ngpw = 0
res!chstate = "Not Avi."
res!pcs1 = 0
res!pcs2 = 0
res!pcs2 = 0
res!gpur = 18
res!minrate1 = 0
res!minrate2 = 0
res!minrate3 = 0
res!minrate4 = 0
res!nmaking1 = 0
res!chquality1 = ""
res!chquality2 = ""
res!chquality3 = ""
res!chcolor1 = ""
res!chcolor2 = ""
res!chcolor3 = ""
res!chsize1 = ""
res!chsize2 = ""
res!chsize3 = ""
res!chcategory = ""
res!chowner = ""
res!opicture = ""
res.update
res.close
'con1.close
MsgBox ("Record Saved")
Clear_Click
save.Enabled = False
Addnew.Enabled = True
update.Enabled = True
load.Enabled = True
End Sub
Private Sub scode_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    load_Click
    rate1.SetFocus
    'rate1.SelText (rate1.Text)
    End If
End Sub


Private Sub sdate_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    duedate.SetFocus
    End If
End Sub

Private Sub through_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    partyname.SetFocus
    End If
End Sub

Private Sub totamt_change()
pno = Round(totamt * 0.01333333)
End Sub

Private Sub update_Click()
On Error Resume Next
If (scode.Text = Empty) Then
MsgBox ("Please Enter The S.Code")
Exit Sub
End If
If (Code.Text = Empty) Then
MsgBox ("Please Enter Valid Item No.")
Exit Sub
End If

If (flag = True) Then
    str1 = "Select chissue,nappno from tblappmaster where nappno="
    str1 = str1 & " ( Select  nappno from tblappdetail where ncode=" & Val(Code.Text) & ") "
    res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
    name1 = ""
  If (res1.EOF = False) Then
    response = MsgBox("This Item Found in the Approval No " & res1!nappno & " And Issue to " & res1!chissue & ". Do You Want To Delete This? ", vbYesNo + vbDefaultButton2, "Confirmation")
 '    MsgBox (response)
    If (Val(response) = 7) Then
        res1.close
        'code.Text = ""
        Code.SetFocus
        flag = False
        Exit Sub
    Else
        str1 = " Select  * from tblappdetail where ncode=" & Val(Code.Text)
        res1.close
        res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
        res1.delete
        res1.close
    End If
  Else
  res1.close
 End If

str1 = "select * from tblcosting where ncode=" & Val(Code.Text)
res.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
res!ncode = Val(Code)
res!chItemname = ""
res!chpcode = ""
'res!dinvoicedate = ""
res!chcontent1 = "Dia."
res!chcontent2 = ""
res!chcontent3 = ""
res!chcontent4 = "Gold"
res!nweight1 = 0
res!nweight2 = 0
res!nweight3 = 0
res!nweight4 = 0
res!nrate1 = 0
res!nrate2 = 0
res!nrate3 = 0
res!nrate4 = 0
res!nigw = 0
res!nmaking = 0
res!nmaking1 = 0
res!chmaker = ""
res!chpcode = ""
'res!drecdate = ""
res!chissueto = ""
res!ngpw = 0
res!chstate = "Not Avi."
res!pcs1 = 0
res!pcs2 = 0
res!pcs2 = 0
res!gpur = 18
res!minrate1 = 0
res!minrate2 = 0
res!minrate3 = 0
res!minrate4 = 0
res!nmaking1 = 0
res!chquality1 = ""
res!chquality2 = ""
res!chquality3 = ""
res!chcolor1 = ""
res!chcolor2 = ""
res!chcolor3 = ""
res!chsize1 = ""
res!chsize2 = ""
res!chsize3 = ""
res!chcategory = ""
res!chowner = ""
res!opicture = ""
res.update
res.close
End If
flag = False

str1 = "Select * from tblestsalevalue where nscode=" & scode.Text
res.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
'res.addnew
'////////////
res!ncode = Val(Code.Text)
res!dinvoicedate = CDate(Invoicedate)
res!chowner = sowner.Text
res!chcategory = Category.Text
res!chpcode = pcode.Text
If (sdate.Text <> Empty) Then
res!dsdate = CDate(sdate.Text)
Else
res!dsdate = CDate(date)
End If

If (duedate.Text <> Empty) Then
res!dduedate = CDate(duedate.Text)
Else
res!dduedate = CDate(date)
End If

If (through.Text <> Empty) Then
res!chthrough = through.Text
Else
res!chthrough = ""
End If

If (picturepath <> Empty) Then
res!chpicture = picturepath
picturepath = ""
Else
res!chpicture = ""
End If

If (partyname.Text <> Empty) Then
res!chpartyname = partyname.Text
Else
res!chpartyname = ""
End If

res!ndno = dno.Text
res!chcontent2 = content2
res!chcontent3 = content3
res!nweight1 = Val(weight1)
res!nweight2 = Val(weight2)
res!nweight3 = Val(weight3)
res!nweight4 = Val(weight4)
res!nsrate1 = Val(rate1)
res!nsrate2 = Val(rate2)
res!nsrate3 = Val(rate3)
res!nsrate4 = Val(rate4)
res!nsmaking = Val(mcharges)
res!npno = Val(pno)
'res!chscode = scode.Text
res!ncom = Val(com)
res!chstatus = "A"
res!chcomment = comment
res.update
res.close
MsgBox ("Record Updated")
Clear_Click
update.Enabled = False
load.Enabled = True
save.Enabled = False
Addnew.Enabled = True
End Sub
