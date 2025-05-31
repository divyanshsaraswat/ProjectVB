VERSION 5.00
Begin VB.Form Form3 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Form3"
   ClientHeight    =   8355
   ClientLeft      =   -675
   ClientTop       =   945
   ClientWidth     =   12045
   FillStyle       =   0  'Solid
   FontTransparent =   0   'False
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   8355
   ScaleWidth      =   12045
   WindowState     =   2  'Maximized
   Begin VB.TextBox igw 
      Appearance      =   0  'Flat
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
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox gpw 
      Appearance      =   0  'Flat
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
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox Invoicedate 
      Appearance      =   0  'Flat
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
      Left            =   6120
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox category 
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
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   81
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox code 
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
      Left            =   8640
      TabIndex        =   82
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton Next1 
      Caption         =   "Next"
      Height          =   375
      Left            =   5640
      TabIndex        =   80
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Prev 
      Caption         =   "Previous"
      Height          =   375
      Left            =   4320
      TabIndex        =   79
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox size 
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
      Height          =   285
      Left            =   2640
      TabIndex        =   75
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox size1 
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
      Height          =   285
      Left            =   2640
      TabIndex        =   74
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox size2 
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
      Left            =   2640
      TabIndex        =   73
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox color 
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
      Height          =   285
      Left            =   1920
      TabIndex        =   72
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox color1 
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
      Height          =   285
      Left            =   1920
      TabIndex        =   71
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox color2 
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
      Left            =   1920
      TabIndex        =   70
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox qul 
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
      Height          =   285
      Left            =   1200
      TabIndex        =   69
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox qul1 
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
      Height          =   285
      Left            =   1200
      TabIndex        =   68
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox qul2 
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
      Left            =   1200
      TabIndex        =   67
      Top             =   3000
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   8280
      ScaleHeight     =   3465
      ScaleWidth      =   2985
      TabIndex        =   60
      Top             =   1560
      Width           =   3015
   End
   Begin VB.TextBox gpur 
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
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   59
      Top             =   3480
      Width           =   615
   End
   Begin VB.TextBox PC1 
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
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   58
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox PC2 
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
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   57
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox PC3 
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
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   56
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
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox pno 
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
      Left            =   10440
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox content3 
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
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox content2 
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
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox content4 
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
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   27
      Text            =   "GOLD"
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox weight1 
      Appearance      =   0  'Flat
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
      TabIndex        =   26
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox weight2 
      Appearance      =   0  'Flat
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
      TabIndex        =   25
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox weight3 
      Appearance      =   0  'Flat
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
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox weight4 
      Appearance      =   0  'Flat
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
      TabIndex        =   23
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
      TabIndex        =   22
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
      TabIndex        =   21
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
      TabIndex        =   20
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
      TabIndex        =   19
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
      TabIndex        =   18
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
      TabIndex        =   17
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
      TabIndex        =   16
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
      TabIndex        =   15
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
      TabIndex        =   14
      Top             =   4680
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
      TabIndex        =   13
      Top             =   4080
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
      TabIndex        =   12
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox mrate1 
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
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "0"
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox mrate2 
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
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "0"
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox mrate3 
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
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "0"
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox mrate4 
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
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   8
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
      TabIndex        =   9
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton clear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   7080
      TabIndex        =   31
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton print1 
      Caption         =   "&Print"
      Height          =   375
      Left            =   4200
      TabIndex        =   30
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton save 
      Caption         =   "&Save"
      Height          =   375
      Left            =   2880
      TabIndex        =   11
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton close 
      Caption         =   "C&lose"
      Height          =   375
      Left            =   5640
      TabIndex        =   10
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Sz."
      Height          =   255
      Left            =   2640
      TabIndex        =   78
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label34 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Col."
      Height          =   255
      Left            =   1920
      TabIndex        =   77
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Qt."
      Height          =   255
      Left            =   1320
      TabIndex        =   76
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Weight(Ct.)"
      Height          =   375
      Left            =   3960
      TabIndex        =   66
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label32 
      BackColor       =   &H8000000E&
      Caption         =   "S.NO."
      Height          =   375
      Left            =   0
      TabIndex        =   65
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "2."
      Height          =   255
      Left            =   -120
      TabIndex        =   64
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "3."
      Height          =   255
      Left            =   -120
      TabIndex        =   63
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "4."
      Height          =   255
      Left            =   -120
      TabIndex        =   62
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "1."
      Height          =   255
      Left            =   -120
      TabIndex        =   61
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
      TabIndex        =   55
      Top             =   1680
      Width           =   615
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   12000
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Label Label13 
      BackColor       =   &H8000000E&
      Caption         =   "Price No.:"
      Height          =   255
      Left            =   9480
      TabIndex        =   54
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label11 
      BackColor       =   &H8000000E&
      Caption         =   "Total Amount:"
      Height          =   255
      Left            =   4080
      TabIndex        =   53
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label Label10 
      BackColor       =   &H8000000E&
      Caption         =   "Making Charges:"
      Height          =   375
      Left            =   4080
      TabIndex        =   52
      Top             =   4080
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
      TabIndex        =   51
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "SRt."
      Height          =   255
      Left            =   5160
      TabIndex        =   50
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Content"
      Height          =   255
      Left            =   360
      TabIndex        =   49
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
      Left            =   7200
      TabIndex        =   45
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Min.Rate"
      Height          =   255
      Left            =   7080
      TabIndex        =   44
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   12000
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Rec2"
      Height          =   375
      Left            =   7080
      TabIndex        =   42
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Rec1"
      Height          =   375
      Left            =   5400
      TabIndex        =   41
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Amount"
      Height          =   375
      Left            =   3840
      TabIndex        =   40
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Comm@(%)"
      Height          =   375
      Left            =   2400
      TabIndex        =   39
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "EST. Amount"
      Height          =   375
      Left            =   120
      TabIndex        =   38
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Rec3"
      Height          =   375
      Left            =   8760
      TabIndex        =   37
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label22 
      BackColor       =   &H8000000E&
      Caption         =   "Through:"
      Height          =   255
      Left            =   5760
      TabIndex        =   36
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Label Label23 
      BackColor       =   &H8000000E&
      Caption         =   "Party Name:"
      Height          =   255
      Left            =   720
      TabIndex        =   35
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label Label24 
      BackColor       =   &H8000000E&
      Caption         =   "Due Date:"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Label25 
      BackColor       =   &H8000000E&
      Caption         =   "Date:"
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   4080
      Width           =   975
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   12000
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Balance"
      Height          =   375
      Left            =   10320
      TabIndex        =   32
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Lahar Exports"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   600
      TabIndex        =   43
      Top             =   120
      Width           =   9375
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Caption         =   "Rec Date:"
      Height          =   255
      Left            =   4920
      TabIndex        =   46
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "GWes.@(%)"
      Height          =   255
      Left            =   2880
      TabIndex        =   48
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "IGW(Cts.)"
      Height          =   255
      Left            =   720
      TabIndex        =   47
      Top             =   960
      Width           =   1095
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim res As New ADODB.Recordset

Private Sub clear_Click()
Form3.code = ""
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
Form3.mrate1 = 0
Form3.mrate2 = 0
Form3.mrate3 = 0
Form3.mrate4 = 0
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
save.Enabled = True
Picture1.Visible = False
End Sub

Private Sub CLOSE_Click()
Unload Me

End Sub

Private Sub code_LostFocus()
If (Val(code.Text) <> Empty) Then
     loading
    Calculate
Else
MsgBox ("Enter The Valid Code For Nevigation")
End If
End Sub
Public Sub Calculate()
amt11 = Val(mrate1) * Val(weight1)
amt21 = Val(mrate2) * Val(weight2)
amt31 = Val(mrate3) * Val(weight3)
'a = Val(igw) - (Val(weight1) + Val(weight2) + Val(weight3)) + (Val(igw) - (Val(weight1) + Val(weight2) + Val(weight3))) * Val(gpw) / 100
'weight4 = a
amt14 = Val(mrate4) * Val(weight4)
totamt = Round(Val(amt11) + Val(amt21) + Val(amt31) + Val(amt14) + Val(makingcharges))
End Sub
Private Sub Form_Load()
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\lahar\vb\db2.mdb;Persist Security Info=False"
con.Open
'res.Open "select ncode from tblcosting", con
'While (res.EOF() = False)
'code.AddItem (res.Fields(0))
'res.MoveNext
'Wend
'res.close
End Sub
Public Sub printing()
    result = MsgBox("Do You want Picture In PritOut!", vbYesNoCancel, "Confirmation")
    If (Val(result) = 7) Then
        Picture1.Visible = False
    ' ElseIf (Val(result) = 2) Then
     '  Else
    End If
    Form3.save.Visible = False
    Form3.save.Visible = False
    Form3.clear.Visible = False
    Form3.CLOSE.Visible = False
    Form3.print1.Visible = False
    Form3.Refresh
    Form3.PrintForm
    Picture1.Visible = True
    Form3.save.Visible = True
    Form3.clear.Visible = True
    Form3.CLOSE.Visible = True
    Form3.print1.Visible = True

End Sub
Public Sub loading()
     str1 = "select * from tblcosting where ncode=" & Val(code.Text)
    res.Open str1, con
    If (res.EOF = True) Then
    MsgBox ("Recod Not Found")
    res.CLOSE
    Else
        code.Text = res!ncode
        Category = res!chcategory
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
        PC1 = res!pcs1
        PC2 = res!pcs2
        PC3 = res!pcs3
        mrate1 = res!minrate1
        mrate2 = res!minrate2
        mrate3 = res!minrate3
        mrate4 = res!minrate4
        gpur = res!gpur
        gpur = gpur + "K"
       
        gpw = res!ngpw
        makingcharges = res!nmaking1
        weight4 = Round(res!nweight4, 2)
        'a = (Val(igw) - (Val(weight1) + Val(weight2) + Val(weight3)) + (Val(igw) - (Val(weight1) + Val(weight2) + Val(weight3))) * Val(gpw) / 100) / 5
        'weight4 = a
        sno = res!chcategory & "-" & code.Text
        qul = res!chquality1
        qul1 = res!chquality2
        qul2 = res!chquality3
        Col = res!chcolor1
        col1 = res!chcolor2
        Col2 = res!chcolor3
        size = res!chsize1
        size1 = res!chsize2
        size2 = res!chsize3
        
        If (res!opicture <> Empty) Then
        Picture1.Visible = True
        Picture1.Picture = LoadPicture(res!opicture)
        Else
        Picture1.Visible = False
        End If
       End If
        res.CLOSE
        End Sub

Private Sub Form_Unload(Cancel As Integer)
con.CLOSE
End Sub

Private Sub makingcharges_Change()
totamt = (Val(mrate1) * Val(weight1)) + (Val(mrate2) * Val(weight2)) + (Val(mrate3) * Val(weight3)) + (Val(mrate4) * Val(weight4)) + Val(makingcharges)
End Sub

Private Sub mcharges_Change()
amt = (Val(rate1) * Val(weight1)) + (Val(rate2) * Val(weight2)) + (Val(rate3) * Val(weight3)) + (Val(rate4) * Val(weight4)) + Val(mcharges)
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

Private Sub Print_Click()
End Sub

Private Sub Next1_Click()
If (code.Text <> Empty) Then
    loading
    Calculate
Else
MsgBox " Enter the valid code"
End If
End Sub

Private Sub Prev_Click()
If (code.Text <> Empty) Then
loading
Calculate
Else
MsgBox ("Enter The Valid Code For Nevigation")
End If
End Sub

Private Sub print1_Click()
printing
End Sub

Private Sub rate1_Change()
amt1 = Val(rate1) * Val(weight1)
amt = (Val(rate1) * Val(weight1)) + (Val(rate2) * Val(weight2)) + (Val(rate3) * Val(weight3)) + (Val(rate4) * Val(weight4)) + Val(mcharges)
End Sub

Private Sub rate2_Change()
amt2 = Val(rate2) * Val(weight2)
amt = (Val(rate1) * Val(weight1)) + (Val(rate2) * Val(weight2)) + (Val(rate3) * Val(weight3)) + (Val(rate4) * Val(weight4)) + Val(mcharges)
End Sub

Private Sub rate3_Change()
amt3 = Val(rate3) * Val(weight3)
amt = (Val(rate1) * Val(weight1)) + (Val(rate2) * Val(weight2)) + (Val(rate3) * Val(weight3)) + (Val(rate4) * Val(weight4)) + Val(mcharges)
End Sub

Private Sub rate4_Change()
amt4 = Val(rate4) * Val(weight4)
amt = (Val(rate1) * Val(weight1)) + (Val(rate2) * Val(weight2)) + (Val(rate3) * Val(weight3)) + (Val(rate4) * Val(weight4)) + Val(mcharges)
End Sub

Private Sub save_Click()
'Dim con1 As New ADODB.Connection
'con1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\lahar\vb\db2.mdb;Persist Security Info=False"
con.Open
res.Open "tblestsalevalue", con1, adOpenDynamic, adLockOptimistic
res.addnew
'////////////
res!ncode = Val(code.Text)
res!dinvoicedate = CDate(Invoicedate)
If (sdate.Text <> Empty) Then
res!sdate = CDate(sdate.Text)
Else
res!sdate = CDate(Date)
End If
res!nweight1 = Val(weight1)
res!nweight2 = Val(weight2)
res!nweight3 = Val(weight3)
res!nweight4 = Val(weight4)
res!nsrate1 = Val(rate1)
res!nsrate2 = Val(rate2)
res!nsrate3 = Val(rate3)
res!nsrate4 = Val(rate4)
res.Update
res.CLOSE

str1 = "select * from tblcosting where ncode=" & Val(code.Text)
res.Open str1, con1, adOpenDynamic, adLockOptimistic
res!ncode = Val(code)
res!chItemname = ""
res!dinvoicedate = ""
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
res!drecdate = ""
res!chissueto = ""
res!ngpw = ""
res!Chstate = "Not Avi."
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
res!opicture = ""
res.Update
res.CLOSE
'con1.close
save.Enabled = False
End Sub



Private Sub totamt_Change()
pno = Round(totamt * 0.0125)
End Sub

