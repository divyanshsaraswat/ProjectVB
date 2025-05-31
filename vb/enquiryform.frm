VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   5715
   ClientLeft      =   240
   ClientTop       =   1545
   ClientWidth     =   6585
   LinkTopic       =   "Form6"
   MDIChild        =   -1  'True
   ScaleHeight     =   5715
   ScaleWidth      =   6585
   WindowState     =   2  'Maximized
   Begin VB.TextBox content1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
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
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   72
      Text            =   "Dia."
      Top             =   2760
      Width           =   615
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
      Left            =   6360
      TabIndex        =   71
      Top             =   4200
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
      Left            =   6360
      TabIndex        =   70
      Top             =   3720
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
      Left            =   6360
      TabIndex        =   69
      Top             =   3240
      Width           =   975
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
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   68
      Top             =   2760
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
      TabIndex        =   67
      Top             =   4200
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
      Left            =   5040
      TabIndex        =   66
      Top             =   3720
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
      Left            =   5040
      TabIndex        =   65
      Top             =   3240
      Width           =   975
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
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   64
      Top             =   2760
      Width           =   975
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
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   63
      Top             =   4200
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
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   62
      Top             =   3720
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
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   61
      Top             =   3240
      Width           =   735
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
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   60
      Top             =   2760
      Width           =   735
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
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   59
      Text            =   "GOLD"
      Top             =   4200
      Width           =   1335
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
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   58
      Top             =   3240
      Width           =   615
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
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   57
      Top             =   3720
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
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   56
      Top             =   3720
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
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   55
      Top             =   3240
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
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   54
      Top             =   2760
      Width           =   615
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
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   53
      Top             =   4200
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
      Left            =   1320
      TabIndex        =   52
      Top             =   3720
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
      Left            =   1320
      TabIndex        =   51
      Top             =   3240
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
      Left            =   1320
      TabIndex        =   50
      Top             =   2760
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
      Left            =   2040
      TabIndex        =   49
      Top             =   3720
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
      Left            =   2040
      TabIndex        =   48
      Top             =   3240
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
      Left            =   2040
      TabIndex        =   47
      Top             =   2760
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
      Left            =   2760
      TabIndex        =   46
      Top             =   3720
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
      Left            =   2760
      TabIndex        =   45
      Top             =   3240
      Width           =   615
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
      Left            =   2760
      TabIndex        =   44
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox date 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox category 
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
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox Code 
      Appearance      =   0  'Flat
      DataField       =   "ncode"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox Itemname1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox igw 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   7320
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1680
      Width           =   1095
   End
   Begin VB.ComboBox gpw 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "enquiryform.frx":0000
      Left            =   9600
      List            =   "enquiryform.frx":0016
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "10%"
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox makingcharges 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   6600
      TabIndex        =   12
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox totamt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   6600
      TabIndex        =   11
      Top             =   5400
      Width           =   975
   End
   Begin VB.TextBox maker 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox issueto 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "C&lose"
      Height          =   375
      Left            =   5640
      TabIndex        =   8
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton load 
      Caption         =   "L&oad"
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton clear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   6120
      Width           =   975
   End
   Begin VB.TextBox smaking 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5520
      TabIndex        =   5
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox GRATE 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   7320
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox stotamt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5520
      TabIndex        =   3
      Top             =   5400
      Width           =   975
   End
   Begin VB.ComboBox state 
      Height          =   315
      ItemData        =   "enquiryform.frx":0036
      Left            =   9240
      List            =   "enquiryform.frx":0040
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Avi."
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox pno 
      Appearance      =   0  'Flat
      DataField       =   "ncode"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "Weight(Ct.)"
      Height          =   375
      Left            =   4200
      TabIndex        =   73
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "Col."
      Height          =   375
      Left            =   1440
      TabIndex        =   43
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "Qt."
      Height          =   375
      Left            =   3000
      TabIndex        =   42
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "Pc."
      Height          =   375
      Left            =   3600
      TabIndex        =   41
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "Sz."
      Height          =   375
      Left            =   2040
      TabIndex        =   40
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000013&
      Caption         =   "Item Code:"
      Height          =   255
      Left            =   120
      TabIndex        =   37
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000013&
      Caption         =   "IName"
      Height          =   255
      Left            =   2280
      TabIndex        =   36
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000013&
      Caption         =   "Mfg. Date"
      Height          =   255
      Left            =   2760
      TabIndex        =   35
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000013&
      Caption         =   "IGW(Ct.)"
      Height          =   255
      Left            =   5640
      TabIndex        =   34
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000013&
      Caption         =   "Gold Wes.@"
      Height          =   255
      Left            =   8520
      TabIndex        =   33
      Top             =   1200
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   8280
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "Stone"
      Height          =   375
      Left            =   480
      TabIndex        =   32
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "Wt.(Ct.)"
      Height          =   375
      Left            =   2760
      TabIndex        =   31
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "Min.Rate1"
      Height          =   375
      Left            =   6480
      TabIndex        =   30
      Top             =   2280
      Width           =   975
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   7680
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000013&
      Caption         =   "Making Charges:"
      Height          =   375
      Left            =   3840
      TabIndex        =   29
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000013&
      Caption         =   "Total Amount:"
      Height          =   375
      Left            =   3840
      TabIndex        =   28
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000013&
      Caption         =   "Maker:"
      Height          =   375
      Left            =   120
      TabIndex        =   27
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000013&
      Caption         =   "Issue To:"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   5400
      Width           =   735
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   11880
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   12000
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
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
      Height          =   855
      Left            =   360
      TabIndex        =   25
      Top             =   0
      Width           =   8895
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   7560
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "SMin. Rate"
      Height          =   375
      Left            =   5160
      TabIndex        =   24
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label17 
      BackColor       =   &H80000013&
      Caption         =   "GRate(24K:10T):"
      Height          =   255
      Left            =   5640
      TabIndex        =   23
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "1."
      Height          =   255
      Left            =   0
      TabIndex        =   22
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "4."
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "3."
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "2."
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label18 
      BackColor       =   &H80000013&
      Caption         =   "S.NO."
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label lState 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "St."
      Height          =   375
      Left            =   8520
      TabIndex        =   17
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "P.No:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1680
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   2895
      Left            =   8400
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   3375
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim res As New ADODB.Recordset
Dim stotamt1 As Double
Dim totamt1 As Double
Dim flag As Boolean
Private Sub calculate()
stotamt = Val(rate1) * Val(weight1) + Val(rate2) * Val(weight2) + Val(rate3) * Val(weight3) + Val(rate4) * Val(weight4) + Val(smaking)
'stotamt1 = Round(Val(stotamt))

totamt = Val(mrate1) * Val(weight1) + Val(mrate2) * Val(weight2) + Val(mrate3) * Val(weight3) + Val(mrate4) * Val(weight4) + Val(makingcharges)
'totamt1 = Round(Val(totamt))
If (Val(usprice) <> 0) Then
amtinus = Round(Val(stotamt) / Val(usprice))
End If
End Sub



Private Sub Clear_Click()
flag = False
'Code = ""
maker = ""
issueto = ""
Category = ""
stotamt = ""
stotamt1 = 0
totamt1 = 0
totamt = ""
content2 = ""
content3 = ""
content4 = ""
igw = 0
gpw = 0
date = ""
pno = ""
makingcharges = 0
mrate1 = 0
mrate2 = 0
mrate3 = 0
mrate4 = 0
rate1 = 0
rate2 = 0
rate3 = 0
rate4 = 0
smaking = 0
weight1 = 0
weight2 = 0
weight3 = 0
weight4 = 0
PC1 = 0
PC2 = 0
PC3 = 0
gpur = 0
End Sub

Private Sub Command5_Click()
Unload Form6
End Sub

Private Sub load_Click()
If (Val(Code.Text) <> Empty) Then
   Clear_Click
   loading
   calculate
   flag = True
Else
MsgBox ("Enter The Valid Code For Nevigation")
End If
End Sub
Private Sub loading()
   str1 = "select * from tblcosting where ncode=" & Val(Code.Text)
    res.Open str1, MDIForm1.con1
    If (res.EOF = True) Then
    MsgBox ("Record Not Found")
    Else
        Code.Text = res!ncode
        Category = res!chcategory
        date = Format(res!dinvoicedate, "dd-mmm-yy")
    
        content2 = res!chcontent2
        content3.Text = res!chcontent3
        content4.Text = res!chcontent4
          
        igw = res!nigw
        weight1 = res!nweight1
        weight2 = res!nweight2
        weight3 = res!nweight3
        PC1 = res!pcs1
        PC2 = res!pcs2
        PC3 = res!pcs3
        mrate1 = res!nrate1
        mrate2 = res!nrate2
        mrate3 = res!nrate3
        mrate4 = res!nrate4
        rate1 = res!minrate1
        rate2 = res!minrate2
        rate3 = res!minrate3
        rate4 = res!minrate4
        gpur = res!gpur
        gpur = gpur + "K"
        gpw = res!ngpw
        makingcharges = res!nmaking
        smaking = res!nmaking1
        weight4 = Round(res!nweight4, 2)
        maker = res!chmaker
        If (IsNull(res!chissueto) = False) Then
        issueto = res!chissueto
        End If
        'a = (Val(igw) - (Val(weight1) + Val(weight2) + Val(weight3)) + (Val(igw) - (Val(weight1) + Val(weight2) + Val(weight3))) * Val(gpw) / 100) / 5
        'weight4 = a
     '   Sno = res!chcategory & "-" & Code.Text
        qul = res!chquality1
        qul1 = res!chquality2
        qul2 = res!chquality3
        col = res!chcolor1
        col1 = res!chcolor2
        Col2 = res!chcolor3
        size = res!chsize1
        size1 = res!chsize2
        size2 = res!chsize3
        
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
       Else
        Image1.Visible = False
        End If
    End If
        res.close
End Sub

Private Sub makingcharges_Change()
If (makingcharges <> Empty) Then
If (flag = True) Then
mrate1 = Round(Val((Val(totamt) - (Val(mrate2) * Val(weight2) + Val(mrate3) * Val(weight3) + Val(mrate4) * Val(weight4) + Val(makingcharges))) / weight1))
End If
End If
End Sub

Private Sub mrate2_Change()
If (mrate2 <> Empty) Then
If (flag = True) Then
mrate1 = Round(Val((Val(totamt) - (Val(mrate2) * Val(weight2) + Val(mrate3) * Val(weight3) + Val(mrate4) * Val(weight4) + Val(makingcharges))) / weight1))
End If
End If
End Sub

Private Sub mrate3_Change()
If (mrate3 <> Empty) Then
If (flag = True) Then
mrate1 = Round(Val((Val(totamt) - (Val(mrate2) * Val(weight2) + Val(mrate3) * Val(weight3) + Val(mrate4) * Val(weight4) + Val(makingcharges))) / weight1))
End If
End If
End Sub

Private Sub mrate4_Change()
If (mrate4 <> Empty) Then
If (flag = True) Then
mrate1 = Round(Val((Val(totamt) - (Val(mrate2) * Val(weight2) + Val(mrate3) * Val(weight3) + Val(mrate4) * Val(weight4) + Val(makingcharges))) / weight1))
End If
End If

End Sub


Private Sub rate2_Change()
If (rate2 <> Empty) Then
If (flag = True) Then
rate1 = Round(Val((Val(stotamt) - (Val(rate2) * Val(weight2) + Val(rate3) * Val(weight3) + Val(rate4) * Val(weight4) + Val(smaking))) / weight1))
End If
End If
End Sub

Private Sub rate3_Change()
If (rate3 <> Empty) Then
If (flag = True) Then
rate1 = Round(Val((Val(stotamt) - (Val(rate2) * Val(weight2) + Val(rate3) * Val(weight3) + Val(rate4) * Val(weight4) + Val(smaking))) / weight1))
End If
End If
End Sub
Private Sub rate4_Change()
If (rate4 <> Empty) Then
If (flag = True) Then
rate1 = Round(Val((Val(stotamt) - (Val(rate2) * Val(weight2) + Val(rate3) * Val(weight3) + Val(rate4) * Val(weight4) + Val(smaking))) / weight1))
End If
End If
End Sub

Private Sub smaking_Change()
If (smaking <> Empty) Then
If (flag = True) Then
rate1 = Round(Val((Val(stotamt) - (Val(rate2) * Val(weight2) + Val(rate3) * Val(weight3) + Val(rate4) * Val(weight4) + Val(smaking))) / weight1))
End If
End If
End Sub

Private Sub stotamt_Change()
If (stotamt <> Empty) Then
If (flag = True) Then
rate1 = Round(Val((Val(stotamt) - (Val(rate2) * Val(weight2) + Val(rate3) * Val(weight3) + Val(rate4) * Val(weight4) + smaking)) / weight1))
stotamt1 = Val(stotamt)
End If
End If
End Sub
Private Sub totamt_change()
If (totamt <> Empty) Then
If (flag = True) Then
mrate1 = Round(Val((Val(totamt) - (Val(mrate2) * Val(weight2) + Val(mrate3) * Val(weight3) + Val(mrate4) * Val(weight4) + Val(makingcharges))) / weight1))
totamt1 = Val(totamt)
End If
End If
End Sub
