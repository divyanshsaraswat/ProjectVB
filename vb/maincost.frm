VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000E&
   Caption         =   "purchase"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   480
   ClientWidth     =   11580
   DrawStyle       =   5  'Transparent
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7215
   ScaleWidth      =   11580
   WindowState     =   2  'Maximized
   Begin VB.TextBox DNO 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2400
      TabIndex        =   96
      Top             =   5400
      Width           =   1095
   End
   Begin VB.ComboBox sowner 
      Height          =   315
      ItemData        =   "maincost.frx":0000
      Left            =   720
      List            =   "maincost.frx":0010
      TabIndex        =   94
      Top             =   1200
      Width           =   615
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   4080
      TabIndex        =   93
      Top             =   1200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yy"
      Format          =   108003329
      CurrentDate     =   42461
   End
   Begin VB.TextBox pno 
      Appearance      =   0  'Flat
      DataField       =   "ncode"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   92
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton delete 
      Appearance      =   0  'Flat
      Caption         =   "Rm. Picture"
      Height          =   375
      Left            =   10440
      TabIndex        =   90
      Top             =   1680
      Width           =   1095
   End
   Begin VB.ComboBox Category 
      Height          =   315
      ItemData        =   "maincost.frx":0024
      Left            =   1320
      List            =   "maincost.frx":004C
      Sorted          =   -1  'True
      TabIndex        =   89
      Top             =   1200
      Width           =   615
   End
   Begin VB.ComboBox content3 
      Height          =   315
      ItemData        =   "maincost.frx":0079
      Left            =   360
      List            =   "maincost.frx":00BF
      TabIndex        =   88
      Top             =   3720
      Width           =   735
   End
   Begin VB.ComboBox content2 
      Height          =   315
      ItemData        =   "maincost.frx":0143
      Left            =   360
      List            =   "maincost.frx":0189
      TabIndex        =   87
      Top             =   3240
      Width           =   735
   End
   Begin VB.ComboBox content1 
      Height          =   315
      ItemData        =   "maincost.frx":020D
      Left            =   360
      List            =   "maincost.frx":0253
      Locked          =   -1  'True
      TabIndex        =   86
      Text            =   "Dia."
      Top             =   2760
      Width           =   735
   End
   Begin VB.ComboBox state 
      Height          =   315
      ItemData        =   "maincost.frx":02D7
      Left            =   9240
      List            =   "maincost.frx":02E1
      Locked          =   -1  'True
      TabIndex        =   84
      Text            =   "Avi."
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox amtinus 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7560
      TabIndex        =   81
      Text            =   "0"
      Top             =   5400
      Width           =   975
   End
   Begin VB.TextBox usprice 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7920
      TabIndex        =   80
      Text            =   "0"
      Top             =   4800
      Width           =   615
   End
   Begin VB.TextBox stotamt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5400
      TabIndex        =   79
      Text            =   "0"
      Top             =   5400
      Width           =   975
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
      Height          =   285
      Left            =   2880
      TabIndex        =   78
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox qul2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   77
      Top             =   3720
      Width           =   735
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
      Height          =   285
      Left            =   2040
      TabIndex        =   76
      Text            =   "0"
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox size2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   75
      Text            =   "0"
      Top             =   3720
      Width           =   735
   End
   Begin VB.ComboBox size 
      Height          =   315
      ItemData        =   "maincost.frx":02F5
      Left            =   2040
      List            =   "maincost.frx":031D
      TabIndex        =   71
      Top             =   2760
      Width           =   735
   End
   Begin VB.ComboBox color2 
      Height          =   315
      ItemData        =   "maincost.frx":0360
      Left            =   1200
      List            =   "maincost.frx":036D
      TabIndex        =   70
      Top             =   3720
      Width           =   735
   End
   Begin VB.ComboBox color1 
      Height          =   315
      ItemData        =   "maincost.frx":037D
      Left            =   1200
      List            =   "maincost.frx":038A
      TabIndex        =   69
      Top             =   3240
      Width           =   735
   End
   Begin VB.ComboBox color 
      Height          =   315
      ItemData        =   "maincost.frx":039A
      Left            =   1200
      List            =   "maincost.frx":03B6
      TabIndex        =   68
      Top             =   2760
      Width           =   735
   End
   Begin VB.ComboBox qul 
      Height          =   315
      ItemData        =   "maincost.frx":03E3
      Left            =   2880
      List            =   "maincost.frx":0408
      TabIndex        =   67
      Top             =   2760
      Width           =   855
   End
   Begin VB.ComboBox content4 
      Height          =   315
      ItemData        =   "maincost.frx":0439
      Left            =   360
      List            =   "maincost.frx":044C
      TabIndex        =   61
      Text            =   "Gold"
      Top             =   4200
      Width           =   1695
   End
   Begin VB.CommandButton printss 
      Caption         =   "Print&SS"
      Height          =   375
      Left            =   6840
      TabIndex        =   60
      Top             =   6120
      Width           =   975
   End
   Begin VB.TextBox GRATE 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   7320
      TabIndex        =   59
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox smaking 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5400
      TabIndex        =   57
      Text            =   "0"
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton ADD 
      Appearance      =   0  'Flat
      Caption         =   "Add Picture"
      Height          =   375
      Left            =   10440
      TabIndex        =   56
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton clear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   8040
      TabIndex        =   55
      Top             =   6120
      Width           =   975
   End
   Begin VB.TextBox minrate4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5400
      TabIndex        =   19
      Text            =   "0"
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox minrate3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5400
      TabIndex        =   53
      Text            =   "0"
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox minrate2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5400
      TabIndex        =   12
      Text            =   "0"
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox minrate1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5400
      TabIndex        =   8
      Text            =   "0"
      Top             =   2760
      Width           =   975
   End
   Begin VB.ComboBox gpur 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "maincost.frx":047B
      Left            =   2280
      List            =   "maincost.frx":049E
      TabIndex        =   16
      Text            =   "18K"
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox pcs3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3840
      TabIndex        =   13
      Text            =   "0"
      Top             =   3720
      Width           =   615
   End
   Begin VB.TextBox pcs2 
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
      Height          =   285
      Left            =   3840
      TabIndex        =   9
      Text            =   "0"
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox pcs1 
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
      Height          =   285
      Left            =   3840
      TabIndex        =   5
      Text            =   "0"
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton SisterConcern 
      Caption         =   "SaleSheet"
      Height          =   375
      Left            =   5640
      TabIndex        =   49
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton Addnew 
      Caption         =   "AddNew"
      Height          =   375
      Left            =   120
      TabIndex        =   48
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton Prev 
      Caption         =   "Previous"
      Height          =   375
      Left            =   3960
      TabIndex        =   47
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton Next1 
      Caption         =   "Next"
      Height          =   375
      Left            =   5400
      TabIndex        =   46
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton Last 
      Caption         =   "Last"
      Height          =   375
      Left            =   6840
      TabIndex        =   45
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton First 
      Caption         =   "First"
      Height          =   375
      Left            =   2520
      TabIndex        =   44
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "L&oad"
      Height          =   375
      Left            =   2400
      TabIndex        =   43
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "C&lose"
      Height          =   375
      Left            =   9240
      TabIndex        =   42
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Update"
      Height          =   375
      Left            =   3480
      TabIndex        =   41
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton dele 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   4560
      TabIndex        =   40
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Height          =   375
      Left            =   1200
      TabIndex        =   23
      Top             =   6120
      Width           =   1095
   End
   Begin VB.TextBox pcode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   22
      Top             =   5400
      Width           =   975
   End
   Begin VB.TextBox maker 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   20
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox weight1 
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
      Height          =   285
      Left            =   4560
      TabIndex        =   6
      Text            =   "0"
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox weight2 
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
      Height          =   285
      Left            =   4560
      TabIndex        =   10
      Text            =   "0"
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox weight3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4560
      TabIndex        =   14
      Text            =   "0"
      Top             =   3720
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
      Height          =   285
      Left            =   4560
      TabIndex        =   17
      Text            =   "0"
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox rate1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   6480
      TabIndex        =   7
      Text            =   "0"
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox rate2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   6480
      TabIndex        =   11
      Text            =   "0"
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox rate3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   6480
      TabIndex        =   15
      Text            =   "0"
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox rate4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   6480
      TabIndex        =   18
      Text            =   "0"
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox amt1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Left            =   7560
      TabIndex        =   52
      Text            =   "0"
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox amt2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Left            =   7560
      TabIndex        =   24
      Text            =   "0"
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox amt3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Left            =   7560
      TabIndex        =   25
      Text            =   "0"
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox amt4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Left            =   7560
      TabIndex        =   26
      Text            =   "0"
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox totamt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Left            =   6480
      TabIndex        =   0
      Text            =   "0"
      Top             =   5400
      Width           =   975
   End
   Begin VB.TextBox makingcharges 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   6480
      TabIndex        =   21
      Text            =   "0"
      Top             =   4800
      Width           =   975
   End
   Begin VB.ComboBox gpw 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "maincost.frx":04C9
      Left            =   9600
      List            =   "maincost.frx":04DF
      TabIndex        =   4
      Text            =   "0%"
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox igw 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   7320
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox Itemname1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3120
      TabIndex        =   3
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox Code 
      Appearance      =   0  'Flat
      DataField       =   "ncode"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   1200
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5040
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open"
      InitDir         =   "c:\lahar\images"
   End
   Begin VB.Label Label25 
      BackColor       =   &H8000000E&
      Caption         =   "NO."
      Height          =   255
      Left            =   2040
      TabIndex        =   95
      Top             =   5400
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   2775
      Left            =   8880
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   2895
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "P.No:"
      Height          =   255
      Left            =   120
      TabIndex        =   91
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lState 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "St."
      Height          =   375
      Left            =   8520
      TabIndex        =   85
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "$@"
      Height          =   255
      Left            =   7440
      TabIndex        =   83
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Rate in US $"
      Height          =   255
      Left            =   7320
      TabIndex        =   82
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Col."
      Height          =   375
      Left            =   1440
      TabIndex        =   73
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Qt."
      Height          =   375
      Left            =   3000
      TabIndex        =   72
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label18 
      BackColor       =   &H8000000E&
      Caption         =   "S.NO."
      Height          =   375
      Left            =   0
      TabIndex        =   66
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "2."
      Height          =   255
      Left            =   0
      TabIndex        =   65
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "3."
      Height          =   255
      Left            =   0
      TabIndex        =   64
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "4."
      Height          =   255
      Left            =   0
      TabIndex        =   63
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "1."
      Height          =   255
      Left            =   0
      TabIndex        =   62
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label17 
      BackColor       =   &H8000000E&
      Caption         =   "GRate(24K:10T):"
      Height          =   255
      Left            =   5640
      TabIndex        =   58
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "SMin. Rate"
      Height          =   375
      Left            =   5400
      TabIndex        =   54
      Top             =   2280
      Width           =   975
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   8400
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Pc."
      Height          =   375
      Left            =   3720
      TabIndex        =   51
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
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
      TabIndex        =   50
      Top             =   0
      Width           =   8895
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   12000
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   11880
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Label Label13 
      BackColor       =   &H8000000E&
      Caption         =   "P.Code"
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label Label12 
      BackColor       =   &H8000000E&
      Caption         =   "Maker:"
      Height          =   375
      Left            =   120
      TabIndex        =   38
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label Label11 
      BackColor       =   &H8000000E&
      Caption         =   "Total Amount:"
      Height          =   375
      Left            =   3840
      TabIndex        =   37
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label Label10 
      BackColor       =   &H8000000E&
      Caption         =   "Making Charges:"
      Height          =   375
      Left            =   3840
      TabIndex        =   36
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   8400
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Amount"
      Height          =   375
      Left            =   7440
      TabIndex        =   35
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Min.Rate1"
      Height          =   375
      Left            =   6480
      TabIndex        =   34
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Wt.(Ct.)"
      Height          =   375
      Left            =   4320
      TabIndex        =   33
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Stone"
      Height          =   375
      Left            =   480
      TabIndex        =   32
      Top             =   2280
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   8280
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "Gold Wes.@"
      Height          =   255
      Left            =   8520
      TabIndex        =   31
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "IGW(Ct.)"
      Height          =   255
      Left            =   5640
      TabIndex        =   30
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Caption         =   "  Date"
      Height          =   255
      Left            =   2760
      TabIndex        =   29
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "IName"
      Height          =   255
      Left            =   2280
      TabIndex        =   28
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "I Code:"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Sz."
      Height          =   375
      Left            =   2040
      TabIndex        =   74
      Top             =   2280
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con1 As New ADODB.Connection
Dim res1 As New ADODB.Recordset
Dim res As New ADODB.Recordset
Dim sas As Integer
Dim loadingstate As Integer
Dim picturepath As String
Dim status As Integer
Private Sub add_Click()
On Error GoTo errorhandler
CommonDialog1.DialogTitle = "open"
CommonDialog1.ShowOpen

Dim picturename As String
pos = InStrRev(CommonDialog1.FileName, "\")
If (pos <> 0) Then
picturename = Mid(CommonDialog1.FileName, pos + 1)
End If
If (picturename <> Empty) Then
Image1.Visible = True
Image1.Picture = LoadPicture(MDIForm1.picturepath & picturename)
picturepath = picturename

'Image1.Picture = LoadPicture(picturepath)
Image1.Visible = True
End If
errorhandler:
If (Err.Number = 53) Then
MsgBox ("No Image Found")
Exit Sub
End If
End Sub

Private Sub Addnew_Click()
On Error Resume Next
clearing
Image1.Visible = True
Image1.Picture = LoadPicture()

'If (Category.Text = Empty) Then
'MsgBox ("Please Select the Category First")
'End Sub
'End If

str1 = "select max(ncode)+1 from tblcosting"
If (res1.state = adStateOpen) Then
res1.CLOSE
End If
res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
If (IsNull(res1.Fields(0)) Or res1.Fields(0) < 1000) Then
res1.CLOSE
res1.Open "select max(ncode) from tblcosting ", MDIForm1.con1, adOpenDynamic, adLockOptimistic
Command1.Enabled = True
Command4.Enabled = False
Next1.Enabled = False

If (IsNull(res1.Fields(0)) Or res1.Fields(0) < 1000) Then
Code.Text = 1001
Else
Code = res1.Fields(0) + 1
End If
Else
Code = res1.Fields(0)
Command1.Enabled = True
Command4.Enabled = False
End If
state = "Avi."
res1.CLOSE

DTPicker1.Value = Date
'Invoicedate.Text = Format(Date, "dd-mmm-yy")

'Prev.Enabled = False
'Next1.Enabled = False
'Command1.Enabled = True
'Command2.Enabled = False
dele.Enabled = False
'Command4.Enabled = False
Command5.Enabled = True
Command6.Enabled = False
Addnew.Enabled = False
sowner.SetFocus
End Sub



Private Sub Category_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
        Code.SetFocus
End If
End Sub

Private Sub Clear_Click()
Command6.Enabled = True
Command4.Enabled = False
Command1.Enabled = False
Addnew.Enabled = True
Code.Enabled = True
Image1.Visible = False
clearing
End Sub





Private Sub Code_GotFocus()
Prev.Enabled = True
Next1.Enabled = True
End Sub

Private Sub code_KeyPress(KeyAscii As Integer)
 If (KeyAscii = 13) Then
    If (Addnew.Enabled = True) Then
    Command6_Click
   ' Itemname1.SetFocus
    'Else
   ' Itemname1.SetFocus
  End If
End If
End Sub

Private Sub Command1_Click()
If (igw.Text = "" Or Code.Text = "" Or Category.Text = "" Or sowner = "") Then
        MsgBox ("Enter The Correct Entry For Item Code or IGW(Ct.) or Category Entry or Owner Entry")
        Exit Sub
ElseIf ((content2.Text <> "" And Val(weight2) <= 0) Or (content2.Text = "" And Val(weight2) > 0)) Then
        MsgBox ("Enter the Valid Entry For Stone 2")
        Exit Sub
ElseIf ((content3.Text <> "" And Val(weight3) <= 0) Or (content3.Text = "" And Val(weight3) > 0)) Then
        MsgBox ("Enter the Valid Entry For Stone 3")
        Exit Sub
ElseIf (maker.Text = "" Or Val(smaking) <= 0) Then
        MsgBox ("Enter The Valid Entry For Maker Or Making Charges")
        Exit Sub
End If


First.Enabled = True
Prev.Enabled = True
Last.Enabled = True
Next1.Enabled = True
Command1.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Addnew.Enabled = True
dele.Enabled = True
Next1.Enabled = True
res.Open "tblcosting", MDIForm1.con1, adOpenDynamic, adLockOptimistic
res.Addnew
res!ncode = Val(Code)
res!chItemname = Itemname1
res!chowner = sowner
'res!dinvoicedate = CDate(Invoicedate)
res!dinvoicedate = CDate(DTPicker1)
res!chcontent1 = content1
res!chcontent2 = content2
res!chcontent3 = content3
res!chcontent4 = content4
res!nweight1 = Val(weight1)
res!nweight2 = Val(weight2)
res!nweight3 = Val(weight3)
res!nweight4 = Val(weight4)
res!nrate1 = Val(rate1)
res!nrate2 = Val(rate2)
res!nrate3 = Val(rate3)
res!nrate4 = Val(rate4)
res!nigw = Val(igw)
res!nmaking = Val(makingcharges)
res!chmaker = maker
'res!drecdate = CDate(Invoicedate)
res!drecdate = CDate(DTPicker1)
res!chissueto = issueto
res!ngpw = Val(gpw)
res!pcs1 = Val(pcs1)
res!pcs2 = Val(pcs2)
res!pcs3 = Val(pcs3)
res!gpur = Val(gpur)
res!minrate1 = Val(minrate1)
res!minrate2 = Val(minrate2)
res!minrate3 = Val(minrate3)
res!minrate4 = Val(minrate4)
res!nmaking1 = Val(smaking)
res!chquality1 = qul
res!chquality2 = qul1
res!chquality3 = qul2
res!chcolor1 = color
res!chcolor2 = color1
res!chcolor3 = color2
res!chsize1 = size
res!chsize2 = size1
res!chsize3 = size2
res!chstate = state
res!chcategory = Category
res!chpcode = pcode.Text
res!ndno = DNO.Text
If (picturepath <> Empty) Then
    res!opicture = picturepath
    picturepath = ""
End If
res.Update
res.CLOSE
res.Open "tblptag", MDIForm1.con1, adOpenDynamic, adLockOptimistic
res.Addnew
res!chcategory = Category
res!ncode = Val(Code)
res.Update
res.CLOSE
MsgBox ("Record Saved")
Command1.Enabled = False
End Sub
Private Sub Command2_Click()
    Dim W As Integer
    Dim H As Integer
    Dim X As Integer
    W = Form1.ScaleWidth / 2
    H = Form1.ScaleHeight / 2
    Form1.Command1.Visible = False
   ' Form1.Command2.Visible = False
    Form1.Command4.Visible = False
    Form1.Command5.Visible = False
    Form1.Command6.Visible = False
    Form1.Addnew.Visible = False
    Form1.SisterConcern.Visible = False
    
   ' X = SetStretchBltMode(Form1.hDC, 3)
   ' X = StretchBlt(Form1.hDC, 0, 0, W, H, Form1.hDC, 0, 0, Form1.ScaleWidth, Form1.ScaleHeight, SRCCOPY)
   ' X = StretchBlt(Form1.hDC, W, 0, W, H, Form1.hDC, 0, 0, Form1.ScaleWidth, Form1.ScaleHeight, SRCCOPY)
   ' X = StretchBlt(Form1.hDC, W, H, W, H, Form1.hDC, 0, 0, Form1.ScaleWidth, Form1.ScaleHeight, SRCCOPY)
   ' X = StretchBlt(Form1.hDC, 0, H, W, H, Form1.hDC, 0, 0, Form1.ScaleWidth, Form1.ScaleHeight, SRCCOPY)
    Form1.Refresh
    Form1.PrintForm
    Form1.Command1.Visible = True
   ' Form1.Command2.Visible = True
    Form1.Command4.Visible = True
    Form1.Command5.Visible = True
    Form1.Addnew.Visible = False
    Form1.SisterConcern.Visible = True
    Form1.Command6.Visible = True
    
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()
If (status <> 1) Then

If (igw.Text = "" Or Code.Text = "" Or Category.Text = "" Or sowner = "") Then
        MsgBox ("Enter The Correct Entry For Item Code or IGW(Ct.) or Category Entry or Owner Enrty")
        Exit Sub
ElseIf ((content2.Text <> "" And Val(weight2) <= 0) Or (content2.Text = "" And Val(weight2) > 0)) Then
        MsgBox ("Enter the Valid Entry For Stone 2")
        Exit Sub
ElseIf ((content3.Text <> "" And Val(weight3) <= 0) Or (content3.Text = "" And Val(weight3) > 0)) Then
        MsgBox ("Enter the Valid Entry For Stone 3")
        Exit Sub
ElseIf (maker.Text = "" Or Val(smaking) <= 0) Then
        MsgBox ("Enter The Valid Entry For Maker Or Making Charges")
        Exit Sub
End If
End If

str1 = "select * from tblcosting where ncode=" & Val(Code.Text)
res.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic

res!ncode = Val(Code)
res!chItemname = Itemname1
res!chowner = sowner
'res!dinvoicedate = CDate(Invoicedate)
'MsgBox (DTPicker1)
res!dinvoicedate = CDate(DTPicker1)
res!chcontent1 = content1
res!chcontent2 = content2
res!chcontent3 = content3
res!chcontent4 = content4
res!chpcode = pcode
res!nweight1 = Val(weight1)
res!nweight2 = Val(weight2)
res!nweight3 = Val(weight3)
res!nweight4 = Val(weight4)
res!nrate1 = Val(rate1)
res!nrate2 = Val(rate2)
res!nrate3 = Val(rate3)
res!nrate4 = Val(rate4)
res!nigw = Val(igw)
res!nmaking = Val(makingcharges)
res!nmaking1 = Val(smaking)
res!chmaker = maker
res!drecdate = CDate(DTPicker1)
res!chissueto = issueto
res!ngpw = Val(gpw)
res!ndno = Val(DNO.Text)

If (status = 1) Then
    res!chstate = "Not Avi."
    state = "Not Avi."
Else
    res!chstate = "Avi."
    state = "Avi."
End If

res!pcs1 = Val(pcs1)
res!pcs2 = Val(pcs2)
res!pcs3 = Val(pcs3)
res!gpur = Val(gpur)
res!minrate1 = Val(minrate1)
res!minrate2 = Val(minrate2)
res!minrate3 = Val(minrate3)
res!minrate4 = Val(minrate4)
res!nmaking1 = Val(smaking)
res!chquality1 = qul
res!chquality2 = qul1
res!chquality3 = qul2
res!chcolor1 = color
res!chcolor2 = color1
res!chcolor3 = color2
res!chsize1 = size
res!chsize2 = size1
res!chsize3 = size2
res!chcategory = Category

If (picturepath <> Empty) Then
    res!opicture = picturepath
    'picturepath = ""
Else
    res!opicture = ""
End If
res.Update
res.CLOSE
If (status <> 1) Then
result = MsgBox("Do You Want Add This Item In Pending Tag List!", vbYesNo, "Confirmation")
    If (Val(result) = 7) Then
        MsgBox ("Record Updated")
    Else
        res.Open "tblptag", MDIForm1.con1, adOpenDynamic, adLockOptimistic
        res.Addnew
        res!chcategory = Category
        res!ncode = Val(Code)
        res.Update
        res.CLOSE
        MsgBox ("Record Updated And Added This Item In The Tag Pending List")
    End If
    Else
        MsgBox ("Record Updated")
    End If
status = 0
Command6.Enabled = True
Addnew.Enabled = True
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Command6_Click()
'loading data

   
str1 = "select * from tblcosting where ncode=" & Val(Code.Text)

res.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
If (res.EOF = True) Then
 MsgBox ("Record Not Found")
res.CLOSE
Else
loading
res.CLOSE
Command1.Enabled = False
dele.Enabled = True
Command4.Enabled = True
First.Enabled = True
Prev.Enabled = True
Last.Enabled = True
Next1.Enabled = True
End If
End Sub

Private Sub content2_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
        weight2.SetFocus
End If

End Sub

Private Sub content3_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
        weight3.SetFocus
End If

End Sub

Private Sub dele_Click()
status = 1


    str1 = "Select chissue,nappno from tblappmaster where nappno="
    str1 = str1 & " ( Select  nappno from tblappdetail where ncode=" & Val(Code.Text) & ") "
    res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
    'MsgBox (str1)
    name1 = ""
    If (res1.EOF = False) Then
     result = MsgBox("This Item Found in the Approval No " & res1!nappno & " And Issue to " & res1!chissue & ". Do You Want To Delete This? ", vbYesNo + vbDefaultButton2, "Confirmation")

    If (Val(result) = 7) Then
        res1.CLOSE
       ' code.Text = ""
        Code.SetFocus
      '  Clear_Click
        Exit Sub
     Else
        str1 = " Select  * from tblappdetail where ncode=" & Val(Code.Text)
        res1.CLOSE
        res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
        res1.delete
        res1.CLOSE
     End If
   Else
    res1.CLOSE
  ' res1.Close
  End If
    
Clear_Click
state = "Not Avi."
Command4_Click
End Sub
Private Sub delete_Click()
str1 = "Select opicture from tblcosting where ncode= " & Val(Code.Text)
res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
res1!opicture = ""
res1.Update
res1.CLOSE
Image1.Visible = False
End Sub


Private Sub DNO_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
        If (Command1.Enabled = True) Then
        Command1.SetFocus
        Else
        Command4.SetFocus
        End If
End If
End Sub

Private Sub First_Click()
res.MoveFirst
loading
First.Enabled = False
Prev.Enabled = False
Next1.Enabled = True
Last.Enabled = True
End Sub

Private Sub Form_Activate()
Code.SetFocus
End Sub

Private Sub Form_Load()
'con1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\lahar\vb\db2.mdb;Persist Security Info=False"
'con1.Open
DTPicker1.Format = dtpCustom
CommonDialog1.InitDir = MDIForm1.picturepath
Command1.Enabled = False
dele.Enabled = False
Command4.Enabled = False
'Prev.Enabled = False
'Next1.Enabled = False
status = 0
clearing

End Sub
Private Sub Form_Unload(Cancel As Integer)
'con1.close
End Sub
Private Sub gpur_LostFocus()
Rateg = (Val(GRATE) * Val(gpur)) / (116.64 * 24)
rate4 = Round(Rateg)
End Sub

Private Sub GRATE_LostFocus()
Rateg = Round((Val(GRATE) * Val(gpur)) / (116.64 * 24))
rate4 = Rateg
End Sub





Private Sub igw_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
        weight1.SetFocus
End If
End Sub

Private Sub Itemname1_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
        igw.SetFocus
End If
End Sub


Private Sub Label24_Click()

End Sub

Private Sub Last_Click()
res.MoveLast
loading
Next1.Enabled = False
Last.Enabled = False
Prev.Enabled = True
First.Enabled = True
End Sub


Private Sub maker_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
        pcode.SetFocus
End If
End Sub

Private Sub makingcharges_Change()
    totamt = Val(amt1) + Val(amt2) + Val(amt3) + Val(amt4) + Val(makingcharges)
End Sub

Private Sub minrate1_Change()
 stotamt = Val(minrate1) * Val(weight1) + Val(minrate2) * Val(weight2) + Val(minrate3) * Val(weight3) + Val(minrate4) * Val(weight4) + Val(smaking)
pcode = Round(Val(minrate1.Text) / 225)
End Sub

Private Sub minrate1_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
        content2.SetFocus
End If
End Sub

Private Sub minrate2_Change()
stotamt = Val(minrate1) * Val(weight1) + Val(minrate2) * Val(weight2) + Val(minrate3) * Val(weight3) + Val(minrate4) * Val(weight4) + Val(smaking)
End Sub

Private Sub minrate2_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
        content3.SetFocus
End If

End Sub

Private Sub minrate3_Change()
stotamt = Val(minrate1) * Val(weight1) + Val(minrate2) * Val(weight2) + Val(minrate3) * Val(weight3) + Val(minrate4) * Val(weight4) + Val(smaking)
End Sub

Private Sub minrate3_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
        weight4.SetFocus
End If
End Sub

Private Sub minrate4_Change()
a = (Val(igw) - (Val(weight1) + Val(weight2) + Val(weight3)) + (Val(igw) - (Val(weight1) + Val(weight2) + Val(weight3))) * Val(gpw) / 100) / 5
weight4 = a
stotamt = Val(minrate1) * Val(weight1) + Val(minrate2) * Val(weight2) + Val(minrate3) * Val(weight3) + Val(minrate4) * Val(weight4) + Val(smaking)
End Sub

Private Sub minrate4_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
        smaking.SetFocus
End If

End Sub

Private Sub Next1_Click()
If (Code.Text <> Empty) Then
str1 = "select * from tblcosting where ncode=" & Val(Code.Text) + 1
res.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
If (res.EOF = True) Then
 MsgBox ("End Of File")
res.CLOSE
Else
loading
res.CLOSE
End If
Else
MsgBox ("Enter The Valid Code For Nevigation")
End If
End Sub







Private Sub pcode_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
        DNO.SetFocus
End If
End Sub





Private Sub pno_Change()

End Sub

Private Sub Prev_Click()
If (Code.Text <> Empty) Then
str1 = "select * from tblcosting where ncode=" & Val(Code.Text) - 1
res.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
If (res.EOF = True) Then
 MsgBox ("Bigning Of File")
res.CLOSE
Else
loading
res.CLOSE
Next1.Enabled = True
dele.Enabled = True
End If
Else
MsgBox ("Enter The Valid Code For Nevigation")
End If

End Sub

Private Sub calculate()
amt1 = Val(rate1) * Val(weight1)
amt2 = Val(rate2) * Val(weight2)
amt3 = Val(rate3) * Val(weight3)
'a = Val(igw) - (Val(weight1) + Val(weight2) + Val(weight3)) + (Val(igw) - (Val(weight1) + Val(weight2) + Val(weight3))) * Val(gpw) / 100
'weight4 = a
amt4 = Val(rate4) * Val(weight4)
totamt = Val(amt1) + Val(amt2) + Val(amt3) + Val(amt4) + Val(makingcharges)
stotamt = Val(minrate1) * Val(weight1) + Val(minrate2) * Val(weight2) + Val(minrate3) * Val(weight3) + Val(minrate4) * Val(weight4) + Val(smaking)
pcode.Text = Round(Val(minrate1) / 225)
If (Val(usprice) <> 0) Then
amtinus = Round(Val(stotamt) / Val(usprice))
End If
End Sub
Private Sub loading()
On Error GoTo errorhandler
clearing
Code = res!ncode
stritemname = res!chItemname
If (stritemname = Empty) Then
    Itemname1 = " "
Else
    Itemname1 = res!chItemname
End If

'Invoicedate = Format(res!dinvoicedate, "dd-mmm-yy")
DTPicker1 = res!dinvoicedate
content1 = res!chcontent1
If (IsNull(res!chowner) = False) Then
   sowner = res!chowner
End If

If (IsNull(res!chpcode) = False) Then
   pcode.Text = res!chpcode
Else
   pcode.Text = Round(res!minrate / 225)
End If

If (IsNull(res!chcontent2)) Then
    content2.Text = ""
Else
    content2 = res!chcontent2
End If

If (IsNull(res!chcontent3)) Then
    content3.Text = ""
Else
    content3 = res!chcontent3
End If

content4 = res!chcontent4
weight1 = res!nweight1
weight2 = res!nweight2

If (res!nweight3 = Null) Then
     weight3 = 0
Else
    weight3 = res!nweight3
End If
weight4 = res!nweight4
rate1 = res!nrate1
rate2 = res!nrate2
If (res!nrate1 = Empty) Then
rate3 = 0
Else
rate3 = res!nrate3
End If
rate4 = res!nrate4
igw = res!nigw
makingcharges = res!nmaking
maker = res!chmaker
'Invoicedate = Format(res!drecdate, "dd-mmm-yy")
issueto = res!chissueto
state = res!chstate
gpw = res!ngpw
pcs1 = res!pcs1
pcs2 = res!pcs2
pcs3 = res!pcs3
minrate1 = res!minrate1
minrate2 = res!minrate2
minrate3 = res!minrate3
minrate4 = res!minrate4
smaking = res!nmaking1
gpur = res!gpur
Category = res!chcategory


If (IsNull(res!chquality1)) Then
qul = ""
Else
qul = res!chquality1
End If


If (IsNull(res!chquality2)) Then
qul1 = ""
Else
qul1 = res!chquality2
End If


If (IsNull(res!chquality3)) Then
qul2 = ""
Else
qul2 = res!chquality3
End If


If (IsNull(res!chcolor1)) Then
color = ""
Else
color = res!chcolor1
End If


If (IsNull(res!chcolor2)) Then
color1 = ""
Else
color1 = res!chcolor2
End If


If (IsNull(res!chcolor3)) Then
color2 = ""
Else
color2 = res!chcolor3
End If


If (IsNull(res!chsize1)) Then
size = ""
Else
size = res!chsize1
End If


If (IsNull(res!chsize2)) Then
size1 = ""
Else
size1 = res!chsize2
End If

If (IsNull(res!chsize3)) Then
size2 = ""
Else
size2 = res!chsize3
End If

If (IsNull(res!ndno) = False) Then
     DNO.Text = res!ndno
     Else
     DNO.Text = 0
End If

If (res!opicture <> Empty) Then
Dim picturename As String

pos = InStrRev(res!opicture, "\")
If (pos <> 0) Then
picturename = Mid(res!opicture, pos + 1)
picturepath = picturename
Else
picturename = res!opicture
picturepath = picturename
End If
Image1.Visible = True
Image1.Picture = LoadPicture(MDIForm1.picturepath & picturename)
Else
picturepath = ""
Image1.Visible = False
End If
calculate
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
Private Sub clearing()
Itemname1.Text = ""
'Invoicedate = Format(Date, "dd-mmm-yy")
Category = ""
sowner = ""
content2 = ""
content3 = ""
If (status <> 1) Then
Category = ""
End If
maker = ""
issueto = ""
weight1 = 0
weight2 = 0
weight3 = 0
weight4 = 0
rate1 = 0
rate2 = 0
rate3 = 0

pcs1 = 0
pcs2 = 0
pcs3 = 0
igw = 0
DNO = 0
minrate1 = 0
minrate2 = 0
minrate3 = 0
pcode = ""
makingcharges = 0
smaking = 0

amtinus = 0
size = 0
size1 = 0
size2 = 0
picturepath = ""

calculate
End Sub

Private Sub printss_Click()
Form3.Code.Text = Form1.Code
Form3.loading
Form3.calculate
Form3.printing
Unload Form3
'If DataEnvironment2.rsCommand7.state = adStateOpen Then
'DataEnvironment2.rsCommand7.close
'End If
'If (code.Text = "") Then
'MsgBox "Enter The Code No."
'Exit Sub
'End If
'DataEnvironment2.Commands(7).Parameters(0).Value = Val(code.Text)

'DataReport6.Show
'DataReport6.PrintReport True
End Sub

Private Sub rate1_Change()
  amt1 = Val(rate1) * Val(weight1)
  totamt = Val(amt1) + Val(amt2) + Val(amt3) + Val(amt4) + Val(makingcharges)
End Sub

Private Sub rate2_Change()
  amt2 = Val(rate2) * Val(weight2)
  totamt = Val(amt1) + Val(amt2) + Val(amt3) + Val(amt4) + Val(makingcharges)
  
End Sub

Private Sub rate3_Change()
amt3 = Val(rate3) * Val(weight3)
totamt = Val(amt1) + Val(amt2) + Val(amt3) + Val(amt4) + Val(makingcharges)
End Sub

Private Sub rate4_Change()
amt4 = Val(rate4) * Val(weight4)
totamt = Val(amt1) + Val(amt2) + Val(amt3) + Val(amt4) + Val(makingcharges)
End Sub

Private Sub rate4_GotFocus()
amt4 = Val(rate4) * Val(weight4)
totamt = Val(amt1) + Val(amt2) + Val(amt3) + Val(amt4) + Val(makingcharges)
End Sub

Private Sub SisterConcern_Click()
load Form3
Form3.Code.Text = Form1.Code
Form3.loading
Form3.calculate
Form3.Show
End Sub







Private Sub smaking_Change()
stotamt = Val(minrate1) * Val(weight1) + Val(minrate2) * Val(weight2) + Val(minrate3) * Val(weight3) + Val(minrate4) * Val(weight4) + Val(smaking)
End Sub



Private Sub smaking_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
        maker.SetFocus
End If
End Sub

Private Sub sowner_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
        Category.SetFocus
End If
End Sub

Private Sub stotamt_Change()
pno = Round(stotamt * 0.01333333)
End Sub

Private Sub Text1_Change()

End Sub

Private Sub usprice_Change()
If (Val(usprice) <> 0) Then
amtinus = Round(Val(stotamt) / Val(usprice))
End If
End Sub
Private Sub usprice_GotFocus()
If (Val(usprice) <> 0) Then
amtinus = Round(Val(stotamt) / Val(usprice))
End If
End Sub
Private Sub weight1_Change()
amt1 = Val(rate1) * Val(weight1)
stotamt = Val(minrate1) * Val(weight1) + Val(minrate2) * Val(weight2) + Val(minrate3) * Val(weight3) + Val(minrate4) * Val(weight4) + Val(smaking)
End Sub


Private Sub weight1_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
        minrate1.SetFocus
End If

End Sub

Private Sub weight2_Change()
amt2 = Val(rate2) * Val(weight2)
stotamt = (Val(minrate1) * Val(weight1)) + (Val(minrate2) * Val(weight2)) + (Val(minrate3) * Val(weight3)) + (Val(minrate4) * Val(weight4)) + Val(smaking)
End Sub

Private Sub weight2_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
        minrate2.SetFocus
End If

End Sub

Private Sub weight3_Change()
amt3 = Val(rate3) * Val(weight4)
stotamt = Val(minrate1) * Val(weight1) + Val(minrate2) * Val(weight2) + Val(minrate3) * Val(weight3) + Val(minrate4) * Val(weight4) + Val(smaking)
End Sub

Private Sub weight3_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
        minrate3.SetFocus
End If

End Sub

Private Sub weight4_GotFocus()
a = (Val(igw) - (Val(weight1) + Val(weight2) + Val(weight3)) + (Val(igw) - (Val(weight1) + Val(weight2) + Val(weight3))) * Val(gpw) / 100) / 5
weight4 = a
End Sub

Private Sub weight4_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
        minrate4.SetFocus
End If
End Sub
