VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form dsails 
   Caption         =   "View Details"
   ClientHeight    =   5715
   ClientLeft      =   1815
   ClientTop       =   1860
   ClientWidth     =   6585
   LinkTopic       =   "Form8"
   MDIChild        =   -1  'True
   ScaleHeight     =   8340
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton export 
      Caption         =   "&Export"
      Height          =   375
      Left            =   11640
      TabIndex        =   72
      Top             =   1080
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "H"
      Height          =   255
      Index           =   1
      Left            =   7560
      TabIndex        =   71
      Top             =   720
      Width           =   495
   End
   Begin VB.OptionButton Option1 
      Caption         =   "C"
      Height          =   255
      Index           =   0
      Left            =   6960
      TabIndex        =   70
      Top             =   720
      Value           =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton remove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   9000
      TabIndex        =   69
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox scode 
      Height          =   315
      Left            =   1080
      TabIndex        =   66
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox bal 
      Enabled         =   0   'False
      Height          =   325
      Left            =   5520
      TabIndex        =   58
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox recdate3 
      Height          =   315
      Left            =   7320
      TabIndex        =   55
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox recdate2 
      Height          =   315
      Left            =   7320
      TabIndex        =   54
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox payrec3 
      Height          =   325
      Left            =   5520
      TabIndex        =   50
      Top             =   2040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox payrec2 
      Height          =   325
      Left            =   5520
      TabIndex        =   46
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox total 
      Enabled         =   0   'False
      Height          =   325
      Left            =   3000
      TabIndex        =   43
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox mak 
      Height          =   325
      Left            =   3000
      TabIndex        =   41
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox mrt 
      Height          =   325
      Left            =   3000
      TabIndex        =   38
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox mwt 
      Height          =   325
      Left            =   1200
      TabIndex        =   37
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox crt2 
      Height          =   325
      Left            =   3000
      TabIndex        =   33
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox cst2 
      Height          =   325
      Left            =   1200
      TabIndex        =   32
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox crt1 
      Height          =   325
      Left            =   3000
      TabIndex        =   29
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox cst1 
      Height          =   325
      Left            =   1200
      TabIndex        =   28
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox drt 
      Height          =   325
      Left            =   3000
      TabIndex        =   23
      Top             =   1200
      Width           =   1335
   End
   Begin VB.ComboBox order 
      Height          =   315
      ItemData        =   "dsails.frx":0000
      Left            =   9960
      List            =   "dsails.frx":0013
      TabIndex        =   18
      Top             =   120
      Width           =   1095
   End
   Begin VB.ComboBox Category 
      Height          =   315
      ItemData        =   "dsails.frx":004C
      Left            =   4320
      List            =   "dsails.frx":0071
      TabIndex        =   15
      Top             =   120
      Width           =   615
   End
   Begin VB.ComboBox sowner 
      Height          =   315
      ItemData        =   "dsails.frx":009B
      Left            =   2640
      List            =   "dsails.frx":00A8
      TabIndex        =   14
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Search 
      Caption         =   "&Search"
      Height          =   375
      Left            =   10200
      TabIndex        =   12
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Clear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   11640
      TabIndex        =   11
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton close 
      Caption         =   "&Close"
      Height          =   375
      Left            =   10200
      TabIndex        =   10
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton report 
      Caption         =   "&Print"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9000
      TabIndex        =   9
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox through 
      Height          =   315
      Left            =   4800
      TabIndex        =   8
      Top             =   600
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   6960
      TabIndex        =   6
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      Format          =   58720257
      CurrentDate     =   38162
   End
   Begin VB.TextBox pname 
      Height          =   315
      Left            =   1200
      TabIndex        =   2
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox sdate 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd-mmm-yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   315
      Left            =   5640
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox edate 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd-mmm-yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   315
      Left            =   7440
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   8640
      TabIndex        =   7
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      Format          =   58720257
      CurrentDate     =   38162
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3855
      Left            =   -120
      TabIndex        =   13
      Top             =   4320
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   6800
      _Version        =   393216
      Cols            =   29
      AllowUserResizing=   1
      MousePointer    =   1
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   8400
      TabIndex        =   51
      Top             =   1320
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      Format          =   58720257
      CurrentDate     =   38162
   End
   Begin VB.TextBox recdate1 
      Height          =   315
      Left            =   7320
      TabIndex        =   52
      Top             =   1320
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   8400
      TabIndex        =   56
      Top             =   1680
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      Format          =   58720257
      CurrentDate     =   38162
   End
   Begin MSComCtl2.DTPicker DTPicker5 
      Height          =   375
      Left            =   8400
      TabIndex        =   57
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      Format          =   58720257
      CurrentDate     =   38162
   End
   Begin VB.Frame Frame1 
      Caption         =   "Details."
      Height          =   3135
      Left            =   0
      TabIndex        =   20
      Top             =   960
      Width           =   8775
      Begin VB.TextBox comment 
         Height          =   495
         Left            =   5520
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   67
         Text            =   "dsails.frx":00B8
         Top             =   1440
         Width           =   3015
      End
      Begin VB.CommandButton save 
         Caption         =   "&Save"
         Height          =   375
         Left            =   3960
         TabIndex        =   60
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox payrec1 
         Height          =   325
         Left            =   5520
         TabIndex        =   44
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox dwt 
         Height          =   325
         Left            =   1200
         TabIndex        =   22
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label31 
         Caption         =   "Comm."
         Height          =   255
         Left            =   4560
         TabIndex        =   68
         Top             =   1560
         Width           =   855
      End
      Begin VB.Line Line5 
         X1              =   120
         X2              =   8640
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label Label29 
         Caption         =   "Dt."
         Height          =   255
         Left            =   6960
         TabIndex        =   64
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label28 
         Caption         =   "Dt."
         Height          =   255
         Left            =   6960
         TabIndex        =   63
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label25 
         Caption         =   "Balance"
         Height          =   375
         Left            =   4560
         TabIndex        =   59
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label24 
         Caption         =   "Dt."
         Height          =   255
         Left            =   6960
         TabIndex        =   53
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label23 
         Caption         =   "Pay.Rec3."
         Height          =   375
         Left            =   4560
         TabIndex        =   49
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label22 
         Caption         =   "Pay.Rec2."
         Height          =   255
         Left            =   4560
         TabIndex        =   48
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label20 
         Caption         =   "Pay.Rec1."
         Height          =   375
         Left            =   4560
         TabIndex        =   45
         Top             =   360
         Width           =   735
      End
      Begin VB.Line Line3 
         X1              =   4440
         X2              =   4440
         Y1              =   120
         Y2              =   2520
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   8640
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label Label19 
         Caption         =   "Net. AMt."
         Height          =   255
         Left            =   2040
         TabIndex        =   42
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "Making"
         Height          =   255
         Left            =   2280
         TabIndex        =   40
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label17 
         Caption         =   "Matel Rt."
         Height          =   255
         Left            =   2280
         TabIndex        =   39
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "Matel Wt."
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Col.Rt."
         Height          =   255
         Left            =   2280
         TabIndex        =   35
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Col St2.Wt"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Col.RT"
         Height          =   255
         Left            =   2280
         TabIndex        =   31
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Col St1.Wt"
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Dia:Rt"
         Height          =   375
         Left            =   2280
         TabIndex        =   26
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Dia:Wt."
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   8520
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label Label30 
      Caption         =   "S.Code"
      Height          =   255
      Left            =   240
      TabIndex        =   65
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label27 
      Caption         =   "Dt."
      Height          =   255
      Left            =   6960
      TabIndex        =   62
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label26 
      Caption         =   "Dt."
      Height          =   255
      Left            =   6960
      TabIndex        =   61
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label21 
      Caption         =   "Pay.Rec1."
      Height          =   375
      Left            =   5400
      TabIndex        =   47
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label11 
      Caption         =   "Col St1.wt"
      Height          =   375
      Left            =   2760
      TabIndex        =   30
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "Dia:Wt."
      Height          =   375
      Left            =   2520
      TabIndex        =   25
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Dia:Wt."
      Height          =   375
      Left            =   2520
      TabIndex        =   24
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Order By:"
      Height          =   255
      Left            =   9000
      TabIndex        =   19
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Category"
      Height          =   255
      Left            =   3360
      TabIndex        =   17
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Owner"
      Height          =   255
      Left            =   1920
      TabIndex        =   16
      Top             =   120
      Width           =   615
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   11880
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Image Image1 
      Height          =   2415
      Left            =   9120
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "Through:"
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Party Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Date"
      Height          =   255
      Left            =   5040
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "dsails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim res As New ADODB.Recordset
Dim res1 As New ADODB.Recordset
Dim str1 As String
Dim recdt1 As String
Dim recdt2 As String
Private Sub Clear_Click()
report.Enabled = False
clearing
End Sub
Private Sub clearing()
    MSFlexGrid1.Rows = 2
    MSFlexGrid1.TextMatrix(1, 0) = ""
    MSFlexGrid1.TextMatrix(1, 1) = ""
    MSFlexGrid1.TextMatrix(1, 2) = ""
    MSFlexGrid1.TextMatrix(1, 3) = ""
    MSFlexGrid1.TextMatrix(1, 4) = ""
    MSFlexGrid1.TextMatrix(1, 5) = ""
    MSFlexGrid1.TextMatrix(1, 6) = ""
    MSFlexGrid1.TextMatrix(1, 7) = ""
    MSFlexGrid1.TextMatrix(1, 8) = ""
    MSFlexGrid1.TextMatrix(1, 9) = ""
    MSFlexGrid1.TextMatrix(1, 10) = ""
    MSFlexGrid1.TextMatrix(1, 11) = ""
    MSFlexGrid1.TextMatrix(1, 12) = ""
    MSFlexGrid1.TextMatrix(1, 13) = ""
    MSFlexGrid1.TextMatrix(1, 14) = ""
    MSFlexGrid1.TextMatrix(1, 15) = ""
    MSFlexGrid1.TextMatrix(1, 16) = ""
    MSFlexGrid1.TextMatrix(1, 17) = ""
    MSFlexGrid1.TextMatrix(1, 18) = ""
    MSFlexGrid1.TextMatrix(1, 19) = ""
    MSFlexGrid1.TextMatrix(1, 20) = ""
    MSFlexGrid1.TextMatrix(1, 21) = ""
    MSFlexGrid1.TextMatrix(1, 22) = ""
    MSFlexGrid1.TextMatrix(1, 23) = ""
    MSFlexGrid1.TextMatrix(1, 24) = ""
    MSFlexGrid1.TextMatrix(1, 25) = ""
    MSFlexGrid1.TextMatrix(1, 26) = ""
    MSFlexGrid1.TextMatrix(1, 27) = ""
End Sub
Private Sub close_Click()
Unload Me
End Sub
Private Sub crt1_Change()
calculate
End Sub
Private Sub crt2_Change()
calculate
End Sub
Private Sub cst1_Change()
calculate
End Sub
Private Sub cst2_Change()
calculate
End Sub
Private Sub drt_Change()
calculate
End Sub
Private Sub DTPicker1_Change()
sdate = DTPicker1
End Sub
Private Sub DTPicker1_Click()
sdate = DTPicker1
End Sub
Private Sub DTPicker2_Click()
sdate = DTPicker2
End Sub
Private Sub DTPicker2_Change()
edate = DTPicker2
End Sub
Private Sub DTPicker3_Change()
recdate1 = DTPicker3
End Sub

Private Sub DTPicker3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
recdate1 = DTPicker3
End Sub
Private Sub DTPicker4_Change()
recdate2 = DTPicker4
End Sub

Private Sub DTPicker5_Change()
recdate3 = DTPicker5
End Sub

Private Sub dwt_Change()
calculate
End Sub

Private Sub export_Click()
On Error Resume Next

Dim objExcel As Excel.Application, objBook As Excel.Workbook, objsheet As Excel.Worksheet

Set objExcel = GetObject(, "Excel.Application") 'if excel already open you can use GetObject
If Err.Number Then
Err.Clear
End If
Set objExcel = CreateObject("Excel.Application") 'or CreateObject to open new Excel Application
''Set objBook = objExcel.Workbooks.Open("c:\daksh\vb\Book3.xls")
Set objBook = objExcel.Workbooks.add
Set objsheet = objBook.Worksheets(1)


Dim col
col = ""
Dim i
objsheet.Cells(2, 1) = "Date"
objsheet.Cells(2, 2) = "Party"
objsheet.Cells(2, 3) = "Through"
objsheet.Cells(2, 4) = "Owner"
objsheet.Cells(2, 5) = "Item"
objsheet.Cells(2, 6) = "D.WT"
objsheet.Cells(2, 7) = "D.RT"
objsheet.Cells(2, 8) = "Stone"
objsheet.Cells(2, 9) = "WT."
objsheet.Cells(2, 10) = "RT."
objsheet.Cells(2, 11) = "Stone2"
objsheet.Cells(2, 12) = "WT."
objsheet.Cells(2, 13) = "RT."
objsheet.Cells(2, 14) = "M.WT."
objsheet.Cells(2, 15) = "M.RT"
objsheet.Cells(2, 16) = "Make"
objsheet.Cells(2, 17) = "Amount"
i = 3
''MsgBox (MSFlexGrid1.Rows)
''Debug.Print MFlexGrid1.Rows
Dim total
For j = 1 To MSFlexGrid1.Rows - 2

objsheet.Cells(i, 1) = MSFlexGrid1.TextMatrix(j, 5)

objsheet.Cells(i, 2) = MSFlexGrid1.TextMatrix(j, 4)
objsheet.Cells(i, 3) = MSFlexGrid1.TextMatrix(j, 3)
objsheet.Cells(i, 4) = MSFlexGrid1.TextMatrix(j, 2)
objsheet.Cells(i, 5) = MSFlexGrid1.TextMatrix(j, 1)
objsheet.Cells(i, 6) = MSFlexGrid1.TextMatrix(j, 7)
objsheet.Cells(i, 7) = MSFlexGrid1.TextMatrix(j, 8)

objsheet.Cells(i, 8) = MSFlexGrid1.TextMatrix(j, 9)
objsheet.Cells(i, 9) = MSFlexGrid1.TextMatrix(j, 10)
objsheet.Cells(i, 10) = MSFlexGrid1.TextMatrix(j, 11)
objsheet.Cells(i, 11) = MSFlexGrid1.TextMatrix(j, 12)
objsheet.Cells(i, 12) = MSFlexGrid1.TextMatrix(j, 13)
objsheet.Cells(i, 13) = MSFlexGrid1.TextMatrix(j, 14)
objsheet.Cells(i, 14) = MSFlexGrid1.TextMatrix(j, 15)
objsheet.Cells(i, 15) = MSFlexGrid1.TextMatrix(j, 16)
objsheet.Cells(i, 16) = MSFlexGrid1.TextMatrix(j, 17)

objsheet.Cells(i, 17) = MSFlexGrid1.TextMatrix(j, 18)
objsheet.Cells(i, 18).Select

Dim path
If (MSFlexGrid1.TextMatrix(j, 19) <> "") Then
path = "D:\DAKSH\images\" & MSFlexGrid1.TextMatrix(j, 19)
ActiveSheet.Pictures.Insert(path).Select
Selection.ShapeRange.Height = 36#
Selection.ShapeRange.Width = 49#
Else
''Selection.delete
End If

i = i + 3
col = ""
Next j


''objBook.SaveAs "DataExpo.xls"
''objExcel.Workbooks.Open "DatExpo.xls"
objExcel.Visible = True
objBook.PrintPreview
''objBook.CLOSE
' cell(1,1) means cell A1 ;
Set objsheet = Null
Set objBook = Null



End Sub

Private Sub Form_Load()
DTPicker1.Format = dtpCustom
DTPicker2.Format = dtpCustom
MSFlexGrid1.TextMatrix(0, 0) = "S.No."
MSFlexGrid1.ColWidth(0) = 450
MSFlexGrid1.TextMatrix(0, 1) = "Item Code"
MSFlexGrid1.ColWidth(1) = 1000
MSFlexGrid1.TextMatrix(0, 2) = "Owner"
MSFlexGrid1.ColWidth(2) = 650
MSFlexGrid1.TextMatrix(0, 3) = "Through"
MSFlexGrid1.ColWidth(3) = 1300
MSFlexGrid1.TextMatrix(0, 4) = "P.Name"
MSFlexGrid1.ColWidth(4) = 1300
MSFlexGrid1.TextMatrix(0, 5) = "S.Date"
MSFlexGrid1.ColWidth(5) = 900
MSFlexGrid1.TextMatrix(0, 6) = "D.Date"
MSFlexGrid1.ColWidth(6) = 900
MSFlexGrid1.TextMatrix(0, 7) = ""
MSFlexGrid1.ColWidth(7) = 0
MSFlexGrid1.TextMatrix(0, 8) = ""
MSFlexGrid1.ColWidth(8) = 0
MSFlexGrid1.TextMatrix(0, 9) = "Stone1."
MSFlexGrid1.ColWidth(9) = 650
MSFlexGrid1.TextMatrix(0, 10) = ""
MSFlexGrid1.ColWidth(10) = 0
MSFlexGrid1.TextMatrix(0, 11) = ""
MSFlexGrid1.ColWidth(11) = 0
MSFlexGrid1.TextMatrix(0, 12) = ""
MSFlexGrid1.ColWidth(12) = 0
MSFlexGrid1.TextMatrix(0, 13) = ""
MSFlexGrid1.ColWidth(13) = 0
MSFlexGrid1.TextMatrix(0, 14) = ""
MSFlexGrid1.ColWidth(14) = 0
MSFlexGrid1.TextMatrix(0, 15) = ""
MSFlexGrid1.ColWidth(15) = 0
MSFlexGrid1.TextMatrix(0, 16) = ""
MSFlexGrid1.ColWidth(16) = 0
MSFlexGrid1.TextMatrix(0, 17) = ""
MSFlexGrid1.ColWidth(17) = 0
MSFlexGrid1.TextMatrix(0, 18) = "Total Amt."
MSFlexGrid1.ColWidth(18) = 1000
MSFlexGrid1.TextMatrix(0, 19) = ""
MSFlexGrid1.ColWidth(19) = 0
MSFlexGrid1.TextMatrix(0, 20) = ""
MSFlexGrid1.ColWidth(20) = 0
MSFlexGrid1.TextMatrix(0, 21) = ""
MSFlexGrid1.ColWidth(21) = 0
MSFlexGrid1.TextMatrix(0, 22) = ""
MSFlexGrid1.ColWidth(22) = 0
MSFlexGrid1.TextMatrix(0, 23) = ""
MSFlexGrid1.ColWidth(23) = 0
MSFlexGrid1.TextMatrix(0, 24) = ""
MSFlexGrid1.ColWidth(24) = 0
MSFlexGrid1.TextMatrix(0, 25) = ""
MSFlexGrid1.ColWidth(25) = 0
MSFlexGrid1.TextMatrix(0, 26) = ""
MSFlexGrid1.ColWidth(26) = 0
MSFlexGrid1.TextMatrix(0, 27) = ""
MSFlexGrid1.ColWidth(27) = 0
MSFlexGrid1.TextMatrix(0, 28) = "Status"
MSFlexGrid1.ColWidth(28) = 650
End Sub
Private Sub mak_Change()
calculate
End Sub
Private Sub mrt_Change()
calculate
End Sub
Private Sub cleardetail()
 scode = ""
 dwt = ""
 drt = ""
 cst1 = ""
 crt1 = ""
 cst2 = ""
 crt2 = ""
 mwt = ""
 mrt = ""
mak = ""
total = ""
bal = ""
payrec1 = ""
payrec2 = ""
payrec3 = ""
recdate1 = ""
recdate2 = ""
recdate3 = ""
comment = ""
End Sub
Private Sub MSFlexGrid1_Click()
On Error Resume Next
If (MSFlexGrid1.col = 28) Then
    If (MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 28) = "C") Then
        MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 28) = "P"
        Else
        MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 28) = "C"
    End If
End If
End Sub
Private Sub MSFlexGrid1_SelChange()
On Error Resume Next
   cleardetail
   save.Enabled = True
   scode = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 26)
   dwt = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7)
   drt = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 8)
   cst1 = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 10)
   crt1 = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 11)
   cst2 = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 13)
   crt2 = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 14)
   mwt = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 15)
   mrt = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 16)
   mak = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 17)
   comment = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 27)
   payrec1 = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 20)
   payrec2 = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 21)
   payrec3 = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 22)
   recdate1 = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 23)
   recdate2 = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 24)
   recdate3 = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 25)
   
   calculate
   
If (MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 19) <> "" And MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 19) <> Empty) Then
     Dim picturename As String
     pos = InStrRev(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 19), "\")
     If (pos <> 0) Then
     picturename = Mid(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 19), pos + 1)
     Else
     picturename = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 19)
     End If
     Image1.Visible = True
     Image1.Picture = LoadPicture(MDIForm1.picturepath & picturename)
     'Image1.Picture = LoadPicture(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 17))
Else
    Image1.Picture = LoadPicture()
End If
End Sub
Private Sub calculate()
total = Round(Val(dwt) * Val(drt) + Val(cst1) * Val(crt1) + Val(cst2) * Val(crt2) + Val(mwt) * Val(mrt) + Val(mak))
bal = total - (Val(payrec1) + Val(payrec2) + Val(payrec3))
End Sub

Private Sub mwt_Change()
calculate
End Sub

Private Sub Option1_Click(Index As Integer)
If (Index = 0) Then
save.Enabled = True
Else
save.Enabled = False
End If
End Sub

Private Sub payrec1_Change()
calculate
End Sub

Private Sub payrec2_Change()
calculate
End Sub

Private Sub payrec3_Change()
calculate
End Sub



Private Sub remove_Click()
On Error Resume Next
str1 = ""
str2 = ""
updatedcode = ""
Dim ritems(350) As Integer
Dim k As Integer
For j = 0 To MSFlexGrid1.Rows - 1
If (MSFlexGrid1.TextMatrix(j, 28) = "C") Then
    num = Val(MSFlexGrid1.TextMatrix(j, 26))
    ritems(k) = j
    k = k + 1

'If (j <> MSFlexGrid1.Rows - 1) Then
str1 = str1 & MSFlexGrid1.TextMatrix(j, 26) & ","
updatedcode = updatedcode & num & ","
'Else
'str1 = str1 & MSFlexGrid1.TextMatrix(j, 1)
'updatedcode = updatedcode & num
'End If
End If
Next j
'MsgBox (str1)
pos = InStrRev(str1, ",")
str1 = Mid(str1, 1, pos - 1)

If (str1 = "") Then
MsgBox ("Please Change The Status Of Any Item")
Exit Sub
End If

pos = InStrRev(updatedcode, ",")
updatedcode = Mid(updatedcode, 1, pos - 1)

result = MsgBox("Do You Want To Change The Status Of Thease Items " & str1, vbYesNo + vbDefaultButton2, "Confirmation")
 If (Val(result) = 7) Then
    Exit Sub
 Else
   ' str2 = "Select * from tblestsalevalue where nscode in (" & updatedcode & ")"
 ' MsgBox (str2)
    res1.Open " Select * from tblestsalevalue where nscode in (" & updatedcode & ")", MDIForm1.con1, adOpenStaic, adLockOptimistic
   
  If (res1.EOF = False) Then
        res1.MoveFirst
    Do While Not res1.EOF
    If (MDIForm1.con2.state = 0) Then
     Exit Sub
    End If
    res.Open "tblestsalevalue", MDIForm1.con2, adOpenDynamic, adLockOptimistic
    res.Addnew
    res!ncode = res1!ncode
    res!dinvoicedate = res1!dinvoicedate
    res!chowner = res1!chowner
    res!chcategory = res1!chcategory
    res!chpcode = res1!chpcode
    
    If (IsNull(res1!npno) = False) Then
    res!npno = res1!npno
    Else
    res!npno = 0
    End If
    
    If (res1!dsdate <> Empty) Then
    res!dsdate = res1!dsdate
    Else
    res!dsdate = CDate(date)
    End If
    
    If (res1!chthrough <> Empty) Then
    res!chthrough = res1!chthrough
    Else
    res!chthrough = ""
    End If
    
    If (res1!chpicture <> Empty) Then
    res!chpicture = res1!chpicture
    Else
    res!chpicture = ""
    End If
    
    If (res1!chpartyname <> Empty) Then
    res!chpartyname = res1!chpartyname
    Else
    res!chpartyname = ""
    End If
    
    res!chcontent2 = res1!chcontent2
    res!chcontent3 = res1!chcontent3
    
    res!nweight1 = res1!nweight1
    res!nweight2 = res1!nweight2
    res!nweight3 = res1!nweight3
    res!nweight4 = res1!nweight4
    
    res!nsrate1 = res1!nsrate1
    res!nsrate2 = res1!nsrate2
    res!nsrate3 = res1!nsrate3
    res!nsrate4 = res1!nsrate4
    
    res!nsmaking = res1!nsmaking
    res!nscode = res1!nscode
    res!dfpaydate = date
    res!chcomment = res1!chcomment
    res.update
    res.close
   'Now setting the current row data empty.
    
    res1!ncode = 0
    res1!npno = 0
    res1!chstatus = "NA"
   ' res1!chItemname = ""
    res1!chpcode = ""
    res1!chowner = ""
    res1!chthrough = ""
    res1!chpartyname = ""
    res1!chpicture = ""
    'res1!dinvoicedate = ""
    res1!chcontent2 = ""
    res1!chcontent3 = ""
    res1!nweight1 = 0
    res1!nweight2 = 0
    res1!nweight3 = 0
    res1!nweight4 = 0
    res1!nsrate1 = 0
    res1!nsrate2 = 0
    res1!nsrate3 = 0
    res1!nsrate4 = 0
    'res1!nigw = 0
    res1!nsmaking = 0
    res1!npayrec1 = 0
    res1!npayrec2 = 0
    res1!npayrec3 = 0
    res1!ncom = 0
    res1!chcomment = ""
    res1.update
    res1.MoveNext
  Loop
End If
    res1.close
   For i = 0 To k - 1
    MSFlexGrid1.RemoveItem (Val(ritems(i)) - i)
    Next
    MsgBox ("Status Changed")
End If
End Sub

Private Sub report_Click()
On Error Resume Next
If (str1 = Empty) Then
If (Option1(0) = True) Then
str1 = "select Tblestsalevalue.*,Tblestsalevalue.ndno,IIF(Tblestsalevalue.chowner ='LE',Tblestsalevalue.chowner,Tblestsalevalue.chpcode) AS code, Tblestsalevalue.nsrate1 * Tblestsalevalue.nweight1 as amt1,Tblestsalevalue.nsrate2 * Tblestsalevalue.nweight2 as amt2,Tblestsalevalue.nsrate3 * Tblestsalevalue.nweight3 as amt3,"
str1 = str1 & " Tblestsalevalue.nsrate4 * Tblestsalevalue.nweight4 as amt4,amt1+amt2+amt3+amt4+nsmaking AS totamt,"
str1 = str1 & "  totamt-(IIF(npayrec1>0,val(npayrec1),0)+IIF(npayrec2>0,val(npayrec2),0)+IIF(npayrec3>0,val(npayrec3),0)) AS bal1,IIF(bal<50,'Rec.',' ') as bal "
str1 = str1 & " from Tblestsalevalue order by dsdate"
Else
str1 = "select Tblestsalevalue.*,Tblestsalevalue.ndno,IIF(Tblestsalevalue.chowner ='LE',Tblestsalevalue.chowner,Tblestsalevalue.chpcode) AS code, Tblestsalevalue.nsrate1 * Tblestsalevalue.nweight1 as amt1,Tblestsalevalue.nsrate2 * Tblestsalevalue.nweight2 as amt2,Tblestsalevalue.nsrate3 * Tblestsalevalue.nweight3 as amt3,"
str1 = str1 & " Tblestsalevalue.nsrate4 * Tblestsalevalue.nweight4 as amt4,amt1+amt2+amt3+amt4+nsmaking AS totamt "
str1 = str1 & " from Tblestsalevalue order by dsdate"
End If
End If

'If (Option1(0) = True) Then
'str1 = "select Tblcosting.*, Tblcosting.minrate1 * tblcosting.nweight1 as amt1,Tblcosting.minrate2 * tblcosting.nweight2 as amt2,Tblcosting.minrate3 * tblcosting.nweight3 as amt3,"
'str1 = str1 & " Tblcosting.minrate4 * tblcosting.nweight4 as amt4,amt1+amt2+amt3+amt4+nmaking1 AS totamt,totamt * 1.333333 as priceno,priceno-((priceno)/100 * " & Val(dis.Text)
'str1 = str1 & ") as ftotamt,(ftotamt-(amt2+amt3+amt4+nmaking1))/nweight1 as dprice "
'str1 = str1 & " from tblcosting "
'Else
'str1 = "select Tblcosting.*, Tblcosting.nrate1 * tblcosting.nweight1 as amt1,Tblcosting.nrate2 * tblcosting.nweight2 as amt2,Tblcosting.nrate3 * tblcosting.nweight3 as amt3,"
'str1 = str1 & " Tblcosting.nrate4 * tblcosting.nweight4 as amt4,amt1+amt2+amt3+amt4+nmaking AS totamt,totamt * 1.333333 as priceno,priceno-((priceno)/100 * " & Val(dis.Text)
'str1 = str1 & ") as ftotamt,(ftotamt-(amt2+amt3+amt4+nmaking1))/nweight1 as dprice "
'str1 = str1 & " from tblcosting "
'End If
'Debug.Print (str1)
If (Option1(0) = True) Then
DataEnvironment1.Connection1 = MDIForm1.con1
Else
DataEnvironment1.Connection1 = MDIForm1.con2
End If

DataEnvironment1.Commands(2).CommandText = str1
If DataEnvironment1.rsCommand2.state = adStateOpen Then
DataEnvironment1.rsCommand2.close
End If

DataEnvironment1.Commands(2).CommandText = str1
DataReport8.PrintReport True
End Sub

Private Sub save_Click()
On Error Resume Next
   MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 26) = scode
   MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 7) = dwt
   MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 8) = drt
   MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 10) = cst1
   MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 11) = crt1
   MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 13) = cst2
   MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 14) = crt2
   MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 15) = mwt
   MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 16) = mrt
   MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 17) = mak
   MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 27) = comment.Text
   str1 = "select npayrec1,npayrec2,npayrec3,drecdate1,drecdate2,drecdate3,nweight1,nsrate1,nweight2,nsrate2,nweight3,nsrate3,nweight4,nsrate4,nsmaking,chcomment from tblestsalevalue where nscode=" & scode.Text
   res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
   res1!nweight1 = dwt
   res1!nsrate1 = drt
   res1!nweight2 = cst1
   res1!nsrate2 = crt1
   res1!nweight3 = cst2
   res1!nsrate3 = crt2
   res1!nweight4 = mwt
   res1!nsrate4 = mrt
   res1!nsmaking = mak
   res1!npayrec1 = Val(payrec1)
   res1!npayrec2 = Val(payrec2)
   res1!npayrec3 = Val(payrec3)
   res1!chcomment = comment.Text
   
   If (recdate1 <> Empty) Then
   res1!drecdate1 = CDate(recdate1)
   End If
   If (recdate2 <> Empty) Then
   res1!drecdate2 = CDate(recdate2)
   End If
   If (recdate3 <> Empty) Then
   res1!drecdate3 = CDate(recdate3)
   End If
   res1.update
   res1.close
   MSFlexGrid1.Refresh
   save.Enabled = False
   MsgBox ("Record Updated")
End Sub



Private Sub Search_Click()
On Error Resume Next
Dim recdt1 As String
Dim recdt2 As String
recdt1 = recdate1.Text
recdt2 = recdate2.Text
clearing
cleardetail
Dim str2 As String
str2 = ""

result = MsgBox("Do You Want To Display Complete List!", vbYesNoCancel, "Confirmation")
    If (Option1(0) = True) Then
     str1 = "select Tblestsalevalue.*,Tblestsalevalue.ndno,IIF(Tblestsalevalue.chowner ='LE',Tblestsalevalue.chowner,Tblestsalevalue.chpcode) AS code, Tblestsalevalue.nsrate1 * Tblestsalevalue.nweight1 as amt1,Tblestsalevalue.nsrate2 * Tblestsalevalue.nweight2 as amt2,Tblestsalevalue.nsrate3 * Tblestsalevalue.nweight3 as amt3,"
     str1 = str1 & " Tblestsalevalue.nsrate4 * Tblestsalevalue.nweight4 as amt4,amt1+amt2+amt3+amt4+nsmaking AS totamt,"
     str1 = str1 & "  totamt-(IIF(npayrec1>0,val(npayrec1),0)+IIF(npayrec2>0,val(npayrec2),0)+IIF(npayrec3>0,val(npayrec3),0)) AS bal1,IIF(bal1<100,'Rec.',' ') as bal "
     str1 = str1 & " from Tblestsalevalue "
Else
     str1 = "select Tblestsalevalue.*,Tblestsalevalue.ndno,IIF(Tblestsalevalue.chowner ='LE',Tblestsalevalue.chowner,Tblestsalevalue.chpcode) AS code, Tblestsalevalue.nsrate1 * Tblestsalevalue.nweight1 as amt1,Tblestsalevalue.nsrate2 * Tblestsalevalue.nweight2 as amt2,Tblestsalevalue.nsrate3 * Tblestsalevalue.nweight3 as amt3,"
     str1 = str1 & " Tblestsalevalue.nsrate4 * Tblestsalevalue.nweight4 as amt4,amt1+amt2+amt3+amt4+nsmaking AS totamt "
     str1 = str1 & " from Tblestsalevalue "
End If
      
    If (Val(result) = 7) Then
    If (str2 <> Empty) Then
     str2 = str2 + " and (Tblestsalevalue.nsrate1 * Tblestsalevalue.nweight1+Tblestsalevalue.nsrate2 * Tblestsalevalue.nweight2+Tblestsalevalue.nsrate3 * Tblestsalevalue.nweight3+Tblestsalevalue.nsrate4 * Tblestsalevalue.nweight4+Tblestsalevalue.nsmaking)"
     str2 = str2 + " - (IIF(npayrec1>0,val(npayrec1),0)+IIF(npayrec2>0,val(npayrec2),0)+IIF(npayrec3>0,val(npayrec3),0))>100"
    Else
     str2 = str2 + " (Tblestsalevalue.nsrate1 * Tblestsalevalue.nweight1+Tblestsalevalue.nsrate2 * Tblestsalevalue.nweight2+Tblestsalevalue.nsrate3 * Tblestsalevalue.nweight3+Tblestsalevalue.nsrate4 * Tblestsalevalue.nweight4+Tblestsalevalue.nsmaking)"
     str2 = str2 + " - (IIF(npayrec1>0,val(npayrec1),0)+IIF(npayrec2>0,val(npayrec2),0)+IIF(npayrec3>0,val(npayrec3),0))>100"
    End If
    
    If (Val(result) = 2) Then
        If (Option1(0) = True) Then
            If (str2 <> Empty) Then
             str2 = str2 + " and (Tblestsalevalue.nsrate1 * Tblestsalevalue.nweight1+Tblestsalevalue.nsrate2 * Tblestsalevalue.nweight2+Tblestsalevalue.nsrate3 * Tblestsalevalue.nweight3+Tblestsalevalue.nsrate4 * Tblestsalevalue.nweight4+Tblestsalevalue.nsmaking)"
             str2 = str2 + " - (IIF(npayrec1>0,val(npayrec1),0)+IIF(npayrec2>0,val(npayrec2),0)+IIF(npayrec3>0,val(npayrec3),0))<100"
            Else
             str2 = str2 + " (Tblestsalevalue.nsrate1 * Tblestsalevalue.nweight1+Tblestsalevalue.nsrate2 * Tblestsalevalue.nweight2+Tblestsalevalue.nsrate3 * Tblestsalevalue.nweight3+Tblestsalevalue.nsrate4 * Tblestsalevalue.nweight4+Tblestsalevalue.nsmaking)"
             str2 = str2 + " - (IIF(npayrec1>0,val(npayrec1),0)+IIF(npayrec2>0,val(npayrec2),0)+IIF(npayrec3>0,val(npayrec3),0))<100"
            End If
        End If
    End If
    
    If (scode.Text <> Empty) Then
    If (str2 <> Empty) Then
    str2 = str2 & " And chscode=" & "'" & scode.Text & "'"
    Else
    str2 = str2 & " chscode=" & "'" & scode.Text & "'"
    End If
    End If
    End If
  
    
   If (sdate.Text <> Empty) Then
    If (str2 <> Empty) Then
    str2 = str2 & " And dsdate>= #" & Format(sdate.Text, "dd-mmm-yy") & "# "
    Else
    str2 = str2 & " dsdate >= #" & Format(sdate.Text, "dd-mmm-yy") & "# "
    End If
  End If

    If (edate.Text <> Empty) Then
        If (str2 <> Empty) Then
            str2 = str2 & " And dsdate<=#" & Format(edate.Text, "dd-mmm-yy") & "# "
        Else
            str2 = str2 & " dsdate<=#" & Format(edate.Text, "dd-mmm-yy") & "# "
        End If
    End If

 If (recdt1 <> Empty) Then
    If (str2 <> Empty) Then
    str2 = str2 & " And dinvoicedate>= #" & Format(recdt1, "dd-mmm-yy") & "# "
    Else
    str2 = str2 & " dinvoicedate >= #" & Format(recdt1, "dd-mmm-yy") & "# "
    End If
  End If
'Debug.Print (str2)
    If (recdt2 <> Empty) Then
        If (str2 <> Empty) Then
            str2 = str2 & " And dinvoicedate<=#" & Format(recdt2, "dd-mmm-yy") & "# "
        Else
            str2 = str2 & " dinvoicedate<=#" & Format(recdt2, "dd-mmm-yy") & "# "
        End If
    End If

If (Category.Text <> Empty) Then
    If (str2 <> Empty) Then
    str2 = str2 & " And chcategory=" & "'" & Category.Text & "'"
    Else
    str2 = str2 & " chcategory=" & "'" & Category.Text & "'"
    End If
End If

If (sowner.Text <> Empty) Then
    If (str2 <> Empty) Then
    str2 = str2 & " And chowner=" & "'" & sowner.Text & "'"
    Else
    str2 = str2 & " chowner=" & "'" & sowner.Text & "'"
End If
End If

If (through.Text <> Empty) Then
    If (str2 <> Empty) Then
    str2 = str2 & " And chthrough=" & "'" & through & "'"
    Else
    str2 = str2 & "  chthrough=" & "'" & through & "'"
    End If
End If

If (pname.Text <> Empty) Then
    If (str2 <> Empty) Then
    str2 = str2 & " And chpartyname=" & "'" & pname.Text & "'"
    Else
    str2 = str2 & "  chpartyname=" & "'" & pname.Text & "'"
   End If
End If

If (order.Text = Empty) Then
    If (str2 <> Empty) Then
        If (Option1(0) = True) Then
        str1 = str1 & " where chstatus='A' And " & str2 & " order by nscode"
        Else
        str1 = str1 & " where " & str2 & " order by nscode"
        End If
    Else
        If (Option1(0) = True) Then
            str1 = str1 & " where chstatus='A' order by nscode"
        Else
            str1 = str1 & " order by nscode"
        End If
    End If
ElseIf (order.Text = "Category") Then
    If (str2 <> Empty) Then
            If (Option1(0) = True) Then
            str1 = str1 & " where chstatus='A' And " & str2 & " chcategory,ncode"
            Else
            str1 = str1 & " where " & str2 & " chcategory,ncode"
            End If
    Else
            If (Option1(0) = True) Then
            str1 = str1 & " where chstatus='A' order by chcategory,ncode"
            Else
            str1 = str1 & " order by chcategory,ncode"
            End If
    End If
Else
    If (str2 <> Empty) Then
        If (Option1(0) = True) Then
            str1 = str1 & " where chstatus='A' And " & str2 & " order by " & order.Text
        Else
            str1 = str1 & " where  " & str2 & " order by " & order.Text
        End If
        Else
           If (Option1(0) = True) Then
           str1 = str1 & " where chstatus='A' order by " & order.Text
           Else
           str1 = str1 & " order by " & order.Text
           End If
    End If
End If
report.Enabled = True
'Debug.Print (str1)
'MsgBox (str1)
If (Option1(0) = True) Then


res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
Else
res1.Open str1, MDIForm1.con2
End If

Dim i As Integer
i = 1
If (res1.EOF = True) Then
    MsgBox ("No Record Found")
Else
Dim gtotal As Double
Dim dwt As Double
Dim metal As Double
Dim metalamt As Double
Dim stonewt1 As Double
Dim stonewt2 As Double
Dim makingtot As Double
Dim diatotrt As Double
Dim stonetotamt1 As Double
Dim stonetotamt2 As Double

While (res1.EOF = False)
    MSFlexGrid1.TextMatrix(i, 0) = i
    MSFlexGrid1.TextMatrix(i, 1) = res1!chcategory & "-" & res1!ncode
    MSFlexGrid1.TextMatrix(i, 2) = res1!chowner
    MSFlexGrid1.TextMatrix(i, 3) = res1!chthrough
    MSFlexGrid1.TextMatrix(i, 4) = res1!chpartyname
    
    MSFlexGrid1.TextMatrix(i, 5) = Format(res1!dsdate, "dd-mmm-yy")
    If (Option1(0) = True) Then
    MSFlexGrid1.TextMatrix(i, 6) = Format(res1!dduedate, "dd-mmm-yy")
    Else
    MSFlexGrid1.TextMatrix(i, 6) = Format(res1!dsdate, "dd-mmm-yy")
    End If
    dweight = res1!nweight1
    MSFlexGrid1.TextMatrix(i, 7) = dweight
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 7
    MSFlexGrid1.CellBackColor = RGB(231, 228, 218)
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 8
    MSFlexGrid1.CellBackColor = RGB(231, 228, 218)
  '  If (Option1(0) = True) Then
  
    MSFlexGrid1.TextMatrix(i, 8) = Round(res1!nsrate1)
  '  Else
  '  MSFlexGrid1.TextMatrix(i, 5) = res1!nrate1
  '  End If
    
    
    MSFlexGrid1.TextMatrix(i, 9) = res1!chcontent2
    swt1 = res1!nweight2
    
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 10
    MSFlexGrid1.CellBackColor = RGB(231, 228, 218)
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 11
    MSFlexGrid1.CellBackColor = RGB(231, 228, 218)
    
    MSFlexGrid1.TextMatrix(i, 10) = swt1
    
    MSFlexGrid1.TextMatrix(i, 11) = res1!nsrate2
    
    
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 13
    MSFlexGrid1.CellBackColor = RGB(231, 228, 218)
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 14
    MSFlexGrid1.CellBackColor = RGB(231, 228, 218)
    
    MSFlexGrid1.TextMatrix(i, 12) = res1!chcontent3
    swt2 = res1!nweight3
    MSFlexGrid1.TextMatrix(i, 13) = swt2
    
    MSFlexGrid1.TextMatrix(i, 14) = res1!nsrate3
    
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 15
    MSFlexGrid1.CellBackColor = RGB(231, 228, 218)
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 16
    MSFlexGrid1.CellBackColor = RGB(231, 228, 218)
    
    mwt = res1!nweight4
    MSFlexGrid1.TextMatrix(i, 15) = Round(mwt, 2)
    
    MSFlexGrid1.TextMatrix(i, 16) = res1!nsrate4
    

    
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 17
    MSFlexGrid1.CellBackColor = RGB(234, 248, 222)
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 18
    MSFlexGrid1.CellBackColor = RGB(234, 248, 222)
    making = res1!nsmaking
    
    
    MSFlexGrid1.TextMatrix(i, 17) = making
    
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 18
    MSFlexGrid1.CellBackColor = RGB(255, 255, 236)
'    If (Option1(0) = True) Then
 '   amt1 = Val(res1!minrate1) * Val(res1!nweight1)
 '   amt2 = Val(res1!minrate2) * Val(res1!nweight2)
  '  amt3 = Val(res1!minrate3) * Val(res1!nweight3)
  '  amt4 = Val(res1!minrate4) * Val(res1!nweight4)
  '  Else
  '  amt1 = Val(res1!nrate1) * Val(res1!nweight1)
  '  amt2 = Val(res1!nrate2) * Val(res1!nweight2)
  '  amt3 = Val(res1!nrate3) * Val(res1!nweight3)
  '  amt4 = Val(res1!nrate4) * Val(res1!nweight4)
  '  End If
    metalamt = metalamt + res1!amt4
    stonetotamt1 = stonetotamt1 + res1!amt2
    stonetotamt2 = stonetotamt2 + res1!amt3
    'diatotrt = diatotrt + res!amt1
    'totamt = Val(amt1) + Val(amt2) + Val(amt3) + Val(amt4) + Val(making)
    'MSFlexGrid1.TextMatrix(i, 15) = Round(Val(totamt) * 00.01333333)
   ' MSFlexGrid1.TextMatrix(I, 15) = Round(Val(res1!priceno) / 100)
    'If (Val(dis) <> 20) Then
    'totamt = Val(totamt) * 1.333333
    'totamt = Val(totamt) - (Val(totamt) * Val(dis) / 100)
    'rate1 = (Val(totamt) - (Val(amt2) + Val(amt3) + Val(amt4) + Val(making))) / Val(dweight)
    'MSFlexGrid1.TextMatrix(i, 5) = Round(rate1)
   ' amt1 = Val(rate1) * Val(dweight)
   ' diatotrt = diatotrt + amt1
   ' End If
    diatotrt = diatotrt + res1!nsrate1 * res1!nweight1
    MSFlexGrid1.TextMatrix(i, 18) = Round(res1!totamt)
    MSFlexGrid1.TextMatrix(i, 28) = "P"
    
    If (IsNull(res1!chpicture)) Then
        MSFlexGrid1.TextMatrix(i, 19) = ""
    ElseIf (res1!chpicture <> "" And res1!chpicture <> Empty) Then
        MSFlexGrid1.TextMatrix(i, 19) = res1!chpicture
    End If
    If (Option1(0) = True) Then
    If (IsNull(res1!npayrec1) = False) Then
    MSFlexGrid1.TextMatrix(i, 20) = Round(res1!npayrec1)
    End If
    If (IsNull(res1!npayrec2) = False) Then
    MSFlexGrid1.TextMatrix(i, 21) = Round(res1!npayrec2)
    End If
    If (IsNull(res1!npayrec3) = False) Then
    MSFlexGrid1.TextMatrix(i, 22) = Round(res1!npayrec3)
    End If
    
    If (IsNull(res1!drecdate1) = False) Then
       MSFlexGrid1.TextMatrix(i, 23) = res1!drecdate1
    End If
    
    If (IsNull(res1!drecdate2) = False) Then
    MSFlexGrid1.TextMatrix(i, 24) = res1!drecdate2
    End If
    
    If (IsNull(res1!drecdate3) = False) Then
    MSFlexGrid1.TextMatrix(i, 25) = res1!drecdate3
    End If
    End If
    If (IsNull(res1!nscode) = False) Then
    MSFlexGrid1.TextMatrix(i, 26) = res1!nscode
    End If
    
    If (IsNull(res1!chcomment) = False) Then
        If (res1!chcomment <> "") Then
            MSFlexGrid1.row = i
            MSFlexGrid1.col = 0
            MSFlexGrid1.CellFontBold = True
            MSFlexGrid1.CellForeColor = vbRed
        End If
    MSFlexGrid1.TextMatrix(i, 27) = res1!chcomment
    End If
    
    
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 16
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.CellForeColor = vbBlue
    
    gtotal = gtotal + res1!totamt
    dwt = dwt + dweight
    metal = metal + mwt
    stonewt1 = stonewt1 + swt1
    stonewt2 = stonewt2 + swt2
    
    makingtot = makingtot + making
     i = i + 1
    MSFlexGrid1.Rows = i + 1
    res1.MoveNext
   
    Wend

    MSFlexGrid1.TextMatrix(i, 1) = "Grand Total:"
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 1
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
       
    MSFlexGrid1.TextMatrix(i, 7) = Round(dwt, 2)
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 7
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
        
    MSFlexGrid1.TextMatrix(i, 8) = Round(diatotrt)
    
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 8
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
        
    MSFlexGrid1.TextMatrix(i, 10) = Round(stonewt1, 2)
    
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 10
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
    
    MSFlexGrid1.TextMatrix(i, 11) = Round(stonetotamt1)
    
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 11
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
    
    MSFlexGrid1.TextMatrix(i, 13) = Round(stonewt2, 2)
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 13
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
    
    MSFlexGrid1.TextMatrix(i, 14) = Round(stonetotamt2)
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 14
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
    
    MSFlexGrid1.TextMatrix(i, 15) = Round(metal, 2)
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 15
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
    
    MSFlexGrid1.TextMatrix(i, 16) = Round(metalamt)
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 16
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
    
    MSFlexGrid1.TextMatrix(i, 17) = Round(makingtot)
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 17
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
    
    MSFlexGrid1.TextMatrix(i, 18) = Round(gtotal)
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 18
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
    
    End If
   ' MSFlexGrid1.RowIsVisible(I - 1) = False
    res1.close
End Sub
