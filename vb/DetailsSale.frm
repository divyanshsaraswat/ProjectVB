VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form2 
   Caption         =   "salesheet"
   ClientHeight    =   5715
   ClientLeft      =   1980
   ClientTop       =   1200
   ClientWidth     =   10635
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5715
   ScaleWidth      =   10635
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "DetailsSale.frx":0000
      Left            =   7800
      List            =   "DetailsSale.frx":0016
      TabIndex        =   40
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Stock"
      Height          =   375
      Left            =   7920
      TabIndex        =   39
      Top             =   1800
      Width           =   735
   End
   Begin VB.ComboBox order 
      Height          =   315
      ItemData        =   "DetailsSale.frx":0077
      Left            =   4440
      List            =   "DetailsSale.frx":008A
      TabIndex        =   36
      Top             =   1920
      Width           =   1095
   End
   Begin VB.ComboBox dis 
      Height          =   315
      ItemData        =   "DetailsSale.frx":00BC
      Left            =   6960
      List            =   "DetailsSale.frx":00D8
      TabIndex        =   34
      Text            =   "25%"
      Top             =   1320
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   31
      Top             =   1920
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   30
      Top             =   1920
      Value           =   -1  'True
      Width           =   375
   End
   Begin VB.ComboBox pstatus 
      Height          =   315
      ItemData        =   "DetailsSale.frx":0103
      Left            =   6840
      List            =   "DetailsSale.frx":010D
      TabIndex        =   28
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox edate 
      Height          =   315
      Left            =   6960
      TabIndex        =   27
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox sdate 
      Height          =   315
      Left            =   4440
      TabIndex        =   26
      Top             =   120
      Width           =   1095
   End
   Begin VB.ComboBox sowner 
      Height          =   315
      ItemData        =   "DetailsSale.frx":0129
      Left            =   1320
      List            =   "DetailsSale.frx":0139
      TabIndex        =   25
      Top             =   120
      Width           =   615
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   8160
      TabIndex        =   24
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yy"
      Format          =   121634817
      CurrentDate     =   37684
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   5640
      TabIndex        =   23
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yy"
      Format          =   121634817
      CurrentDate     =   37684
   End
   Begin VB.CommandButton report 
      Caption         =   "Report"
      Height          =   375
      Left            =   11160
      TabIndex        =   22
      Top             =   1320
      Width           =   1335
   End
   Begin VB.ComboBox content2 
      Height          =   315
      ItemData        =   "DetailsSale.frx":014D
      Left            =   2040
      List            =   "DetailsSale.frx":0193
      TabIndex        =   21
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton close 
      Caption         =   "&Close"
      Height          =   375
      Left            =   11160
      TabIndex        =   20
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Clear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   11160
      TabIndex        =   19
      Top             =   1680
      Width           =   1335
   End
   Begin VB.ComboBox content4 
      Height          =   315
      ItemData        =   "DetailsSale.frx":0217
      Left            =   4440
      List            =   "DetailsSale.frx":022A
      TabIndex        =   18
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox issueto 
      Height          =   315
      Left            =   4440
      TabIndex        =   14
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox maker 
      Height          =   315
      Left            =   1320
      TabIndex        =   13
      Top             =   1320
      Width           =   1095
   End
   Begin VB.ComboBox gpur1 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "DetailsSale.frx":0259
      Left            =   6960
      List            =   "DetailsSale.frx":027C
      TabIndex        =   9
      Top             =   720
      Width           =   735
   End
   Begin VB.ComboBox content1 
      Height          =   315
      ItemData        =   "DetailsSale.frx":02A7
      Left            =   1320
      List            =   "DetailsSale.frx":02ED
      TabIndex        =   7
      Top             =   720
      Width           =   615
   End
   Begin VB.ComboBox Category 
      Height          =   315
      ItemData        =   "DetailsSale.frx":0371
      Left            =   2040
      List            =   "DetailsSale.frx":0399
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5295
      Left            =   0
      TabIndex        =   2
      Top             =   2640
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9340
      _Version        =   393216
      Cols            =   18
      AllowUserResizing=   1
      MousePointer    =   1
   End
   Begin VB.CommandButton Search 
      Caption         =   "&Search"
      Height          =   375
      Left            =   11160
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox code 
      Height          =   315
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.ComboBox state 
      Height          =   315
      ItemData        =   "DetailsSale.frx":03C6
      Left            =   11160
      List            =   "DetailsSale.frx":03D0
      TabIndex        =   15
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label stlab 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   9120
      TabIndex        =   38
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Order By:"
      Height          =   255
      Left            =   3240
      TabIndex        =   37
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "Discount:"
      Height          =   255
      Left            =   5880
      TabIndex        =   35
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   " Cost Details:"
      Height          =   375
      Left            =   1560
      TabIndex        =   33
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "Sales Details:"
      Height          =   375
      Left            =   120
      TabIndex        =   32
      Top             =   1920
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   1695
      Left            =   9240
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Print St."
      Height          =   255
      Left            =   5880
      TabIndex        =   29
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "Metal:"
      Height          =   375
      Left            =   3600
      TabIndex        =   17
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lState 
      Alignment       =   2  'Center
      Caption         =   "St."
      Height          =   255
      Left            =   8520
      TabIndex        =   16
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label13 
      Caption         =   "Issue To:"
      Height          =   255
      Left            =   3360
      TabIndex        =   12
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Maker:"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Metal Qly.:"
      Height          =   375
      Left            =   5880
      TabIndex        =   10
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Stone:"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "End Date:"
      Height          =   375
      Left            =   6120
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "St. Date:"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Item Code:"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   11880
      Y1              =   2520
      Y2              =   2520
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con2 As New ADODB.Connection
Dim res As New ADODB.Recordset
Dim res1 As New ADODB.Recordset
Dim str1 As String
Private Sub Clear_Click()
clearing
End Sub
Private Sub close_Click()
Unload Me
'con2.close
End Sub
Private Sub DTPicker1_Change()
sdate = DTPicker1
End Sub
Private Sub DTPicker2_Change()
edate = DTPicker2
End Sub
Private Sub Form_Load()
'con2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\lahar\vb\db2.mdb;Persist Security Info=False"
'con2.Open
DTPicker1.Format = dtpCustom
DTPicker2.Format = dtpCustom
MSFlexGrid1.TextMatrix(0, 0) = "S.No."
MSFlexGrid1.ColWidth(0) = 450
MSFlexGrid1.TextMatrix(0, 1) = "Item Code"
MSFlexGrid1.ColWidth(1) = 850
MSFlexGrid1.TextMatrix(0, 2) = "Rec.Date"
MSFlexGrid1.ColWidth(2) = 900
MSFlexGrid1.TextMatrix(0, 3) = "Item Name"
MSFlexGrid1.ColWidth(3) = 1200
MSFlexGrid1.TextMatrix(0, 4) = "Dia.Wt."
MSFlexGrid1.ColWidth(4) = 600
MSFlexGrid1.TextMatrix(0, 5) = "Dia.Rt."
MSFlexGrid1.ColWidth(5) = 700
MSFlexGrid1.TextMatrix(0, 6) = "Stone1."
MSFlexGrid1.ColWidth(6) = 650
MSFlexGrid1.TextMatrix(0, 7) = "St1 Wt."
MSFlexGrid1.ColWidth(7) = 600
MSFlexGrid1.TextMatrix(0, 8) = "St1 Rt."
MSFlexGrid1.ColWidth(8) = 600
MSFlexGrid1.TextMatrix(0, 9) = "Stone2"
MSFlexGrid1.ColWidth(9) = 650
MSFlexGrid1.TextMatrix(0, 10) = "St2 Wt."
MSFlexGrid1.ColWidth(10) = 600
MSFlexGrid1.TextMatrix(0, 11) = "St2 Rt."
MSFlexGrid1.ColWidth(11) = 600
MSFlexGrid1.TextMatrix(0, 12) = "MetalWt."
MSFlexGrid1.ColWidth(12) = 700
MSFlexGrid1.TextMatrix(0, 13) = "MetatRt."
MSFlexGrid1.ColWidth(13) = 600
MSFlexGrid1.TextMatrix(0, 14) = "Making"
MSFlexGrid1.ColWidth(14) = 700
MSFlexGrid1.TextMatrix(0, 15) = "Price No."
MSFlexGrid1.ColWidth(15) = 700
MSFlexGrid1.TextMatrix(0, 16) = "Total Amt."
MSFlexGrid1.ColWidth(16) = 950
MSFlexGrid1.TextMatrix(0, 17) = " "
MSFlexGrid1.ColWidth(17) = 100
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
End Sub


Private Sub MSFlexGrid1_DblClick()
On Error Resume Next
Dim i As Integer
Dim j As Integer
Dim exc As Excel.Application
Set exc = CreateObject("Excel.Application")
exc.Workbooks.add
exc.Visible = True


With MSFlexGrid1
For i = 0 To .Rows - 1
For j = 0 To .Cols - 2
exc.Cells(i + 1, j + 1) = .TextMatrix(i, j)
exc.Cells(i + 1, j + 1).Borders.LineStyle = xlDouble
exc.Cells(i + 1, j + 1).Borders.color = vbBlue
Next j
Next i
exc.range("A1:" & Chr(65 + j) & 1).Font.Bold = True
exc.Columns("$A:" & "$" & Chr(65 + j)).AutoFit
End With
End Sub

Private Sub MSFlexGrid1_SelChange()
On Error Resume Next
If (MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 17) <> "" And MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 17) <> Empty) Then
     Dim picturename As String
     pos = InStrRev(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 17), "\")
     If (pos <> 0) Then
     picturename = Mid(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 17), pos + 1)
     Else
     picturename = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 17)
     End If
     Image1.Visible = True
     Image1.Picture = LoadPicture(MDIForm1.picturepath & picturename)
     'Image1.Picture = LoadPicture(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 17))
Else
    Image1.Picture = LoadPicture()
End If
        pos = InStr(1, MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1), "-", vbTextCompare)
        num = Val(Mid(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1), pos + 1))
        str1 = "select chissue,nappno from tblappmaster where nappno = (select nappno from tblappdetail where ncode=" & num & ")"
        'MsgBox (str1)
        res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
        If (res1.EOF = False) Then
              stlab.Caption = "Appno:" & res1!nappno & " And Issue to " & res1!chissue
        Else
            stlab.Caption = "Stock In Hand"
        End If
        res1.CLOSE
End Sub

Private Sub report_Click()
If (str1 = Empty) Then
str1 = "SELECT tblcosting.*, (Tblcosting.minrate1 * tblcosting.nweight1 + Tblcosting.minrate2 * tblcosting.nweight2 + Tblcosting.minrate3 * tblcosting.nweight3+Tblcosting.minrate4 * tblcosting.nweight4)+nmaking1 AS total From tblcosting where chstate='Avi.' ORDER BY chcategory,ncode"
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

If DataEnvironment2.rsCommand1.state = adStateOpen Then
DataEnvironment2.rsCommand1.CLOSE
End If
DataEnvironment2.Commands(1).CommandText = str1
DataReport1.PrintReport True, rptRangeAllPages
End Sub

Private Sub Search_Click()
clearing
Dim str2 As String
str2 = ""
Dim str3 As String
str3 = ""
If (Option1(0) = True) Then
str1 = "select Tblcosting.*,Tblcosting.minrate1 * tblcosting.nweight1 as amt1,Tblcosting.minrate2 * tblcosting.nweight2 as amt2,Tblcosting.minrate3 * tblcosting.nweight3 as amt3,"
str1 = str1 & " Tblcosting.minrate4 * tblcosting.nweight4 as amt4,tblcosting.nmaking1 as nmaking1,amt1+amt2+amt3+amt4+nmaking1 AS totamt,totamt * 1.333333 as priceno,priceno-((priceno)/100 * " & Val(dis.Text)
str1 = str1 & ") as ftotamt,(ftotamt-(amt2+amt3+amt4+nmaking1))/nweight1 as dprice "
str1 = str1 & " from tblcosting "
str3 = " val((Tblcosting.minrate1 * tblcosting.nweight1+Tblcosting.minrate2 * tblcosting.nweight2+Tblcosting.minrate3 * tblcosting.nweight3+Tblcosting.minrate4 * tblcosting.nweight4)+nmaking1) "
Else
str1 = "select Tblcosting.*, Tblcosting.nrate1 * tblcosting.nweight1 as amt1,Tblcosting.nrate2 * tblcosting.nweight2 as amt2,Tblcosting.nrate3 * tblcosting.nweight3 as amt3,"
str1 = str1 & " Tblcosting.nrate4 * tblcosting.nweight4 as amt4,tblcosting.nmaking as nmaking,amt1+amt2+amt3+amt4+nmaking AS totamt,totamt * 1.333333 as priceno,priceno-((priceno)/100 * " & Val(dis.Text)
str1 = str1 & ") as ftotamt,(ftotamt-(amt2+amt3+amt4+nmaking))/nweight1 as dprice "
str1 = str1 & " from tblcosting "
str3 = " val((Tblcosting.nrate1 * tblcosting.nweight1+Tblcosting.nrate2 * tblcosting.nweight2+Tblcosting.nrate3 * tblcosting.nweight3+Tblcosting.nrate4 * tblcosting.nweight4)+nmaking) "
End If
If (code.Text <> Empty) Then
    str2 = " ncode=" & Val(code.Text)
End If

    If (sdate.Text <> Empty) Then
    If (str2 <> Empty) Then
    str2 = str2 & " And dinvoicedate>= #" & Format(sdate.Text, "dd-mmm-yy") & "# "
    Else
    str2 = str2 & " dinvoicedate >= #" & Format(sdate.Text, "dd-mmm-yy") & "# "
    End If
    End If

    If (edate.Text <> Empty) Then
    If (str2 <> Empty) Then
    str2 = str2 & " And dinvoicedate<=#" & Format(edate.Text, "dd-mmm-yy") & "# "
    Else
    str2 = str2 & " dinvoicedate<=#" & Format(edate.Text, "dd-mmm-yy") & "# "
    End If
    End If

    If (state.Text <> Empty) Then
    If (str2 <> Empty) Then
    str2 = str2 & " And chstate=" & "'" & state.Text & "'"
    Else
    str2 = str2 & " chstate=" & "'" & state.Text & "'"
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

If (content1.Text = "Dia.") Then
    If (str2 <> Empty) Then
    str2 = str2 & " And  chcontent1=" & "'" & content1.Text & "'"
    Else
    str2 = str2 & " chcontent1=" & "'" & content1.Text & "'"
    End If
End If
If (content1.Text <> "Dia." And content1.Text <> Empty) Then
    If (str2 <> Empty) Then
    str2 = str2 & " And  chcontent2=" & "'" & content1.Text & "'"
    Else
    str2 = str2 & " chcontent2=" & "'" & content1.Text & "'"
    End If
End If
If (content2.Text <> "Dia." And content2.Text <> Empty) Then
    If (str2 <> Empty) Then
    str2 = str2 & " And  chcontent3=" & "'" & content2.Text & "'"
    Else
    str2 = str2 & " chcontent3=" & "'" & content2.Text & "'"
    End If
End If

If (content4.Text <> Empty) Then
    If (str2 <> Empty) Then
    str2 = str2 & " And chcontent4=" & "'" & content4.Text & "'"
    Else
    str2 = str2 & " chcontent4=" & "'" & content4.Text & "'"
    End If
End If

If (gpur1.Text <> Empty) Then
    If (str2 <> Empty) Then
        str2 = str2 & " And gpur=" & Val(gpur1.Text)
    Else
        str2 = str2 & " gpur=" & Val(gpur1.Text)
    End If
End If

If (maker.Text <> Empty) Then
    If (str2 <> Empty) Then
    str2 = str2 & " And chmaker=" & "'" & maker & "'"
    Else
    str2 = str2 & "  chmaker=" & "'" & maker & "'"
    End If
End If

If (Combo1.Text = "<=10000") Then
    If (str2 <> Empty) Then
    str2 = str2 & " And minrate1<=10000"
    Else
    str2 = str2 & "  minrate1<=10000"
    End If
ElseIf (Combo1.Text = ">10000 And <20000") Then
    If (str2 <> Empty) Then
    str2 = str2 & " And minrate1>10000 and minrate1<=20000"
    Else
    str2 = str2 & "  minrate1>10000 and minrate1<=20000"
    End If
ElseIf (Combo1.Text = ">20000 And <40000") Then
    If (str2 <> Empty) Then
    str2 = str2 & " And minrate1>20000 and minrate1<=40000"
    Else
    str2 = str2 & "  minrate1>20000 and minrate1<=40000"
    End If
ElseIf (Combo1.Text = ">40000 And <60000") Then
    If (str2 <> Empty) Then
    str2 = str2 & " And minrate1>40000 and minrate1<=60000"
    Else
    str2 = str2 & "  minrate1>40000 and minrate1<=60000"
    End If
ElseIf (Combo1.Text = ">60000 And <1 Lac") Then
    If (str2 <> Empty) Then
    str2 = str2 & " And minrate1>60000 and minrate1<=100000"
    Else
    str2 = str2 & "  minrate1>60000 and minrate1<=100000"
    End If
ElseIf (Combo1.Text = ">1 Lac") Then
    If (str2 <> Empty) Then
    str2 = str2 & " And minrate1>100000"
    Else
    str2 = str2 & "  minrate1>100000"
    End If

End If


If (issueto.Text <> Empty) Then
    If (str2 <> Empty) Then
    str2 = str2 & " And chissueto=" & "'" & issueto & "'"
    Else
    str2 = str2 & "  chissueto=" & "'" & issueto & "'"
    End If
End If

If (pstatus.Text <> Empty) Then
    If (str2 <> Empty) Then
    str2 = str2 & " And pflag=" & "'" & pstatus & "'"
   Else
   str2 = str2 & "  pflag=" & "'" & pstatus & "'"
   End If
End If

If (Check1.Value = 1) Then
    If (str2 <> Empty) Then
    str2 = str2 & " And tblcosting.ncode not in (select ncode from tblappdetail) "
   Else
    str2 = str2 & "  tblcosting.ncode not in (select ncode from tblappdetail) "
   End If
End If

If (order.Text = Empty Or order.Text = "Item Code") Then
    If (str2 <> Empty) Then
    str1 = str1 & " where " & str2 & " And chstate='Avi.' order by chcategory,ncode"
    Else
    str1 = str1 & " where chstate='Avi.' order by chcategory,ncode"
    End If
 ElseIf (order.Text = "Maker") Then
    If (str2 <> Empty) Then
    str1 = str1 & " where " & str2 & "  And chstate='Avi.' order by val(chmaker)"
    Else
    str1 = str1 & " where chstate='Avi.' order by val(chmaker)"
    End If
  ElseIf (order.Text = "Gold Rate") Then
    If (str2 <> Empty) Then
    str1 = str1 & " where " & str2 & "  And chstate='Avi.' order by val(minrate4)"
    Else
    str1 = str1 & " where chstate='Avi.' order by val(minrate4)"
    End If
  ElseIf (order.Text = "Price No.") Then
    If (str2 <> Empty) Then
    str1 = str1 & " where " & str2 & "  And chstate='Avi.' order by " & str3
    Else
    str1 = str1 & " where chstate='Avi.' order by " & str3
    End If
 Else
   If (str2 <> Empty) Then
    str1 = str1 & " where " & str2 & "  And chstate='Avi.' order by dinvoicedate"
    Else
    str1 = str1 & " where chstate='Avi.' order by dinvoicedate"
    End If
  End If


'Debug.Print (str1)
On Error GoTo errorhandler
res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
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
    MSFlexGrid1.TextMatrix(i, 2) = Format(res1!dinvoicedate, "dd-mmm-yy")
    MSFlexGrid1.TextMatrix(i, 3) = res1!chItemname
    dweight = res1!nweight1
    MSFlexGrid1.TextMatrix(i, 4) = dweight
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 4
    MSFlexGrid1.CellBackColor = RGB(231, 228, 218)
  '  If (Option1(0) = True) Then
    MSFlexGrid1.TextMatrix(i, 5) = Round(res1!dprice)
  '  Else
  '  MSFlexGrid1.TextMatrix(i, 5) = res1!nrate1
  '  End If
    
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 5
    MSFlexGrid1.CellBackColor = RGB(231, 228, 218)
    
    MSFlexGrid1.TextMatrix(i, 6) = res1!chcontent2
    swt1 = res1!nweight2
    MSFlexGrid1.TextMatrix(i, 7) = swt1
    If (Option1(0) = True) Then
    MSFlexGrid1.TextMatrix(i, 8) = res1!minrate2
    Else
    MSFlexGrid1.TextMatrix(i, 8) = res1!nrate2
    End If
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 7
    MSFlexGrid1.CellBackColor = RGB(231, 228, 218)
    
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 8
    MSFlexGrid1.CellBackColor = RGB(231, 228, 218)
    
    MSFlexGrid1.TextMatrix(i, 9) = res1!chcontent3
    swt2 = res1!nweight3
    MSFlexGrid1.TextMatrix(i, 10) = swt2
    If (Option1(0) = True) Then
    MSFlexGrid1.TextMatrix(i, 11) = res1!minrate3
    Else
    MSFlexGrid1.TextMatrix(i, 11) = res1!nrate3
    End If
    
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 10
    MSFlexGrid1.CellBackColor = RGB(231, 228, 218)
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 11
    MSFlexGrid1.CellBackColor = RGB(231, 228, 218)
    
    mwt = res1!nweight4
    MSFlexGrid1.TextMatrix(i, 12) = Round(mwt, 2)
    If (Option1(0) = True) Then
    MSFlexGrid1.TextMatrix(i, 13) = res1!minrate4
    Else
    MSFlexGrid1.TextMatrix(i, 13) = res1!nrate4
    End If
    
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 12
    MSFlexGrid1.CellBackColor = RGB(234, 248, 222)
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 13
    MSFlexGrid1.CellBackColor = RGB(234, 248, 222)
    If (Option1(0) = True) Then
    making = res1!nmaking1
    Else
    making = res1!nmaking
    End If
    
    MSFlexGrid1.TextMatrix(i, 14) = making
    
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 14
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
    MSFlexGrid1.TextMatrix(i, 15) = Round(Val(res1!priceno) / 100)
    'If (Val(dis) <> 20) Then
    'totamt = Val(totamt) * 1.333333
    'totamt = Val(totamt) - (Val(totamt) * Val(dis) / 100)
    'rate1 = (Val(totamt) - (Val(amt2) + Val(amt3) + Val(amt4) + Val(making))) / Val(dweight)
    'MSFlexGrid1.TextMatrix(i, 5) = Round(rate1)
   ' amt1 = Val(rate1) * Val(dweight)
   ' diatotrt = diatotrt + amt1
   ' End If
    diatotrt = diatotrt + res1!dprice * res1!nweight1
    MSFlexGrid1.TextMatrix(i, 16) = Round(res1!ftotamt)
    If (IsNull(res1!opicture)) Then
        MSFlexGrid1.TextMatrix(i, 17) = ""
    ElseIf (res1!opicture <> "" And res1!opicture <> Empty) Then
        MSFlexGrid1.TextMatrix(i, 17) = res1!opicture
    End If
    
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 16
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.CellForeColor = vbBlue
    
    gtotal = gtotal + res1!ftotamt
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
       
    MSFlexGrid1.TextMatrix(i, 4) = Round(dwt, 2)
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 4
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
        
    MSFlexGrid1.TextMatrix(i, 5) = Round(diatotrt)
    
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 5
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
        
    MSFlexGrid1.TextMatrix(i, 7) = Round(stonewt1, 2)
    
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 7
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
    
    MSFlexGrid1.TextMatrix(i, 8) = Round(stonetotamt1)
    
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 8
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
    
    MSFlexGrid1.TextMatrix(i, 10) = Round(stonewt2, 2)
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 10
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
    
    MSFlexGrid1.TextMatrix(i, 11) = Round(stonetotamt2)
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 11
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
    
    MSFlexGrid1.TextMatrix(i, 12) = Round(metal, 2)
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 12
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
    
    MSFlexGrid1.TextMatrix(i, 13) = Round(metalamt)
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 13
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
    
    MSFlexGrid1.TextMatrix(i, 14) = Round(makingtot)
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 14
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
    
    MSFlexGrid1.TextMatrix(i, 16) = Round(gtotal)
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 16
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
    
    End If
    res1.CLOSE
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

