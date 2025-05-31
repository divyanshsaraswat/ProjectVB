VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   5715
   ClientLeft      =   840
   ClientTop       =   2625
   ClientWidth     =   6675
   LinkTopic       =   "Form5"
   ScaleHeight     =   8130
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cprint 
      Caption         =   "&Print"
      Height          =   435
      Left            =   9240
      TabIndex        =   7
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Clear 
      Caption         =   "C&lear"
      Height          =   435
      Left            =   9240
      TabIndex        =   6
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Search 
      Caption         =   "&Search"
      Height          =   375
      Left            =   9240
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton close 
      Caption         =   "&Close"
      Height          =   435
      Left            =   9240
      TabIndex        =   4
      Top             =   1680
      Width           =   1335
   End
   Begin VB.ListBox List2 
      Height          =   960
      ItemData        =   "DCost.frx":0000
      Left            =   4200
      List            =   "DCost.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   120
      Width           =   4335
   End
   Begin VB.TextBox sdate 
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox edate 
      Height          =   315
      Left            =   4560
      TabIndex        =   1
      Top             =   1440
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   960
      ItemData        =   "DCost.frx":0004
      Left            =   1080
      List            =   "DCost.frx":0006
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   6240
      TabIndex        =   8
      Top             =   1440
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yy"
      Format          =   61145089
      CurrentDate     =   37684
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   1440
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yy"
      Format          =   61145089
      CurrentDate     =   37684
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   6375
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   11245
      _Version        =   393216
      Cols            =   26
      RowHeightMin    =   300
      FocusRect       =   2
      HighLight       =   2
      AllowUserResizing=   3
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   26
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Line Line1 
      X1              =   10800
      X2              =   10800
      Y1              =   0
      Y2              =   2760
   End
   Begin VB.Image Image1 
      Height          =   2415
      Left            =   11160
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Invoice."
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   240
      Width           =   975
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   14160
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label Label1 
      Caption         =   "St. Date:"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "End Date:"
      Height          =   375
      Left            =   3480
      TabIndex        =   12
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Party"
      Height          =   255
      Left            =   2880
      TabIndex        =   11
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim res1 As New ADODB.Recordset
Dim res As New ADODB.Recordset
Dim str1, str2 As String

Private Sub Clear_Click()
clearing
sdate = ""
edate = ""
End Sub
Private Sub clearing()
    MSHFlexGrid1.Rows = 2
    MSHFlexGrid1.TextMatrix(1, 0) = ""
    MSHFlexGrid1.TextMatrix(1, 1) = ""
    MSHFlexGrid1.TextMatrix(1, 2) = ""
    MSHFlexGrid1.TextMatrix(1, 3) = ""
    MSHFlexGrid1.TextMatrix(1, 4) = ""
    MSHFlexGrid1.TextMatrix(1, 5) = ""
    MSHFlexGrid1.TextMatrix(1, 6) = ""
    MSHFlexGrid1.TextMatrix(1, 7) = ""
    MSHFlexGrid1.TextMatrix(1, 8) = ""
    MSHFlexGrid1.TextMatrix(1, 9) = ""
    MSHFlexGrid1.TextMatrix(1, 10) = ""
    MSHFlexGrid1.TextMatrix(1, 11) = ""
    MSHFlexGrid1.TextMatrix(1, 12) = ""
    MSHFlexGrid1.TextMatrix(1, 13) = ""
    MSHFlexGrid1.TextMatrix(1, 14) = ""
    MSHFlexGrid1.TextMatrix(1, 15) = ""
    MSHFlexGrid1.TextMatrix(1, 16) = ""
    MSHFlexGrid1.TextMatrix(1, 17) = ""
    MSHFlexGrid1.TextMatrix(1, 18) = ""
    MSHFlexGrid1.TextMatrix(1, 19) = ""
    MSHFlexGrid1.TextMatrix(1, 20) = ""
    MSHFlexGrid1.TextMatrix(1, 21) = ""
    MSHFlexGrid1.TextMatrix(1, 22) = ""
    MSHFlexGrid1.TextMatrix(1, 23) = ""
    MSHFlexGrid1.TextMatrix(1, 24) = ""
    MSHFlexGrid1.TextMatrix(1, 25) = ""

End Sub
Private Sub close_Click()
Unload Me
End Sub

Private Sub cprint_Click()
On Error Resume Next
Dim objExcel As Excel.Application, objBook As Excel.Workbook, objsheet As Excel.Worksheet

Set objExcel = GetObject(, "Excel.Application") 'if excel already open you can use GetObject
If Err.Number Then
Err.Clear
End If

Set objExcel = CreateObject("Excel.Application") 'or CreateObject to open new Excel Application
Set objBook = objExcel.Workbooks.Open("c:\\daksh\vb\Sales Details.xls")
Set objsheet = objBook.Worksheets(1)


Dim i

i = 3
For j = 1 To MSHFlexGrid1.Rows - 1
objsheet.Cells(i, 1) = MSHFlexGrid1.TextMatrix(j, 0)
objsheet.Cells(i, 2) = MSHFlexGrid1.TextMatrix(j, 1)
objsheet.Cells(i, 3) = MSHFlexGrid1.TextMatrix(j, 2)
objsheet.Cells(i, 4) = MSHFlexGrid1.TextMatrix(j, 3)
objsheet.Cells(i, 5) = MSHFlexGrid1.TextMatrix(j, 4)
objsheet.Cells(i, 6) = MSHFlexGrid1.TextMatrix(j, 5)
objsheet.Cells(i, 7) = MSHFlexGrid1.TextMatrix(j, 6)
objsheet.Cells(i, 8) = MSHFlexGrid1.TextMatrix(j, 7)
objsheet.Cells(i, 9) = MSHFlexGrid1.TextMatrix(j, 8)
objsheet.Cells(i, 10) = MSHFlexGrid1.TextMatrix(j, 9)
objsheet.Cells(i, 11) = MSHFlexGrid1.TextMatrix(j, 10)
objsheet.Cells(i, 12) = MSHFlexGrid1.TextMatrix(j, 11)
objsheet.Cells(i, 13) = MSHFlexGrid1.TextMatrix(j, 12)
objsheet.Cells(i, 14) = MSHFlexGrid1.TextMatrix(j, 13)
objsheet.Cells(i, 15) = MSHFlexGrid1.TextMatrix(j, 14)
objsheet.Cells(i, 16) = MSHFlexGrid1.TextMatrix(j, 15)
objsheet.Cells(i, 17) = MSHFlexGrid1.TextMatrix(j, 16)
objsheet.Cells(i, 18) = MSHFlexGrid1.TextMatrix(j, 18)
objsheet.Cells(i, 19) = MSHFlexGrid1.TextMatrix(j, 19)
objsheet.Cells(i, 20) = MSHFlexGrid1.TextMatrix(j, 20)
i = i + 1
Next j
''objBook.save
objExcel.Visible = True
objBook.PrintPreview
''objBook.CLOSE
' cell(1,1) means cell A1 ;
''Set objsheet = Null
''Set objExcel = Null

End Sub

Private Sub DTPicker1_Change()
sdate = DTPicker1.Value
End Sub
Private Sub DTPicker2_Change()
edate = DTPicker2.Value
End Sub

Private Sub Form_Load()
loaditem
MSHFlexGrid1.TextMatrix(0, 0) = "S.No."
MSHFlexGrid1.ColWidth(0) = 350
MSHFlexGrid1.TextMatrix(0, 1) = "Item Code"
MSHFlexGrid1.ColWidth(1) = 700
MSHFlexGrid1.TextMatrix(0, 2) = "Invoice."
MSHFlexGrid1.ColWidth(2) = 700
MSHFlexGrid1.TextMatrix(0, 3) = "Party Name"
MSHFlexGrid1.ColWidth(3) = 1500
MSHFlexGrid1.TextMatrix(0, 4) = "Through"
MSHFlexGrid1.ColWidth(4) = 1250
MSHFlexGrid1.TextMatrix(0, 5) = "Dia.Wt."
MSHFlexGrid1.ColWidth(5) = 750
MSHFlexGrid1.TextMatrix(0, 6) = "Dia.Rt."
MSHFlexGrid1.ColWidth(6) = 950
MSHFlexGrid1.TextMatrix(0, 7) = "Stone1."
MSHFlexGrid1.ColWidth(7) = 850
MSHFlexGrid1.TextMatrix(0, 8) = "St1 Wt."
MSHFlexGrid1.ColWidth(8) = 750
MSHFlexGrid1.TextMatrix(0, 9) = "St1 Rt."
MSHFlexGrid1.ColWidth(9) = 750
MSHFlexGrid1.TextMatrix(0, 10) = "Stone2."
MSHFlexGrid1.ColWidth(10) = 850
MSHFlexGrid1.TextMatrix(0, 11) = "St1 Wt."
MSHFlexGrid1.ColWidth(11) = 750
MSHFlexGrid1.TextMatrix(0, 12) = "St1 Rt."
MSHFlexGrid1.ColWidth(12) = 750
MSHFlexGrid1.TextMatrix(0, 13) = "MetalWt."
MSHFlexGrid1.ColWidth(13) = 600
MSHFlexGrid1.TextMatrix(0, 14) = "MetatRt."
MSHFlexGrid1.ColWidth(14) = 850
MSHFlexGrid1.TextMatrix(0, 15) = "Making"
MSHFlexGrid1.ColWidth(15) = 850
MSHFlexGrid1.TextMatrix(0, 16) = "Total Amt."
MSHFlexGrid1.ColWidth(16) = 1200
MSHFlexGrid1.TextMatrix(0, 17) = "Picture"
MSHFlexGrid1.ColWidth(17) = 0
MSHFlexGrid1.TextMatrix(0, 18) = "NDNO"
MSHFlexGrid1.ColWidth(18) = 0
MSHFlexGrid1.TextMatrix(0, 19) = "CHPCODE"
MSHFlexGrid1.ColWidth(19) = 0
MSHFlexGrid1.TextMatrix(0, 20) = "CHOWNER"
MSHFlexGrid1.ColWidth(20) = 0
MSHFlexGrid1.TextMatrix(0, 21) = "Dia Min. RT"
MSHFlexGrid1.ColWidth(21) = 0
MSHFlexGrid1.TextMatrix(0, 22) = "Col.Min RT1"
MSHFlexGrid1.ColWidth(22) = 0
MSHFlexGrid1.TextMatrix(0, 23) = "Col.Min RT2"
MSHFlexGrid1.ColWidth(23) = 0
MSHFlexGrid1.TextMatrix(0, 24) = "Making"
MSHFlexGrid1.ColWidth(24) = 0
MSHFlexGrid1.TextMatrix(0, 25) = "Metal Rt."
MSHFlexGrid1.ColWidth(25) = 0
End Sub
Private Sub loaditem()
DTPicker1.Format = dtpCustom
DTPicker2.Format = dtpCustom
str1 = "Select distinct(chissue) from tblsale order by chissue"
res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
If (List2.ListCount > 0) Then
  List2.Clear
End If
While (res1.EOF = False)
List2.AddItem (res1!chissue)
 res1.MoveNext
Wend

res1.close

str1 = "Select distinct(nappno) from tblsale order by nappno"
res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
If (List1.ListCount > 0) Then
List1.Clear
End If
While (res1.EOF = False)
List1.AddItem (res1!nappno)
'appno.AddItem (res1!nappno)
res1.MoveNext
Wend
res1.close

str1 = "Select distinct(chthrough) from tblsale order by chthrough"
res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
through.Clear
While (res1.EOF = False)
through.AddItem (res1!chthrough)
'appno.AddItem (res1!nappno)
res1.MoveNext
Wend
res1.close

str1 = "Select distinct(chcategory) from tblsale order by chcategory"
res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
Category.Clear
While (res1.EOF = False)
Category.AddItem (res1!chcategory)
'appno.AddItem (res1!nappno)
res1.MoveNext
Wend
res1.close

End Sub

Private Sub MSHFlexGrid1_DblClick()
Load Invoice
Invoice.qut.Text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 2)
Invoice.load_Click
Invoice.Show
End Sub

Private Sub MSHFlexGrid1_SelChange()
On Error Resume Next
If (MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 17) <> "" And MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 17) <> Empty) Then
     Dim picturename As String
     pos = InStrRev(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 17), "\")
     If (pos <> 0) Then
     picturename = Mid(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 17), pos + 1)
     Else
     picturename = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 17)
     End If
     Image1.Visible = True
     Image1.Picture = LoadPicture(MDIForm1.picturepath & picturename)
     'Image1.Picture = LoadPicture(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 17))
Else
    Image1.Picture = LoadPicture()
End If

End Sub

Private Sub Search_Click()
''On Error Resume Next
str2 = ""
str1 = "select * from tblissuerec "
Dim listdt, listapp
listdt = ""
listapp = ""

Dim issuein As String
    For Id = 0 To List2.ListCount - 1

     If (List2.Selected(Id) <> False) Then
     listdt = listdt & "'" & List2.List(Id) & "',"
     End If
    Next
    'MsgBox (listdt)
    
    If (listdt <> "") Then
     pos = InStrRev(listdt, ",")
     issuein = Mid(listdt, 1, pos - 1)
     str2 = str2 & " chissue in(" & issuein & ")"
    End If

For Id = 0 To List1.ListCount - 1
If (List1.Selected(Id) <> False) Then
     listapp = listapp & List1.List(Id) & ","
     End If
    Next

If (listapp <> "") Then
     pos = InStrRev(listapp, ",")
     listapp = Mid(listapp, 1, pos - 1)
    If (str2 <> Empty) Then
     str2 = str2 & " And tblsale.nappno in(" & listapp & ")"
     Else
     str2 = str2 & " tblsale.nappno in(" & listapp & ")"
    End If
End If


If (Category.Text <> Empty) Then
    If (str2 <> Empty) Then
    str2 = str2 & " And tblsale.chcategory=" & "'" & Category.Text & "'"
    Else
    str2 = str2 & " tblsale.chcategory=" & "'" & Category.Text & "'"
    End If
End If

If (sowner.Text <> Empty) Then
    If (str2 <> Empty) Then
    str2 = str2 & " And tblsale.chowner=" & "'" & MDIForm1.parserText(sowner.Text) & "'"
    Else
    str2 = str2 & " tblsale.chowner=" & "'" & MDIForm1.parserText(sowner.Text) & "'"
    End If
End If


If (through.Text <> Empty) Then
    If (str2 <> Empty) Then
    str2 = str2 & " And tblsale.chthrough=" & "'" & MDIForm1.parserText(through.Text) & "'"
    Else
    str2 = str2 & " tblsale.chthrough=" & "'" & MDIForm1.parserText(through.Text) & "'"
    End If
End If

 If (sdate.Text <> Empty) Then
    If (str2 <> Empty) Then
    str2 = str2 & " And tblsale.ddate>= #" & Format(sdate.Text, "dd-mmm-yy") & "# "
    Else
    str2 = str2 & " tblsale.ddate >= #" & Format(sdate.Text, "dd-mmm-yy") & "# "
    End If
  End If

    If (edate.Text <> Empty) Then
        If (str2 <> Empty) Then
            str2 = str2 & " And tblsale.ddate<=#" & Format(edate.Text, "dd-mmm-yy") & "# "
        Else
            str2 = str2 & " tblsale.ddate<=#" & Format(edate.Text, "dd-mmm-yy") & "# "
        End If
    End If

If (str2 <> Empty) Then
str1 = str1 & " where " & str2 & " order by tblsale.nappno,tblsale.chcategory,tblsale.ncode"
Else
str1 = str1 & " order by tblsale.nappno,tblsale.chcategory,tblsale.ncode"
End If

Dim gtotal As Double
Dim dwt As Double
Dim metal As Double
Dim metalamt As Double
Dim stonewt1 As Double
Dim stonewt2 As Double
Dim makingtot As Double
Dim diatotrt As Double
Dim stonetotamt1 As Double


res.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
Dim i
i = 1
While (res.EOF = False)
    MSHFlexGrid1.TextMatrix(i, 0) = i
    MSHFlexGrid1.TextMatrix(i, 1) = res!chcategory & "-" & res!ncode
    MSHFlexGrid1.TextMatrix(i, 2) = res!nappno
    MSHFlexGrid1.TextMatrix(i, 3) = res!chissue
    If (IsNull(res!chthrough) = False) Then
      MSHFlexGrid1.TextMatrix(i, 4) = res!chthrough
    End If
    
    If (IsNull(res!minrate1) = False) Then
      MSHFlexGrid1.TextMatrix(i, 6) = res!minrate1
    End If
    If (IsNull(res!minrate2) = False) Then
    MSHFlexGrid1.TextMatrix(i, 9) = res!minrate2
    End If
    If (IsNull(res!minrate3) = False) Then
    MSHFlexGrid1.TextMatrix(i, 12) = res!minrate3
    End If
    If (IsNull(res!minrate4) = False) Then
    MSHFlexGrid1.TextMatrix(i, 14) = res!minrate4
    End If
    dweight = res!nweight1
    MSHFlexGrid1.TextMatrix(i, 5) = dweight
    swt1 = res!nweight2
    MSHFlexGrid1.TextMatrix(i, 8) = swt1
    swt2 = res!nweight3
    MSHFlexGrid1.TextMatrix(i, 11) = swt2
    mwt = res!nweight4
    MSHFlexGrid1.TextMatrix(i, 13) = Round(mwt, 2)
    
    MSHFlexGrid1.TextMatrix(i, 7) = res!chcontent1
    MSHFlexGrid1.TextMatrix(i, 10) = res!chcontent2
    
    If (IsNull(res!opicture) = False) Then
    MSHFlexGrid1.TextMatrix(i, 17) = res!opicture
    End If
  
    ''MSHFlexGrid1.TextMatrix(i, 14) = res!npno
    
    If (IsNull(res!ndno) = False) Then
    MSHFlexGrid1.TextMatrix(i, 18) = res!ndno
    End If
    
   If (IsNull(res!chpcode) = False) Then
    MSHFlexGrid1.TextMatrix(i, 19) = res!chpcode
   End If
   
    If (IsNull(res!chowner) = False) Then
    MSHFlexGrid1.TextMatrix(i, 20) = res!chowner
    End If
    
        
    If (IsNull(res!nrate1) = False) Then
        MSHFlexGrid1.TextMatrix(i, 21) = res!nrate1
    End If
    If (IsNull(res!nrate2) = False) Then
        MSHFlexGrid1.TextMatrix(i, 22) = res!nrate2
    End If
    If (IsNull(res!nrate3) = False) Then
    MSHFlexGrid1.TextMatrix(i, 23) = res!nrate3
    End If
    If (IsNull(res!nmaking) = False) Then
    
    MSHFlexGrid1.TextMatrix(i, 24) = res!nmaking
    End If
    
    If (IsNull(res!nrate4) = False) Then
    MSHFlexGrid1.TextMatrix(i, 25) = res!nrate4
    End If

   ' If (res!opicture <> "") Then
   ' MSHFlexGrid1.Row = I
   ' MSHFlexGrid1.Col = 14
   ' Set MSHFlexGrid1.CellPicture = LoadPicture(MDIForm1.picturepath & picturename)
   ' End If
    
    If (IsNull(res!nmaking1) = False) Then
    making = res!nmaking1
    End If
    MSHFlexGrid1.TextMatrix(i, 15) = making
    amt1 = Val(MSHFlexGrid1.TextMatrix(i, 6)) * Val(res!nweight1)
    amt2 = Val(MSHFlexGrid1.TextMatrix(i, 9)) * Val(res!nweight2)
    amt3 = Val(MSHFlexGrid1.TextMatrix(i, 12)) * Val(res!nweight3)
    amt4 = Val(MSHFlexGrid1.TextMatrix(i, 14)) * Val(res!nweight4)
    metalamt = metalamt + amt4
    stonetotamt1 = stonetotamt1 + amt2
    stonetotamt2 = stonetotamt2 + amt3
    diatotrt = diatotrt + amt1
    totamt = Val(amt1) + Val(amt2) + Val(amt3) + Val(amt4) + Val(MSHFlexGrid1.TextMatrix(i, 15))
    MSHFlexGrid1.TextMatrix(i, 16) = Round(totamt)
    pno = res!npno
    ''MSHFlexGrid1.TextMatrix(i, 14) = pno
     ''MSHFlexGrid1.TextMatrix(i, 15) = Round(100 - (Val(totamt) / (Val(pno) * 100)) * 100, 1)
     ''gpno = gpno + pno
    gtotal = gtotal + totamt
    dwt = dwt + dweight
    metal = metal + mwt
    stonewt1 = stonewt1 + swt1
    stonewt2 = stonewt2 + swt2
    makingtot = makingtot + making
    res.MoveNext
    i = i + 1
    MSHFlexGrid1.Rows = i + 1
    Wend
    MSHFlexGrid1.TextMatrix(i, 1) = "Grand Total:"
    MSHFlexGrid1.TextMatrix(i, 5) = Round(dwt, 2)
    MSHFlexGrid1.TextMatrix(i, 6) = Round(diatotrt)
    MSHFlexGrid1.TextMatrix(i, 8) = Round(stonewt1, 2)
    MSHFlexGrid1.TextMatrix(i, 9) = Round(stonetotamt1)
    MSHFlexGrid1.TextMatrix(i, 11) = Round(stonewt2, 2)
    MSHFlexGrid1.TextMatrix(i, 12) = Round(stonetotamt2)
    MSHFlexGrid1.TextMatrix(i, 13) = Round(metal, 2)
    MSHFlexGrid1.TextMatrix(i, 14) = Round(metalamt)
    MSHFlexGrid1.TextMatrix(i, 15) = Round(makingtot)
    MSHFlexGrid1.TextMatrix(i, 16) = Round(gtotal)
    ''MSHFlexGrid1.TextMatrix(i, 14) = Round(gpno)
  
  res.close

End Sub

