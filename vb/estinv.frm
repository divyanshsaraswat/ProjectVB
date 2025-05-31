VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form estinv 
   Caption         =   "Invoice"
   ClientHeight    =   5715
   ClientLeft      =   2370
   ClientTop       =   1545
   ClientWidth     =   6585
   LinkTopic       =   "Form8"
   MDIChild        =   -1  'True
   ScaleHeight     =   5715
   ScaleWidth      =   6585
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   600
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton Update 
      Caption         =   "&Update"
      Height          =   375
      Left            =   4200
      TabIndex        =   21
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton Print 
      Caption         =   "&Print"
      Height          =   375
      Left            =   9600
      TabIndex        =   20
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton save 
      Caption         =   "&Save"
      Height          =   375
      Left            =   2880
      TabIndex        =   19
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton addnew 
      Caption         =   "A&dd New"
      Height          =   375
      Left            =   240
      TabIndex        =   18
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton load 
      Caption         =   "&Load"
      Height          =   375
      Left            =   1560
      TabIndex        =   17
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton CLOSE 
      Caption         =   "C&lose"
      Height          =   375
      Left            =   8160
      TabIndex        =   16
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton clear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   6840
      TabIndex        =   15
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton delete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   5520
      TabIndex        =   14
      Top             =   7080
      Width           =   1095
   End
   Begin VB.TextBox through 
      Height          =   315
      Left            =   7920
      TabIndex        =   12
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox code 
      Height          =   315
      Left            =   1440
      TabIndex        =   6
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton add 
      Caption         =   "&Add"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox party 
      Height          =   315
      Left            =   5400
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox qut 
      Height          =   315
      Left            =   960
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton rmv 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.ComboBox dis 
      Height          =   315
      ItemData        =   "estinv.frx":0000
      Left            =   5640
      List            =   "estinv.frx":001C
      TabIndex        =   0
      Text            =   "20%"
      Top             =   840
      Width           =   735
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yy"
      Format          =   24641537
      CurrentDate     =   37684
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   5055
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   8916
      _Version        =   393216
      Cols            =   17
      RowHeightMin    =   300
      HighLight       =   2
      AllowUserResizing=   3
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   17
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   9360
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1935
   End
   Begin VB.Line Line3 
      X1              =   240
      X2              =   12000
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   9240
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   9240
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label4 
      Caption         =   "Through"
      Height          =   255
      Left            =   7080
      TabIndex        =   11
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Party Name:"
      Height          =   255
      Left            =   4320
      TabIndex        =   10
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Date:"
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   240
      Width           =   615
   End
   Begin VB.Label icode 
      Caption         =   "S. NO.:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Qut. No."
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "estinv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim res1 As New ADODB.Recordset
Dim res As New ADODB.Recordset
Dim I As Integer
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
Dim gpno As Double
Dim flag As Boolean
Dim celltotal As Double
Dim row As Integer
Dim col As Integer
Private Sub add_Click()
On Error Resume Next
Dim num As Integer
For j = 0 To I - 1
pos = InStr(1, MSHFlexGrid1.TextMatrix(j, 1), "-", vbTextCompare)
num = Val(Mid(MSHFlexGrid1.TextMatrix(j, 1), pos + 1))
If (num = Val(Code.Text)) Then
MsgBox ("Item All Ready Added In The List.")
Exit Sub
End If
Next
str1 = "select Tblcosting.*, Tblcosting.minrate1 * tblcosting.nweight1 as amt1,Tblcosting.minrate2 * tblcosting.nweight2 as amt2,Tblcosting.minrate3 * tblcosting.nweight3 as amt3,"
str1 = str1 & " Tblcosting.minrate4 * tblcosting.nweight4 as amt4,amt1+amt2+amt3+amt4+nmaking1 AS totamt,totamt * 1.25 as priceno,priceno-((priceno)/100 * " & Val(dis.Text)
str1 = str1 & ") as ftotamt,(ftotamt-(amt2+amt3+amt4+nmaking1))/nweight1 as dprice,dprice*Tblcosting.nweight1 as damt "
str1 = str1 & " from tblcosting  where ncode= " & Val(Code.Text)

'str1 = "SELECT * from tblcosting Where ncode ="
'If (Option1(2).Value = False) Then
''If (I = 22) Then
'    MsgBox ("You Can't Add More Than 21 Items In The One Form")
'Exit Sub
'End If
'End If

If (Val(Code.Text) <= 0) Then
    MsgBox ("Enter the Valid S.No")
Else
'str1 = str1 & Val(code.Text) & " and ncode not in (Select ncode from tblappdetail order by ncode) "
res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic

'If (res1.EOF = True) Then
'    res1.CLOSE
'    str1 = "Select chissue,nappno from tblappmaster where nappno="
 '   str1 = str1 & " ( Select  nappno from tblappdetail where ncode=" & Val(code.Text) & ") "
'    res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
 '   name1 = ""
 ' If (res1.EOF = False) Then
   ' result = MsgBox("This Item Found in the Approval No " & res1!nappno & " And Issue to " & res1!chissue & ". Do You Want To add This? ", vbYesNo + vbDefaultButton2, "Confirmation")

  '  If (Val(result) = 7) Then
   '     res1.CLOSE
   '     code.Text = ""
   '     code.SetFocus
   '     Exit Sub
  '  Else
  '      str1 = " Select  * from tblappdetail where ncode=" & Val(code.Text)
  '      res1.CLOSE
  '      res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
  '      res1.delete
  '      res1.CLOSE
   '     add_Click
  '  End If
'  Else
   
   If (res1.EOF = True) Then
    MsgBox ("Item Not Found ")
    res1.close
    Code.Text = ""
    Code.SetFocus
    Exit Sub
   End If

Code.Text = ""
Code.SetFocus
   
While (res1.EOF = False)
    
    MSHFlexGrid1.TextMatrix(I, 0) = I
    MSHFlexGrid1.TextMatrix(I, 1) = res1!chcategory & "-" & res1!ncode
    dweight = res1!nweight1
    MSHFlexGrid1.TextMatrix(I, 2) = dweight
'    MSHFlexgrid1.TextMatrix(i, 3) = res1!minrate1
     MSHFlexGrid1.TextMatrix(I, 3) = Round(res1!dprice)
     
    MSHFlexGrid1.TextMatrix(I, 4) = res1!chcontent2
    swt1 = res1!nweight2
    MSHFlexGrid1.TextMatrix(I, 5) = swt1
    MSHFlexGrid1.TextMatrix(I, 6) = res1!minrate2
    MSHFlexGrid1.TextMatrix(I, 7) = res1!chcontent3
    swt2 = res1!nweight3
    MSHFlexGrid1.TextMatrix(I, 8) = swt2
    MSHFlexGrid1.TextMatrix(I, 9) = res1!minrate3
    mwt = res1!nweight4
    MSHFlexGrid1.TextMatrix(I, 10) = mwt
    MSHFlexGrid1.TextMatrix(I, 11) = res1!minrate4
    making = res1!nmaking1
    MSHFlexGrid1.TextMatrix(I, 12) = making
   
    
    If (IsNull(res1!opicture) = False) Then
    
    Dim picturename As String
    pos = InStrRev(res1!opicture, "\")
    
    If (pos <> 0) Then
    picturename = Mid(res1!opicture, pos + 1)
    Else
    picturename = res1!opicture
    End If
     MSHFlexGrid1.TextMatrix(I, 16) = picturename
    End If
    
   ' If (picturename <> "") Then
  '  Dim stdPic As StdPicture
  '  Dim d As Single
  '  Set stdPic = LoadPicture(MDIForm1.picturepath & picturename, 0)
    
   ' Image1.Width = 1815
  '  Image1.Height = 615
   ' Image1.Picture = stdPic
    
    
 '   MSHFlexGrid1.Row = I
   ' MSHFlexGrid1.Col = 14
    'Set Image2(I).Picture = LoadPicture(MDIForm1.picturepath & picturename)
  ' ' load Image1(I - 1)
   ' With Image1(I - 1)
 '   .Picture = LoadPicture(MDIForm1.picturepath & picturename)
    '.Stretch = True
  '  .Width = 750
  '  .Height = MSHFlexGrid1.CellHeight
  '  .top = 1350 + MSHFlexGrid1.CellTop
  '  .Left = 10500 '50 + MSHFlexGrid1.CellLeft ' + MSHFlexGrid1.Left
   ' .Visible = True
    
 ' End With
 

   ' Set MSHFlexGrid1.CellPicture = Image2(I)
   ' End If
    'amt1 = Val(res1!minrate1) * Val(res1!nweight1)
    'amt2 = Val(res1!minrate2) * Val(res1!nweight2)
    'amt3 = Val(res1!minrate3) * Val(res1!nweight3)
    'amt4 = Val(res1!minrate4) * Val(res1!nweight4)
    
    metalamt = metalamt + res1!amt4
    stonetotamt1 = stonetotamt1 + res1!amt2
    stonetotamt2 = stonetotamt2 + res1!amt3
    diatotrt = diatotrt + res1!damt
    totamt = Val(res1!damt) + Val(res1!amt2) + Val(res1!amt3) + Val(res1!amt4) + Val(res1!nmaking1)
    MSHFlexGrid1.TextMatrix(I, 13) = Round(totamt)
    pno = res1!priceno
    MSHFlexGrid1.TextMatrix(I, 14) = Round(pno / 100)
     MSHFlexGrid1.TextMatrix(I, 15) = Round(100 - (Val(totamt) / (Val(pno) * 100)) * 100, 1)
    gtotal = gtotal + totamt
    dwt = dwt + dweight
    gpno = gpno + Round(pno / 100)
    metal = metal + mwt
    stonewt1 = stonewt1 + swt1
    stonewt2 = stonewt2 + swt2
    
    makingtot = makingtot + making
    res1.MoveNext
    I = I + 1
    MSHFlexGrid1.Rows = I + 1
    Wend

    MSHFlexGrid1.TextMatrix(I, 1) = "Grand Total:"
    MSHFlexGrid1.TextMatrix(I, 2) = Round(dwt, 2)
    MSHFlexGrid1.TextMatrix(I, 3) = Round(diatotrt)
    MSHFlexGrid1.TextMatrix(I, 5) = Round(stonewt1, 2)
    MSHFlexGrid1.TextMatrix(I, 6) = Round(stonetotamt1)
    MSHFlexGrid1.TextMatrix(I, 8) = Round(stonewt2, 2)
    
    MSHFlexGrid1.TextMatrix(I, 9) = Round(stonetotamt2)
    MSHFlexGrid1.TextMatrix(I, 10) = Round(metal, 2)
    MSHFlexGrid1.TextMatrix(I, 11) = Round(metalamt)
    
    MSHFlexGrid1.TextMatrix(I, 12) = Round(makingtot)
    MSHFlexGrid1.TextMatrix(I, 13) = Round(gtotal)
    MSHFlexGrid1.TextMatrix(I, 14) = Round(gpno)
   res1.close

End If

End Sub

Private Sub Addnew_Click()
On Error Resume Next
clearing
res1.Open "select max(nappno) from tblqut ", MDIForm1.con1, adOpenDynamic, adLockOptimistic
If (IsNull(res1.Fields(0))) Then
qut = 1
Else
qut = res1.Fields(0) + 1
End If
res1.close
DTPicker1 = Format(date, "dd-mmm-yy")
'apdate = Format(Date, "dd-mmm-yy")
save.Enabled = True
load.Enabled = False
update.Enabled = False
delete.Enabled = False
End Sub

Private Sub Clear_Click()
On Error Resume Next
clearing
save.Enabled = False
delete.Enabled = True
load.Enabled = True
update.Enabled = True
End Sub
Private Sub clearing()
DTPicker1 = Format(date, "dd-mmm-yy")
through = ""
I = 1
dwt = 0
metal = 0
stonewt1 = 0
stonewt2 = 0
gtotal = 0
metalamt = 0
makingtot = 0
diatotrt = 0
stonetotamt1 = 0
stonetotamt2 = 0
party = ""
Code = ""
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
End Sub
Private Sub close_Click()
Unload Me
End Sub

Private Sub code_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
add_Click
End If
End Sub

Private Sub delete_Click()
On Error Resume Next
deleterec
clearing
delete.Enabled = False
End Sub

Private Sub deleterec()
On Error Resume Next
str1 = "Select * from tblqut where nappno=" & Val(qut)
res.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
If (res.EOF = True) Then
MsgBox ("NO RECORD FOUND!")
Exit Sub
End If
If (res.RecordCount > 0) Then
res.MoveFirst
End If
Do While Not res.EOF
    res.delete
    res.MoveNext
Loop
   MsgBox ("Record Deleted")
   res.close
End Sub
Private Sub Form_Activate()
row = 5000
col = 5000
qut.SetFocus
Text1.Visible = False
End Sub

Private Sub Form_Load()
I = 1
DTPicker1.Format = dtpCustom
'apdate = Format(Date, "dd-mmm-yy")
save.Enabled = False
MSHFlexGrid1.TextMatrix(0, 0) = "S.No."
MSHFlexGrid1.ColWidth(0) = 650
MSHFlexGrid1.TextMatrix(0, 1) = "Item Code"
MSHFlexGrid1.ColWidth(1) = 850
MSHFlexGrid1.TextMatrix(0, 2) = "Dia.Wt."
MSHFlexGrid1.ColWidth(2) = 650
MSHFlexGrid1.TextMatrix(0, 3) = "Dia.Rt."
MSHFlexGrid1.ColWidth(3) = 850
MSHFlexGrid1.TextMatrix(0, 4) = "Stone1."
MSHFlexGrid1.ColWidth(4) = 900
MSHFlexGrid1.TextMatrix(0, 5) = "St1 Wt."
MSHFlexGrid1.ColWidth(5) = 650
MSHFlexGrid1.TextMatrix(0, 6) = "St1 Rt."
MSHFlexGrid1.ColWidth(6) = 850
MSHFlexGrid1.TextMatrix(0, 7) = "Stone2"
MSHFlexGrid1.ColWidth(7) = 0
MSHFlexGrid1.TextMatrix(0, 8) = "St2 Wt."
MSHFlexGrid1.ColWidth(8) = 0
MSHFlexGrid1.TextMatrix(0, 9) = "St2 Rt."
MSHFlexGrid1.ColWidth(9) = 0
MSHFlexGrid1.TextMatrix(0, 10) = "MetalWt."
MSHFlexGrid1.ColWidth(10) = 700
MSHFlexGrid1.TextMatrix(0, 11) = "MetalRt."
MSHFlexGrid1.ColWidth(11) = 650
MSHFlexGrid1.TextMatrix(0, 12) = "Making."
MSHFlexGrid1.ColWidth(12) = 900
MSHFlexGrid1.TextMatrix(0, 13) = "Total."
MSHFlexGrid1.ColWidth(13) = 1200
MSHFlexGrid1.TextMatrix(0, 14) = "P.No"
MSHFlexGrid1.ColWidth(14) = 850
MSHFlexGrid1.TextMatrix(0, 15) = "Dis.(%)"
MSHFlexGrid1.ColWidth(15) = 650
MSHFlexGrid1.TextMatrix(0, 16) = ""
MSHFlexGrid1.ColWidth(16) = 0
'con1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\lahar\vb\db2.mdb;Persist Security Info=False"
'con1.Open
End Sub

Private Sub load_Click()
On Error Resume Next
str1 = "select * from tblqut where nappno=" & Val(qut.Text)
If (res.state <> 0) Then
res.close
End If
res.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
If (res.EOF = True) Then
 MsgBox ("Record Not Found")
res.close
Else
clearing
loading
End If
End Sub
Private Sub loading()
'apdate = Format(res!dappdate, "dd-mmm-yy")
On Error Resume Next
DTPicker1 = Format(res!ddate, "dd-mmm-yy")
party = res!chissue
If (IsNull(res!chthrough)) Then
through = ""
Else
through = res!chthrough
End If
'res.CLOSE
'str1 = "select * from tblqut where nappno= " & Val(qut.Text)
'res.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
'Clearing
I = 1
While (res.EOF = False)
    MSHFlexGrid1.TextMatrix(I, 0) = I
    MSHFlexGrid1.TextMatrix(I, 1) = res!chcategory & "-" & res!ncode
    If (IsNull(res!nrate1) = False) Then
      MSHFlexGrid1.TextMatrix(I, 3) = res!nrate1
    End If
    If (IsNull(res!nrate2) = False) Then
    MSHFlexGrid1.TextMatrix(I, 6) = res!nrate2
    End If
    If (IsNull(res!nrate3) = False) Then
    MSHFlexGrid1.TextMatrix(I, 9) = res!nrate3
    End If
    If (IsNull(res!nrate4) = False) Then
    MSHFlexGrid1.TextMatrix(I, 11) = res!nrate4
    End If
    dweight = res!nweight1
    MSHFlexGrid1.TextMatrix(I, 2) = dweight
    swt1 = res!nweight2
    MSHFlexGrid1.TextMatrix(I, 5) = swt1
    swt2 = res!nweight3
    MSHFlexGrid1.TextMatrix(I, 8) = swt2
    mwt = res!nweight4
    MSHFlexGrid1.TextMatrix(I, 10) = Round(mwt, 2)
    
    MSHFlexGrid1.TextMatrix(I, 4) = res!chcontent1
    MSHFlexGrid1.TextMatrix(I, 7) = res!chcontent2
    'MSHFlexGrid1.TextMatrix(I, 14) = res!npno
    If (IsNull(res!opicture) = False) Then
    MSHFlexGrid1.TextMatrix(I, 16) = res!opicture
    End If
   ' If (res!opicture <> "") Then
   ' MSHFlexGrid1.Row = I
   ' MSHFlexGrid1.Col = 14
   ' Set MSHFlexGrid1.CellPicture = LoadPicture(MDIForm1.picturepath & picturename)
   ' End If
    
    If (IsNull(res!nmaking1) = False) Then
    making = res!nmaking1
    End If
    MSHFlexGrid1.TextMatrix(I, 12) = making
    amt1 = Val(MSHFlexGrid1.TextMatrix(I, 3)) * Val(res!nweight1)
    amt2 = Val(MSHFlexGrid1.TextMatrix(I, 6)) * Val(res!nweight2)
    amt3 = Val(MSHFlexGrid1.TextMatrix(I, 9)) * Val(res!nweight3)
    amt4 = Val(MSHFlexGrid1.TextMatrix(I, 11)) * Val(res!nweight4)
    metalamt = metalamt + amt4
    stonetotamt1 = stonetotamt1 + amt2
    stonetotamt2 = stonetotamt2 + amt3
    diatotrt = diatotrt + amt1
    totamt = Val(amt1) + Val(amt2) + Val(amt3) + Val(amt4) + Val(MSHFlexGrid1.TextMatrix(I, 12))
    MSHFlexGrid1.TextMatrix(I, 13) = Round(totamt)
    pno = res!npno
    MSHFlexGrid1.TextMatrix(I, 14) = pno
     MSHFlexGrid1.TextMatrix(I, 15) = Round(100 - (Val(totamt) / (Val(pno) * 100)) * 100, 1)
     gpno = gpno + pno
    gtotal = gtotal + totamt
    dwt = dwt + dweight
    metal = metal + mwt
    stonewt1 = stonewt1 + swt1
    stonewt2 = stonewt2 + swt2
    makingtot = makingtot + making
    res.MoveNext
    I = I + 1
    MSHFlexGrid1.Rows = I + 1
    Wend
    MSHFlexGrid1.TextMatrix(I, 1) = "Grand Total:"
    MSHFlexGrid1.TextMatrix(I, 2) = Round(dwt, 2)
    MSHFlexGrid1.TextMatrix(I, 3) = Round(diatotrt)
    MSHFlexGrid1.TextMatrix(I, 5) = Round(stonewt1, 2)
    MSHFlexGrid1.TextMatrix(I, 6) = Round(stonetotamt1)
    MSHFlexGrid1.TextMatrix(I, 8) = Round(stonewt2, 2)
    MSHFlexGrid1.TextMatrix(I, 9) = Round(stonetotamt2)
    MSHFlexGrid1.TextMatrix(I, 10) = Round(metal, 2)
    MSHFlexGrid1.TextMatrix(I, 11) = Round(metalamt)
    MSHFlexGrid1.TextMatrix(I, 12) = Round(makingtot)
    MSHFlexGrid1.TextMatrix(I, 13) = Round(gtotal)
    MSHFlexGrid1.TextMatrix(I, 14) = Round(gpno)
  
  res.close
End Sub




Private Sub MSHFlexGrid1_DblClick()
On Error Resume Next

If (MSHFlexGrid1.col <> 1 And MSHFlexGrid1.col <> 2 And MSHFlexGrid1.col <> 4 And MSHFlexGrid1.col <> 5 And MSHFlexGrid1.col <> 7 And MSHFlexGrid1.col <> 8 And MSHFlexGrid1.col <> 10 And MSHFlexGrid1.col <> 16 And MSHFlexGrid1.col <> 15 And MSHFlexGrid1.col <> 14) Then
    Text1.SetFocus
    Text1.Visible = True
    Text1.Width = MSHFlexGrid1.CellWidth - 20
    Text1.Height = MSHFlexGrid1.CellHeight - 20
    Text1.top = MSHFlexGrid1.CellTop + MSHFlexGrid1.top
    Text1.Left = MSHFlexGrid1.CellLeft + MSHFlexGrid1.Left
    Text1.Text = MSHFlexGrid1.Text
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
    Text1.ZOrder
    Text1.SetFocus

End If
End Sub

Private Sub MSHFlexGrid1_LeaveCell()
Text1.Visible = False
End Sub

Private Sub MSHFlexGrid1_LostFocus()
'Text1.Visible = False
End Sub

Private Sub MSHFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'flag = True
End Sub

Private Sub MSHFlexGrid1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'If (flag = True) Then
'If (col <> MSHFlexGrid1.col) Then
'celltotal = celltotal + Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, MSHFlexGrid1.col))
'End If
'col = MSHFlexGrid1.col
'End If
End Sub
Private Sub MSHFlexGrid1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'flag = False
'MsgBox (celltotal)
'celltotal = 0
End Sub

Private Sub MSHFlexGrid1_SelChange()
On Error Resume Next
Dim picturename As String
'MsgBox (MSHFlexGrid1.Row)
'MsgBox (MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 16))
If (MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 16) <> "" And MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 16) <> Empty) Then
     pos = InStrRev(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 16), "\")
     If (pos <> 0) Then
     picturename = Mid(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 16), pos + 1)
     Else
     picturename = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 16)
     End If
     Image1.Visible = True
     Image1.Picture = LoadPicture(MDIForm1.picturepath & picturename)
     'Image1.Picture = LoadPicture(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 17))
Else
    Image1.Picture = LoadPicture()
End If
End Sub

Private Sub print_Click()
result = PrintMSFlexgrid("", vbPRORPortrait, MSHFlexGrid1, 1)

End Sub
Public Function PrintMSFlexgrid(title As String, orientation As Integer, grdName As MSHFlexGrid, line_spaces As Integer)

' Local declares
Dim lRowCount As Long
Dim iColCount As Integer
Dim lRowLoop As Long
Dim iColLoop As Integer
Dim iTabPos As Integer
Dim found_legend As Boolean
Dim print_line As Boolean
Dim top As Integer
Dim outstr As String

Printer.orientation = orientation
' print the title if one has been  entered.

If title <> "" Then
    Printer.FontName = grdName.Font
    Printer.FontName = "Comic Sans MS"
    Printer.FontSize = 14
    Printer.Print title
    Printer.Print ""
    Printer.Print ""
End If
Printer.Print ""
Printer.Print ""
Printer.Print ""
iTabPos = 0
Printer.Print Tab(iTabPos + 1); " Qut No:  " & qut.Text;
iTabPos = iTabPos + (grdName.CellWidth / 60)
Printer.Print Tab(iTabPos + 1); DTPicker1 & "     Party Name : " & party.Text
iTabPos = iTabPos + (grdName.CellWidth / 60)

iTabPos = 0

' Start function
lRowCount = grdName.Rows - 1
iColCount = grdName.Cols - 1
For lRowLoop = 0 To lRowCount
    grdName.row = lRowLoop
    For iColLoop = 0 To iColCount
        If (grdName.ColWidth(iColLoop) <> 0 And iColLoop <> 15) Then
            grdName.col = iColLoop
            ' Grab the flexgrid font properties.

           ' ptr.FontName = "Comic Sans MS" 'grdName.CellFontName
          '  ptr.FontSize = 8 'grdName.CellFontSize
           ' ptr.FontBold = grdName.CellFontBold
           ' ptr.FontItalic = grdName.CellFontItalic
           ' ptr.FontUnderline = grdName.CellFontUnderline
           ' ptr.ForeColor = grdName.CellForeColor
         
            Printer.Print Tab(iTabPos + 1); grdName.Text;
            iTabPos = iTabPos + (grdName.CellWidth / 60)

'           Printer.NewPage
        End If
    Next iColLoop
    Printer.Print ""
    For I = 1 To line_spaces
        Printer.Print ""
    Next I
    iTabPos = 0
Next lRowLoop
Printer.EndDoc
End Function
Private Sub rmv_Click()
On Error Resume Next
Dim num As Integer
Dim flag As Boolean
For j = 0 To I - 1
pos = InStr(1, MSHFlexGrid1.TextMatrix(j, 1), "-", vbTextCompare)
num = Val(Mid(MSHFlexGrid1.TextMatrix(j, 1), pos + 1))
If (num = Val(Code.Text)) Then
flag = True
MSHFlexGrid1.RemoveItem (j)
str1 = "Select * from tblqut where ncode=" & Val(Code.Text)
If (res.state <> 0) Then
res.close
End If
res.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
If (res.EOF = False) Then
res.delete
'load_Click
Exit Sub
End If
End If
Next
If (flag = False) Then
MsgBox ("Item Not Found In the List")
End If
If (res.state <> 0) Then
res.close
End If
End Sub

Private Sub save_Click()
On Error Resume Next
If (party.Text = Empty Or through = Empty) Then
MsgBox ("Party Name And Through Cannot Be Empty")
Exit Sub
End If
res.Open "tblqut", MDIForm1.con1, adOpenDynamic, adLockOptimistic
For j = 1 To I - 1
res.Addnew
res!nappno = Val(qut)
res!ddate = CDate(DTPicker1)
res!chissue = party
res!chthrough = through
pos = InStr(1, MSHFlexGrid1.TextMatrix(j, 1), "-", vbTextCompare)
res!ncode = Val(Mid(MSHFlexGrid1.TextMatrix(j, 1), pos + 1))
pos = InStr(1, MSHFlexGrid1.TextMatrix(j, 1), "-", vbTextCompare)
res!chcategory = Mid(MSHFlexGrid1.TextMatrix(j, 1), 1, pos - 1)
' MsgBox (Mid(MSHFlexgrid1.TextMatrix(j, 1), 1, pos - 1))
res!nweight1 = Val(MSHFlexGrid1.TextMatrix(j, 2))
res!nrate1 = Val(MSHFlexGrid1.TextMatrix(j, 3))
res!chcontent1 = MSHFlexGrid1.TextMatrix(j, 4)
res!nweight2 = Val(MSHFlexGrid1.TextMatrix(j, 5))
res!nrate2 = Val(MSHFlexGrid1.TextMatrix(j, 6))
res!chcontent2 = MSHFlexGrid1.TextMatrix(j, 7)
res!nweight3 = Val(MSHFlexGrid1.TextMatrix(j, 8))
res!nrate3 = Val(MSHFlexGrid1.TextMatrix(j, 9))
res!nweight4 = Val(MSHFlexGrid1.TextMatrix(j, 10))
res!nrate4 = Val(MSHFlexGrid1.TextMatrix(j, 11))
res!nmaking1 = Val(MSHFlexGrid1.TextMatrix(j, 12))
res!npno = Val(MSHFlexGrid1.TextMatrix(j, 14))
res!opicture = MSHFlexGrid1.TextMatrix(j, 16)
res.update
Next j
res.close
save.Enabled = False
load.Enabled = True
update.Enabled = True
delete.Enabled = True
MsgBox ("Record Saved")
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
'MsgBox (KeyCode)
On Error Resume Next
If (KeyCode = 13 Or KeyCode = 39) Then
MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, MSHFlexGrid1.col) = Text1.Text
weight1 = Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 2))
rate1 = Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 3))
weight2 = Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 5))
rate2 = Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 6))
weight3 = Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 8))
rate3 = Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 9))
weight4 = Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 10))
rate4 = Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 11))
making = Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 12))
MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 13) = Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 13)) - Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 13))
tot = Val(weight1) * Val(rate1) + Val(weight2) * Val(rate2) + Val(rate3) * Val(weight3) + Val(rate4) * Val(weight4) + Val(making)
MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 13) = tot
MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 13) = Round(Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 13)) + Val(tot))
'MSHFlexGrid1.TextMatrix(I, 15) = Round(100 - (Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 14)) * 100) / (Round(tot)), 2)

pno = Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 14)) * 100
MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 15) = Round((1 - (tot / pno)) * 100, 2)
Text1.Visible = False
End If
End Sub

Private Sub update_Click()
deleterec
save_Click
End Sub
