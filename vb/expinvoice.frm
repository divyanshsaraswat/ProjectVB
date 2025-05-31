VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form estinv 
   Caption         =   "Invoice"
   ClientHeight    =   6915
   ClientLeft      =   2370
   ClientTop       =   1545
   ClientWidth     =   8835
   LinkTopic       =   "Form8"
   MDIChild        =   -1  'True
   ScaleHeight     =   6915
   ScaleWidth      =   8835
   WindowState     =   2  'Maximized
   Begin VB.CheckBox btcheck 
      Caption         =   "BT"
      Height          =   255
      Left            =   5880
      TabIndex        =   27
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton export 
      Caption         =   "EXPORT"
      Height          =   375
      Left            =   10080
      TabIndex        =   26
      Top             =   7080
      Width           =   1095
   End
   Begin VB.ComboBox qut 
      Height          =   315
      Left            =   840
      TabIndex        =   25
      Top             =   240
      Width           =   1212
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sm.Print"
      Height          =   375
      Left            =   9000
      TabIndex        =   24
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy"
      Height          =   375
      Left            =   6840
      TabIndex        =   23
      Top             =   840
      Width           =   735
   End
   Begin VB.ListBox List2 
      Height          =   510
      Left            =   7800
      Style           =   1  'Checkbox
      TabIndex        =   22
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton Update 
      Caption         =   "&Update"
      Height          =   375
      Left            =   3360
      TabIndex        =   21
      Top             =   7080
      Width           =   972
   End
   Begin VB.CommandButton Print 
      Caption         =   "&Print"
      Height          =   375
      Left            =   7920
      TabIndex        =   20
      Top             =   7080
      Width           =   972
   End
   Begin VB.CommandButton save 
      Caption         =   "&Save"
      Height          =   375
      Left            =   2280
      TabIndex        =   19
      Top             =   7080
      Width           =   972
   End
   Begin VB.CommandButton addnew 
      Caption         =   "A&dd New"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   7080
      Width           =   972
   End
   Begin VB.CommandButton load 
      Caption         =   "&Load"
      Height          =   375
      Left            =   1200
      TabIndex        =   17
      Top             =   7080
      Width           =   972
   End
   Begin VB.CommandButton CLOSE 
      Caption         =   "C&lose"
      Height          =   375
      Left            =   6720
      TabIndex        =   16
      Top             =   7080
      Width           =   1092
   End
   Begin VB.CommandButton clear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   5520
      TabIndex        =   15
      Top             =   7080
      Width           =   1092
   End
   Begin VB.CommandButton delete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   4440
      TabIndex        =   14
      Top             =   7080
      Width           =   972
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
      Left            =   2400
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
   Begin VB.CommandButton rmv 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.ComboBox dis 
      Height          =   315
      ItemData        =   "expinvoice.frx":0000
      Left            =   5040
      List            =   "expinvoice.frx":001C
      TabIndex        =   1
      Text            =   "25%"
      Top             =   840
      Width           =   735
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yy"
      Format          =   59179009
      CurrentDate     =   42461
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4935
      Left            =   0
      TabIndex        =   13
      Top             =   1560
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   8705
      _Version        =   393216
      Cols            =   19
      RowHeightMin    =   300
      FocusRect       =   2
      HighLight       =   2
      AllowUserResizing=   3
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   19
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Image Image1 
      Height          =   1695
      Left            =   9480
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2295
   End
   Begin VB.Line Line3 
      X1              =   240
      X2              =   12000
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   6480
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
      Height          =   252
      Left            =   0
      TabIndex        =   7
      Top             =   240
      Width           =   852
   End
End
Attribute VB_Name = "estinv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim res1 As New ADODB.Recordset
Dim res As New ADODB.Recordset
Dim i As Integer
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
For j = 0 To i - 1
pos = InStr(1, MSHFlexGrid1.TextMatrix(j, 1), "-", vbTextCompare)
num = Val(Mid(MSHFlexGrid1.TextMatrix(j, 1), pos + 1))
If (num = Val(Code.Text)) Then
MsgBox ("Item All Ready Added In The List.")
Exit Sub
End If
Next
str1 = "select Tblcosting.*, Tblcosting.minrate1 * tblcosting.nweight1 as amt1,Tblcosting.minrate2 * tblcosting.nweight2 as amt2,Tblcosting.minrate3 * tblcosting.nweight3 as amt3,"
str1 = str1 & " Tblcosting.minrate4 * tblcosting.nweight4 as amt4,amt1+amt2+amt3+amt4+nmaking1 AS totamt,totamt * 1.333333 as priceno,priceno-((priceno)/100 * " & Val(dis.Text)
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
    
    MSHFlexGrid1.TextMatrix(i, 0) = i
    MSHFlexGrid1.TextMatrix(i, 1) = res1!chcategory & "-" & res1!ncode
    dweight = res1!nweight1
    MSHFlexGrid1.TextMatrix(i, 2) = dweight
'    MSHFlexgrid1.TextMatrix(i, 3) = res1!minrate1
     MSHFlexGrid1.TextMatrix(i, 3) = Round(res1!dprice)
     
    MSHFlexGrid1.TextMatrix(i, 4) = res1!chcontent2
    swt1 = res1!nweight2
    MSHFlexGrid1.TextMatrix(i, 5) = swt1
    MSHFlexGrid1.TextMatrix(i, 6) = res1!minrate2
    MSHFlexGrid1.TextMatrix(i, 7) = res1!chcontent3
    swt2 = res1!nweight3
    MSHFlexGrid1.TextMatrix(i, 8) = swt2
    MSHFlexGrid1.TextMatrix(i, 9) = res1!minrate3
    mwt = res1!nweight4
    MSHFlexGrid1.TextMatrix(i, 10) = Round(mwt, 2)
    MSHFlexGrid1.TextMatrix(i, 11) = res1!minrate4
    making = res1!nmaking1
    MSHFlexGrid1.TextMatrix(i, 12) = making
   
    MSHFlexGrid1.TextMatrix(i, 17) = res1!chpcode
    MSHFlexGrid1.TextMatrix(i, 18) = res1!ndno
    
    If (IsNull(res1!opicture) = False) Then
    
    Dim picturename As String
    pos = InStrRev(res1!opicture, "\")
    
    If (pos <> 0) Then
    picturename = Mid(res1!opicture, pos + 1)
    Else
    picturename = res1!opicture
    End If
     MSHFlexGrid1.TextMatrix(i, 16) = picturename
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
    MSHFlexGrid1.TextMatrix(i, 13) = Round(totamt)
    pno = res1!priceno
    MSHFlexGrid1.TextMatrix(i, 14) = Round(pno / 100)
     MSHFlexGrid1.TextMatrix(i, 15) = Round(100 - (Val(totamt) / (Val(pno) * 100)) * 100, 1)
    gtotal = gtotal + totamt
    dwt = dwt + dweight
    gpno = gpno + Round(pno / 100)
    metal = metal + mwt
    stonewt1 = stonewt1 + swt1
    stonewt2 = stonewt2 + swt2
    
    makingtot = makingtot + making
    res1.MoveNext
    i = i + 1
    MSHFlexGrid1.Rows = i + 1
    Wend

    MSHFlexGrid1.TextMatrix(i, 1) = "Grand Total:"
    MSHFlexGrid1.TextMatrix(i, 2) = Round(dwt, 2)
    MSHFlexGrid1.TextMatrix(i, 3) = Round(diatotrt)
    MSHFlexGrid1.TextMatrix(i, 5) = Round(stonewt1, 2)
    MSHFlexGrid1.TextMatrix(i, 6) = Round(stonetotamt1)
    MSHFlexGrid1.TextMatrix(i, 8) = Round(stonewt2, 2)
    
    MSHFlexGrid1.TextMatrix(i, 9) = Round(stonetotamt2)
    MSHFlexGrid1.TextMatrix(i, 10) = Round(metal, 2)
    MSHFlexGrid1.TextMatrix(i, 11) = Round(metalamt)
    
    MSHFlexGrid1.TextMatrix(i, 12) = Round(makingtot)
    MSHFlexGrid1.TextMatrix(i, 13) = Round(gtotal)
    MSHFlexGrid1.TextMatrix(i, 14) = Round(gpno)
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
loaditem
save.Enabled = False
delete.Enabled = True
load.Enabled = True
update.Enabled = True
End Sub
Private Sub clearing()
DTPicker1 = Format(date, "dd-mmm-yy")
through = ""
i = 1
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
gpno = 0
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
End Sub
Private Sub close_Click()
Unload Me
End Sub

Private Sub code_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
add_Click
End If
End Sub

Private Sub Command1_Click()
'clearing
'This query is changed on the 10-12-2007 bcs of the fields value just changing the query without  optimise so many fiedls are repeated in this query don't confuse.

str1 = " SELECT tblappmaster.chissue, tblappdetail.chcategory, tblappdetail.*,tblappdetail.ncode as ncode ,[Tblcosting].[minrate1]*[Tblcosting].[nweight1] AS amt1, [Tblcosting].[minrate2]*[Tblcosting].[nweight2] AS amt2, [Tblcosting].[minrate3]*[Tblcosting].[nweight3] AS amt3,"
str1 = str1 & "[Tblcosting].[minrate4]*[Tblcosting].[nweight4] AS amt4, (amt1+amt2+amt3+amt4+[Tblcosting].[nmaking1]) AS total, (total * 1.333333)/100 AS priceno, total * 1.333333-(((total * 1.333333) * " & Val(dis.Text) & ")/100) AS ftotamt, tblappmaster.nappno, Tblcosting.nweight1,Tblcosting.chowner,Tblcosting.chpcode,Tblcosting.ndno,Tblcosting.nweight1 as nweight1,(ftotamt-(amt2+amt3+amt4+Tblcosting.nmaking1))/Tblcosting.nweight1 AS dprice , dprice*Tblcosting.nweight1 AS damt"
str1 = str1 & " FROM Tblcosting rIGHT JOIN (tblappmaster INNER JOIN tblappdetail ON tblappmaster.nappno = tblappdetail.nappno) ON Tblcosting.ncode = tblappdetail.ncode "
'str1 = "SELECT tblappmaster.chissue, tblappdetail.chcategory, tblappdetail.*, [tblappdetail].[nrate1]*[tblappdetail].[nweight1]+[tblappdetail].[nrate2]*[tblappdetail].[nweight2]+[tblappdetail].[nrate3]*[tblappdetail].[nweight3]+[tblappdetail].[nrate4]*[tblappdetail].[nweight4]+[tblappdetail].[nmaking1] AS total, (total * 1.333333)/100 AS priceno,priceno-((priceno)/100 * " & Val(dis.Text) & " AS ftotamt,tblappmaster.nappno, Tblcosting.chowner "
'str1 = str1 & ", (ftotamt-(amt2+amt3+amt4+nmaking1))/nweight1 as dprice,dprice*Tblcosting.nweight1 as damt FROM Tblcosting RIGHT JOIN (tblappmaster INNER JOIN tblappdetail ON tblappmaster.nappno = tblappdetail.nappno) ON Tblcosting.ncode = tblappdetail.ncode"
'MsgBox (str1)
'(ftotamt-(amt2+amt3+amt4+Tblcosting.nmaking1))/Tblcosting.nweight1 AS dprice
listdt = ""
listapp = ""
Dim issuein As String
   
For Id = 0 To list2.ListCount - 1
If (list2.Selected(Id) <> False) Then
     pos = InStrRev(list2.List(Id), "(")
     listdt = Mid(list2.List(Id), 1, pos - 1)
     listapp = listapp & listdt & ","
     End If
    Next

If (listapp <> "") Then
     pos = InStrRev(listapp, ",")
     listapp = Mid(listapp, 1, pos - 1)
    If (str2 <> Empty) Then
     str2 = str2 & " And tblappmaster.nappno in(" & listapp & ")"
     Else
     str2 = str2 & " tblappmaster.nappno in(" & listapp & ")"
    End If
End If

If (str2 <> Empty) Then
str1 = str1 & " where " & str2 & " order by chissue,tblappmaster.nappno"
Else
str1 = str1 & " order by chissue,tblappmaster.nappno"
End If
'MsgBox (str1)

res1.Open str1, MDIForm1.con1, adOpenStatic, adLockOptimistic
While (res1.EOF = False)
    
    MSHFlexGrid1.TextMatrix(i, 0) = i
    MSHFlexGrid1.TextMatrix(i, 1) = res1.Fields(1) & "-" & res1!ncode
    dweight = res1!nweight1
    MSHFlexGrid1.TextMatrix(i, 2) = dweight
'    MSHFlexgrid1.TextMatrix(i, 3) = res1!minrate1
     MSHFlexGrid1.TextMatrix(i, 3) = Round(res1!dprice)
     MSHFlexGrid1.TextMatrix(i, 4) = res1!chcontent1
    swt1 = res1!nweight2
    MSHFlexGrid1.TextMatrix(i, 5) = swt1
    MSHFlexGrid1.TextMatrix(i, 6) = res1!nrate2
    MSHFlexGrid1.TextMatrix(i, 7) = res1!chcontent2
    swt2 = res1!nweight3
    MSHFlexGrid1.TextMatrix(i, 8) = swt2
    MSHFlexGrid1.TextMatrix(i, 9) = res1!nrate3
    mwt = res1!nweight4
    MSHFlexGrid1.TextMatrix(i, 10) = mwt
    MSHFlexGrid1.TextMatrix(i, 11) = res1!nrate4
    making = res1!nmaking1
    MSHFlexGrid1.TextMatrix(i, 12) = making
    If (IsNull(res1!chpcode) = False) Then
       MSHFlexGrid1.TextMatrix(i, 17) = res1!chpcode
    Else
       MSHFlexGrid1.TextMatrix(i, 17) = ""
   End If
    If (IsNull(res1!ndno) = False) Then
    MSHFlexGrid1.TextMatrix(i, 18) = res1!ndno
    Else
    MSHFlexGrid1.TextMatrix(i, 18) = 0
    End If
    
    If (IsNull(res1!opicture) = False) Then
    
    Dim picturename As String
    pos = InStrRev(res1!opicture, "\")
    
    If (pos <> 0) Then
    picturename = Mid(res1!opicture, pos + 1)
    Else
    picturename = res1!opicture
    End If
     MSHFlexGrid1.TextMatrix(i, 16) = picturename
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
    MSHFlexGrid1.TextMatrix(i, 13) = Round(totamt)
    pno = res1!priceno
    MSHFlexGrid1.TextMatrix(i, 14) = Round(pno)
     MSHFlexGrid1.TextMatrix(i, 15) = Round(100 - (Val(totamt) / (Val(pno) * 100)) * 100, 1)
    gtotal = gtotal + totamt
    dwt = dwt + dweight
    gpno = gpno + Round(pno / 100)
    metal = metal + mwt
    stonewt1 = stonewt1 + swt1
    stonewt2 = stonewt2 + swt2
    
    makingtot = makingtot + making
    res1.MoveNext
    i = i + 1
    MSHFlexGrid1.Rows = i + 1
    Wend

    MSHFlexGrid1.TextMatrix(i, 1) = "Grand Total:"
    MSHFlexGrid1.TextMatrix(i, 2) = Round(dwt, 2)
    MSHFlexGrid1.TextMatrix(i, 3) = Round(diatotrt)
    MSHFlexGrid1.TextMatrix(i, 5) = Round(stonewt1, 2)
    MSHFlexGrid1.TextMatrix(i, 6) = Round(stonetotamt1)
    MSHFlexGrid1.TextMatrix(i, 8) = Round(stonewt2, 2)
    
    MSHFlexGrid1.TextMatrix(i, 9) = Round(stonetotamt2)
    MSHFlexGrid1.TextMatrix(i, 10) = Round(metal, 2)
    MSHFlexGrid1.TextMatrix(i, 11) = Round(metalamt)
    
    MSHFlexGrid1.TextMatrix(i, 12) = Round(makingtot)
    MSHFlexGrid1.TextMatrix(i, 13) = Round(gtotal)
    MSHFlexGrid1.TextMatrix(i, 14) = Round(gpno)
   res1.close

End Sub

Private Sub Command2_Click()
result = PrintMSHFlexGrid1("", vbPRORPortrait, MSHFlexGrid1, 1)
End Sub

Private Sub Command3_Click()

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
res.Open str1, MDIForm1.con1, adOpenStatic, adLockOptimistic
If (res.EOF = True) Then
MsgBox ("NO RECORD FOUND!")
Exit Sub
End If
'If (res.RecordCount > 0) Then
res.MoveFirst
'End If
While (res.EOF = False)
    res.delete
    res.MoveNext
Wend
   'MsgBox ("Record Deleted")
   res.close
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
objsheet.Cells(3, 2) = "Name"
objsheet.Cells(3, 2).Select
Selection.HorizontalAlignment = xlRight
objsheet.Cells(3, 3).Select
Selection.EntireColumn.AutoFit = True

objsheet.Cells(3, 3) = ""
objsheet.Cells(3, 3) = party.Text

objsheet.Cells(4, 2) = "Through"
objsheet.Cells(4, 2).Select
Selection.HorizontalAlignment = xlRight

objsheet.Cells(4, 3) = ""
objsheet.Cells(4, 3) = through.Text
objsheet.Cells(4, 3).Select
Selection.EntireColumn.AutoFit = True

objsheet.Cells(5, 2) = "Date"
objsheet.Cells(5, 2).Select
Selection.HorizontalAlignment = xlRight

objsheet.Cells(5, 3) = DTPicker1.Value

objsheet.Cells(6, 2) = "Quotation No."
objsheet.Cells(6, 2).Select
Selection.HorizontalAlignment = xlRight
objsheet.Cells(6, 3).Select
Selection.EntireColumn.AutoFit = True
Selection.NumberFormat = "@"
objsheet.Cells(6, 3) = qut.Text & "/" & DTPicker1.Month & "/" & DTPicker1.Year


Dim col
col = ""
Dim i
i = 9

For j = 1 To MSHFlexGrid1.Rows - 2
objsheet.Cells(i, 1) = MSHFlexGrid1.TextMatrix(j, 0)
objsheet.Cells(i, 1).Select
Selection.Font.Bold = True
Selection.HorizontalAlignment = xlCenter
''pos = InStr(1, MSHFlexGrid1.TextMatrix(j, 1), "-", vbTextCompare)
''num = Mid(MSHFlexGrid1.TextMatrix(j, 1), 1, pos - 1)

''ActiveSheet.Pictures.Insert("D:\DAKSH\images\C3OCT30.jpg").Select
''Selection.ShapeRange.ScaleWidth 0.4, msoFalse, msoScaleFromTopLeft

objsheet.Cells(i, 2) = MSHFlexGrid1.TextMatrix(j, 1)
If (MSHFlexGrid1.TextMatrix(j, 4) <> "") Then
col = " " & MSHFlexGrid1.TextMatrix(j, 4) & "-" & MSHFlexGrid1.TextMatrix(j, 5) & " Cts."
End If

If (MSHFlexGrid1.TextMatrix(j, 7) <> "") Then
col = col & " " & MSHFlexGrid1.TextMatrix(j, 7) & "-" & MSHFlexGrid1.TextMatrix(j, 8) & " Cts."
End If


col = col & " " & "Gold Wt." & MSHFlexGrid1.TextMatrix(j, 10) & " Gms."

objsheet.Cells(i + 1, 2) = "Dia. " & MSHFlexGrid1.TextMatrix(j, 2)
objsheet.Cells(i + 2, 2) = "Cts." & col
objsheet.Cells(i + 2, 2).Select
Selection.EntireColumn.AutoFit = True
objsheet.Cells(i + 3, 2) = "Amount :" & MSHFlexGrid1.TextMatrix(j, 13)

objsheet.Cells(i, 5).Select
Dim path
If (MSHFlexGrid1.TextMatrix(j, 16) <> "") Then
path = "D:\DAKSH\images\" & MSHFlexGrid1.TextMatrix(j, 16)
ActiveSheet.Pictures.Insert(path).Select
Selection.ShapeRange.Height = 54#
Selection.ShapeRange.Width = 72#
Else
''Selection.delete
End If

i = i + 5
col = ""
Next j

For i = i To 500
objsheet.Cells(i, 1) = ""
objsheet.Cells(i, 2) = ""
objsheet.Cells(i, 8) = ""
''objsheet.Cells(i, 5).Select
''Selection.delete
Next i


objBook.SaveAs "DataExpo.xls"
objExcel.Workbooks.Open "DatExpo.xls"
objExcel.Visible = True
objBook.PrintPreview
''objBook.CLOSE
' cell(1,1) means cell A1 ;
Set objsheet = Null
Set objBook = Null


End Sub

''Private Sub edatabase_Click()
''str1 = "select * from tblcosting where ncode in (Select ncode from tblqut where nappno=" & Val(qut) & ")"
''res.Open str1, MDIForm1.con1, adOpenStatic, adLockOptimistic
''If (res.EOF = True) Then
''MsgBox ("Please Save The Record First")
''Exit Sub
''End If
''End Sub

Private Sub Form_Activate()
row = 5000
col = 5000
qut.SetFocus
Text1.Visible = False
End Sub

Private Sub Form_Load()
i = 1
DTPicker1.Format = dtpCustom
'apdate = Format(Date, "dd-mmm-yy")
save.Enabled = False
MSHFlexGrid1.TextMatrix(0, 0) = "S.No."
MSHFlexGrid1.ColWidth(0) = 600
MSHFlexGrid1.TextMatrix(0, 1) = "Item Code"
MSHFlexGrid1.ColWidth(1) = 800
MSHFlexGrid1.TextMatrix(0, 2) = "Dia.Wt."
MSHFlexGrid1.ColWidth(2) = 650
MSHFlexGrid1.TextMatrix(0, 3) = "Dia.Rt."
MSHFlexGrid1.ColWidth(3) = 750
MSHFlexGrid1.TextMatrix(0, 4) = "Stone1."
MSHFlexGrid1.ColWidth(4) = 800
MSHFlexGrid1.TextMatrix(0, 5) = "St1 Wt."
MSHFlexGrid1.ColWidth(5) = 650
MSHFlexGrid1.TextMatrix(0, 6) = "St1 Rt."
MSHFlexGrid1.ColWidth(6) = 750
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
MSHFlexGrid1.ColWidth(12) = 850
MSHFlexGrid1.TextMatrix(0, 13) = "Total."
MSHFlexGrid1.ColWidth(13) = 1000
MSHFlexGrid1.TextMatrix(0, 14) = "P.No"
MSHFlexGrid1.ColWidth(14) = 750
MSHFlexGrid1.TextMatrix(0, 15) = "Dis.(%)"
MSHFlexGrid1.ColWidth(15) = 600
MSHFlexGrid1.TextMatrix(0, 16) = ""
MSHFlexGrid1.ColWidth(16) = 0
MSHFlexGrid1.TextMatrix(0, 17) = "Code"
MSHFlexGrid1.ColWidth(17) = 800
MSHFlexGrid1.TextMatrix(0, 18) = "P.No"
MSHFlexGrid1.ColWidth(18) = 650
loaditem
'con1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\lahar\vb\db2.mdb;Persist Security Info=False"
'con1.Open

End Sub
Private Sub loaditem()
str1 = "Select nappno,chissue from tblappmaster order by nappno"
res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
If (list2.ListCount > 0) Then
list2.Clear
End If

While (res1.EOF = False)
 list2.AddItem (res1!nappno & "(" & res1!chissue & ")")
'List2.AddItem (res1!nappo)

'appno.AddItem (res1!nappno)
res1.MoveNext
Wend
res1.close

res1.Open "select distinct(nappno) from tblqut order by nappno", MDIForm1.con1, adOpenDynamic, adLockOptimistic
While (res1.EOF = False)
qut.AddItem (res1!nappno)
res1.MoveNext
Wend
res1.close
End Sub
Sub load_Click()
On Error Resume Next
str1 = "select * from tblqut where nappno=" & Val(qut.Text) & " order by chcategory,ncode"
If (res.state <> 0) Then
res.close
End If
res.Open str1, MDIForm1.con1, adOpenStatic, adLockOptimistic
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
i = 1
While (res.EOF = False)
    MSHFlexGrid1.TextMatrix(i, 0) = i
    MSHFlexGrid1.TextMatrix(i, 1) = res!chcategory & "-" & res!ncode
    If (IsNull(res!nrate1) = False) Then
      MSHFlexGrid1.TextMatrix(i, 3) = res!nrate1
    End If
    If (IsNull(res!nrate2) = False) Then
    MSHFlexGrid1.TextMatrix(i, 6) = res!nrate2
    End If
    If (IsNull(res!nrate3) = False) Then
    MSHFlexGrid1.TextMatrix(i, 9) = res!nrate3
    End If
    If (IsNull(res!nrate4) = False) Then
    MSHFlexGrid1.TextMatrix(i, 11) = res!nrate4
    End If
    dweight = res!nweight1
    MSHFlexGrid1.TextMatrix(i, 2) = dweight
    swt1 = res!nweight2
    MSHFlexGrid1.TextMatrix(i, 5) = swt1
    swt2 = res!nweight3
    MSHFlexGrid1.TextMatrix(i, 8) = swt2
    mwt = res!nweight4
    MSHFlexGrid1.TextMatrix(i, 10) = Round(mwt, 2)
    
    MSHFlexGrid1.TextMatrix(i, 4) = res!chcontent1
    MSHFlexGrid1.TextMatrix(i, 7) = res!chcontent2
    'MSHFlexGrid1.TextMatrix(I, 14) = res!npno
    
    If (IsNull(res!chpcode) = False) Then
    MSHFlexGrid1.TextMatrix(i, 17) = res!chpcode
    End If
    
    If (IsNull(res!ndno) = False) Then
    MSHFlexGrid1.TextMatrix(i, 18) = res!ndno
    End If
    
    If (IsNull(res!opicture) = False) Then
    MSHFlexGrid1.TextMatrix(i, 16) = res!opicture
    End If
   ' If (res!opicture <> "") Then
   ' MSHFlexGrid1.Row = I
   ' MSHFlexGrid1.Col = 14
   ' Set MSHFlexGrid1.CellPicture = LoadPicture(MDIForm1.picturepath & picturename)
   ' End If
    
    If (IsNull(res!nmaking1) = False) Then
    making = res!nmaking1
    End If
    MSHFlexGrid1.TextMatrix(i, 12) = making
    amt1 = Val(MSHFlexGrid1.TextMatrix(i, 3)) * Val(res!nweight1)
    amt2 = Val(MSHFlexGrid1.TextMatrix(i, 6)) * Val(res!nweight2)
    amt3 = Val(MSHFlexGrid1.TextMatrix(i, 9)) * Val(res!nweight3)
    amt4 = Val(MSHFlexGrid1.TextMatrix(i, 11)) * Val(res!nweight4)
    metalamt = metalamt + amt4
    stonetotamt1 = stonetotamt1 + amt2
    stonetotamt2 = stonetotamt2 + amt3
    diatotrt = diatotrt + amt1
    totamt = Val(amt1) + Val(amt2) + Val(amt3) + Val(amt4) + Val(MSHFlexGrid1.TextMatrix(i, 12))
    MSHFlexGrid1.TextMatrix(i, 13) = Round(totamt)
    pno = res!npno
    MSHFlexGrid1.TextMatrix(i, 14) = pno
     MSHFlexGrid1.TextMatrix(i, 15) = Round(100 - (Val(totamt) / (Val(pno) * 100)) * 100, 1)
     gpno = gpno + pno
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
    MSHFlexGrid1.TextMatrix(i, 2) = Round(dwt, 2)
    MSHFlexGrid1.TextMatrix(i, 3) = Round(diatotrt)
    MSHFlexGrid1.TextMatrix(i, 5) = Round(stonewt1, 2)
    MSHFlexGrid1.TextMatrix(i, 6) = Round(stonetotamt1)
    MSHFlexGrid1.TextMatrix(i, 8) = Round(stonewt2, 2)
    MSHFlexGrid1.TextMatrix(i, 9) = Round(stonetotamt2)
    MSHFlexGrid1.TextMatrix(i, 10) = Round(metal, 2)
    MSHFlexGrid1.TextMatrix(i, 11) = Round(metalamt)
    MSHFlexGrid1.TextMatrix(i, 12) = Round(makingtot)
    MSHFlexGrid1.TextMatrix(i, 13) = Round(gtotal)
    MSHFlexGrid1.TextMatrix(i, 14) = Round(gpno)
  
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



Private Sub MSHFlexGrid1_LostFocus()
'Text1.Visible = False
End Sub

Private Sub MSHFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'flag = True
End Sub

Private Sub MSHFlexGrid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If (flag = True) Then
'If (col <> MSHFlexGrid1.col) Then
'celltotal = celltotal + Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, MSHFlexGrid1.col))
'End If
'col = MSHFlexGrid1.col
'End If
End Sub
Private Sub MSHFlexGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
iTabPos = iTabPos + (grdName.CellWidth / 55)
Printer.Print Tab(iTabPos + 1); DTPicker1 & "     Party Name : " & party.Text
iTabPos = iTabPos + (grdName.CellWidth / 55)

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
            iTabPos = iTabPos + (grdName.CellWidth / 65)

'           Printer.NewPage
        End If
    Next iColLoop
    Printer.Print ""
    For i = 1 To line_spaces
        Printer.Print ""
    Next i
    iTabPos = 0
Next lRowLoop
Printer.EndDoc
End Function

Public Function PrintMSHFlexGrid1(title As String, orientation As Integer, grdName As MSHFlexGrid, line_spaces As Integer)

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

If title = "" Then
    Printer.FontName = grdName.Font
    Printer.FontName = "Arial"
    Printer.FontSize = 10
    Printer.Print title
    Printer.Print ""
    Printer.Print ""
End If
Printer.Print ""
Printer.Print ""
Printer.Print ""
iTabPos = 0
Printer.Print Tab(iTabPos + 1); " Qut No:  " & qut.Text;
iTabPos = iTabPos + (10)
Printer.Print Tab(iTabPos + 1); DTPicker1 & "     Party Name : " & party.Text
iTabPos = iTabPos + (10)

iTabPos = 0

' Start function
lRowCount = grdName.Rows - 1
iColCount = grdName.Cols - 1
For lRowLoop = 0 To lRowCount
    grdName.row = lRowLoop
    For iColLoop = 0 To iColCount
        If (grdName.ColWidth(iColLoop) <> 0 And (iColLoop = 1 Or iColLoop = 14 Or iColLoop = 18)) Then
            grdName.col = iColLoop
            ' Grab the flexgrid font properties.

           ' ptr.FontName = "Comic Sans MS" 'grdName.CellFontName
          '  ptr.FontSize = 8 'grdName.CellFontSize
           ' ptr.FontBold = grdName.CellFontBold
           ' ptr.FontItalic = grdName.CellFontItalic
           ' ptr.FontUnderline = grdName.CellFontUnderline
           ' ptr.ForeColor = grdName.CellForeColor
         
            
            Printer.Print Tab(iTabPos + 1); grdName.Text;
            iTabPos = iTabPos + (9)

'           Printer.NewPage
        End If
    Next iColLoop
    Printer.Print ""
    For i = 1 To line_spaces
        Printer.Print ""
    Next i
    iTabPos = 0
Next lRowLoop
Printer.EndDoc
End Function

Private Sub qut_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
load_Click
End If
End Sub

Private Sub rmv_Click()
On Error Resume Next
Dim num As Integer
Dim flag As Boolean
For j = 0 To MSHFlexGrid1.Rows - 1
pos = InStr(1, MSHFlexGrid1.TextMatrix(j, 1), "-", vbTextCompare)
num = Val(Mid(MSHFlexGrid1.TextMatrix(j, 1), pos + 1))
If (num = Val(Code.Text)) Then
flag = True
MSHFlexGrid1.RemoveItem (j)
i = i - 1

str1 = "Select * from tblqut where ncode=" & Val(Code.Text)
If (res.state <> 0) Then
res.close
End If

res.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
If (res.EOF = False) Then
res.delete
res.close
'load_Click
Exit Sub
End If
End If
Next j

If (flag = False) Then
MsgBox ("Item Not Found In the List")
End If
If (res.state <> 0) Then
res.close
End If
End Sub

Private Sub save_Click()
'On Error Resume Next
If (party.Text = Empty Or through = Empty) Then
MsgBox ("Party Name And Through Cannot Be Empty")
Exit Sub
End If
res.Open "tblqut", MDIForm1.con1, adOpenStatic, adLockOptimistic

For j = 1 To MSHFlexGrid1.Rows - 2
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
If (MSHFlexGrid1.TextMatrix(j, 17) = Empty) Then
res!chpcode = ""
Else
res!chpcode = MSHFlexGrid1.TextMatrix(j, 17)
End If
If (MSHFlexGrid1.TextMatrix(j, 18) = Empty) Then
res!ndno = 0
Else
res!ndno = MSHFlexGrid1.TextMatrix(j, 18)
End If
res.update
Next j

res.close

If (btcheck.Value = 1) Then

res.Open "tblcosting", MDIForm1.con3, adOpenDynamic, adLockOptimistic

For j = 1 To i - 1
Dim Code
pos = InStr(1, MSHFlexGrid1.TextMatrix(j, 1), "-", vbTextCompare)
Code = Val(Mid(MSHFlexGrid1.TextMatrix(j, 1), pos + 1))
res1.Open "select * from tblcosting where ncode=" & Val(Code), MDIForm1.con1, adOpenDynamic, adLockOptimistic
res.Addnew

res!ncode = res1!ncode
res!chItemname = res1!chItemname
res!chowner = res1!chowner
res!dinvoicedate = res1!dinvoicedate
res!chcontent1 = res1!chcontent1

res!chcontent2 = res1!chcontent2
res!chcontent3 = res1!chcontent3
res!chcontent4 = res1!chcontent4

res!nweight1 = res1!nweight1
res!nweight2 = res1!nweight2
res!nweight3 = res1!nweight3
res!nweight4 = res1!nweight4
res!nrate1 = 0
res!nrate2 = 0
res!nrate3 = 0
res!nrate4 = 0
res!nigw = res1!nigw
res!nmaking = 0
res!chmaker = res1!chmaker
'res!drecdate = CDate(Invoicedate)
res!drecdate = res1!drecdate
res!chissueto = res1!chissueto
res!ngpw = res1!ngpw
res!pcs1 = res1!pcs1
res!pcs2 = res1!pcs2
res!pcs3 = res1!pcs3
res!gpur = res1!gpur
res!minrate1 = res1!minrate1
res!minrate2 = res1!minrate2
res!minrate3 = res1!minrate3
res!minrate4 = res1!minrate4
res!nmaking1 = res1!nmaking1
res!chquality1 = res1!chquality1
res!chquality2 = res1!chquality2
res!chquality3 = res1!chquality3
res!chcolor1 = res1!chcolor1
res!chcolor2 = res1!chcolor2
res!chcolor3 = res1!chcolor3
res!chsize1 = res1!chsize1
res!chsize2 = res1!chsize2
res!chsize3 = res1!chsize3
res!chstate = res1!chstate
res!chcategory = res1!chcategory
res!chpcode = ""
res!ndno = 0  '"" 'res1!ndno
'If (picturepath <> Empty) Then
    res!opicture = res1!opicture
 '   picturepath = ""
'End If

res.update
res1.close

Next j

res.close
End If

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
'MSHFlexGrid1.SetFocus
End Sub

Private Sub Text1_LostFocus()
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
'MSHFlexGrid1.SetFocus
End If
End Sub

Private Sub update_Click()
If (party.Text = Empty Or through = Empty) Then
MsgBox ("Party Name And Through Cannot Be Empty")
Exit Sub
End If
deleterec
save_Click
End Sub
