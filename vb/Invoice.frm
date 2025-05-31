VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Invoice 
   Caption         =   "Sale Invoice"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   9975
   LinkTopic       =   "Form8"
   MDIChild        =   -1  'True
   ScaleHeight     =   3030
   ScaleWidth      =   9975
   WindowState     =   2  'Maximized
   Begin VB.ComboBox party 
      Height          =   315
      ItemData        =   "Invoice.frx":0000
      Left            =   1440
      List            =   "Invoice.frx":0002
      TabIndex        =   28
      Top             =   840
      Width           =   8175
   End
   Begin VB.TextBox state 
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   2160
      Width           =   4335
   End
   Begin VB.TextBox City 
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Text            =   "Jaipur"
      Top             =   2160
      Width           =   2535
   End
   Begin VB.TextBox address 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Text            =   "address"
      Top             =   1440
      Width           =   8175
   End
   Begin VB.ComboBox dis 
      Height          =   315
      ItemData        =   "Invoice.frx":0004
      Left            =   7320
      List            =   "Invoice.frx":0020
      TabIndex        =   19
      Text            =   "15%"
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton rmv 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   4320
      TabIndex        =   18
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton add 
      Caption         =   "&Add"
      Height          =   375
      Left            =   2880
      TabIndex        =   17
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox code 
      Height          =   315
      Left            =   1680
      TabIndex        =   16
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox through 
      Height          =   315
      Left            =   6360
      TabIndex        =   2
      Top             =   360
      Width           =   3255
   End
   Begin VB.CommandButton delete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   4440
      TabIndex        =   15
      Top             =   9120
      Width           =   972
   End
   Begin VB.CommandButton clear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   5520
      TabIndex        =   14
      Top             =   9120
      Width           =   1092
   End
   Begin VB.CommandButton CLOSE 
      Caption         =   "C&lose"
      Height          =   375
      Left            =   6720
      TabIndex        =   13
      Top             =   9120
      Width           =   1092
   End
   Begin VB.CommandButton load 
      Caption         =   "&Load"
      Height          =   375
      Left            =   1200
      TabIndex        =   12
      Top             =   9120
      Width           =   972
   End
   Begin VB.CommandButton addnew 
      Caption         =   "A&dd New"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   9120
      Width           =   972
   End
   Begin VB.CommandButton save 
      Caption         =   "&Save"
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   9120
      Width           =   972
   End
   Begin VB.CommandButton Print 
      Caption         =   "& Print Invoice"
      Height          =   375
      Left            =   7920
      TabIndex        =   9
      Top             =   9120
      Width           =   1455
   End
   Begin VB.CommandButton Update 
      Caption         =   "&Update"
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   9120
      Width           =   972
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   9600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Det.Print Invoice"
      Height          =   375
      Left            =   9600
      TabIndex        =   6
      Top             =   9120
      Width           =   1335
   End
   Begin VB.ComboBox qut 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   1212
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yy"
      Format          =   59441153
      CurrentDate     =   37684
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4935
      Left            =   120
      TabIndex        =   29
      Top             =   3960
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   8705
      _Version        =   393216
      Cols            =   25
      RowHeightMin    =   300
      FocusRect       =   2
      HighLight       =   2
      AllowUserResizing=   3
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   25
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label7 
      Caption         =   "State"
      Height          =   255
      Left            =   4080
      TabIndex        =   27
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "City"
      Height          =   255
      Left            =   480
      TabIndex        =   26
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Address:"
      Height          =   255
      Left            =   360
      TabIndex        =   25
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Invoice No."
      Height          =   255
      Left            =   360
      TabIndex        =   24
      Top             =   360
      Width           =   855
   End
   Begin VB.Label icode 
      Caption         =   "S. NO.:"
      Height          =   255
      Left            =   360
      TabIndex        =   23
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Date:"
      Height          =   255
      Left            =   3000
      TabIndex        =   22
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Party Name:"
      Height          =   255
      Left            =   360
      TabIndex        =   21
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Through"
      Height          =   255
      Left            =   5400
      TabIndex        =   20
      Top             =   360
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   11760
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   11760
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line3 
      X1              =   240
      X2              =   11880
      Y1              =   9000
      Y2              =   9000
   End
   Begin VB.Image Image1 
      Height          =   2055
      Left            =   9720
      Stretch         =   -1  'True
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "Invoice"
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

''str1 = "SELECT * from tblcosting Where ncode ="
'If (Option1(2).Value = False) Then
''If (I = 22) Then
'    MsgBox ("You Can't Add More Than 21 Items In The One Form")
'Exit Sub
'End If
'End If

If (Val(Code.Text) <= 0) Then
    MsgBox ("Enter the Valid Item No")
Else
str1 = str1 & " and ncode not in (Select ncode from tblappdetail order by ncode) "
res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic

If (res1.EOF = True) Then
    res1.close
    str1 = "Select chissue,nappno from tblappmaster where nappno="
    str1 = str1 & " ( Select  nappno from tblappdetail where ncode=" & Val(Code.Text) & ") "
    res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
    name1 = ""
  If (res1.EOF = False) Then
    result = MsgBox("This Item Found in the Approval No " & res1!nappno & " And Issue to " & res1!chissue & ". Do You Want To add This? ", vbYesNo + vbDefaultButton2, "Confirmation")

    If (Val(result) = 7) Then
        res1.close
        Code.Text = ""
        Code.SetFocus
        Exit Sub
    Else
        str1 = " Select  * from tblappdetail where ncode=" & Val(Code.Text)
        res1.close
        res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
        res1.delete
        res1.close
        add_Click
    End If
  Else
   MsgBox ("Item Not Found ")
    res1.close
    Code.Text = ""
    Code.SetFocus
    Exit Sub
   End If
Else
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
    MSHFlexGrid1.TextMatrix(i, 19) = res1!chowner
    MSHFlexGrid1.TextMatrix(i, 20) = res1!nrate1
    MSHFlexGrid1.TextMatrix(i, 21) = res1!nrate2
    MSHFlexGrid1.TextMatrix(i, 22) = res1!nrate3
    MSHFlexGrid1.TextMatrix(i, 23) = res1!nmaking
    MSHFlexGrid1.TextMatrix(i, 24) = res!nrate4


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
End If
End Sub
Private Sub Addnew_Click()
On Error Resume Next
clearing
res1.Open "select max(nappno) from tblsale ", MDIForm1.con1, adOpenDynamic, adLockOptimistic
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
party = ""
address = ""
city = ""
state = ""
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
str1 = " SELECT tblappmaster.chissue, tblappdetail.chcategory, tblappdetail.*, [tblappdetail].[nrate1]*[tblappdetail].[nweight1] AS amt1, [tblappdetail].[nrate2]*[tblappdetail].[nweight2] AS amt2, [tblappdetail].[nrate3]*[tblappdetail].[nweight3] AS amt3,"
str1 = str1 & "[tblappdetail].[nrate4]*[tblappdetail].[nweight4] AS amt4, amt1+amt2+amt3+amt4+[tblappdetail].[nmaking1] AS total, (total * 1.333333)/100 AS priceno, total * 1.333333-(((total * 1.333333) * " & Val(dis.Text) & ")/100) AS ftotamt, tblappmaster.nappno, Tblcosting.chowner,Tblcosting.chpcode,Tblcosting.ndno,(ftotamt-(amt2+amt3+amt4+tblappdetail.nmaking1))/tblappdetail.nweight1 AS dprice, dprice*Tblcosting.nweight1 AS damt"
str1 = str1 & " FROM Tblcosting RIGHT JOIN (tblappmaster INNER JOIN tblappdetail ON tblappmaster.nappno = tblappdetail.nappno) ON Tblcosting.ncode = tblappdetail.ncode "
'str1 = "SELECT tblappmaster.chissue, tblappdetail.chcategory, tblappdetail.*, [tblappdetail].[nrate1]*[tblappdetail].[nweight1]+[tblappdetail].[nrate2]*[tblappdetail].[nweight2]+[tblappdetail].[nrate3]*[tblappdetail].[nweight3]+[tblappdetail].[nrate4]*[tblappdetail].[nweight4]+[tblappdetail].[nmaking1] AS total, (total * 1.333333)/100 AS priceno,priceno-((priceno)/100 * " & Val(dis.Text) & " AS ftotamt,tblappmaster.nappno, Tblcosting.chowner "
'str1 = str1 & ", (ftotamt-(amt2+amt3+amt4+nmaking1))/nweight1 as dprice,dprice*Tblcosting.nweight1 as damt FROM Tblcosting RIGHT JOIN (tblappmaster INNER JOIN tblappdetail ON tblappmaster.nappno = tblappdetail.nappno) ON Tblcosting.ncode = tblappdetail.ncode"
'MsgBox (str1)
listdt = ""
listapp = ""
Dim issuein As String
   
For Id = 0 To List2.ListCount - 1
If (List2.Selected(Id) <> False) Then
     pos = InStrRev(List2.List(Id), "(")
     listdt = Mid(List2.List(Id), 1, pos - 1)
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

On Error Resume Next

Dim objExcel As Excel.Application, objBook As Excel.Workbook, objsheet As Excel.Worksheet

Set objExcel = GetObject(, "Excel.Application") 'if excel already open you can use GetObject
If Err.Number Then
Err.Clear
End If
Set objExcel = CreateObject("Excel.Application") 'or CreateObject to open new Excel Application
Set objBook = objExcel.Workbooks.Open("c:\\daksh\vb\Book2.xls")
Set objsheet = objBook.Worksheets(1)
''objsheet.Range(B9).Select
objsheet.Cells(3, 3) = ""
objsheet.Cells(3, 3) = party.Text
objsheet.Cells(4, 3) = ""
objsheet.Cells(4, 3) = address.Text & " " & city.Text
objsheet.Cells(5, 3) = ""
''objsheet.Cells(5, 3) = City.Text
objsheet.Cells(4, 8) = DTPicker1.Value
objsheet.Cells(6, 3) = ""
objsheet.Cells(6, 3) = through.Text

objsheet.Cells(5, 8) = qut.Text & "/" & DTPicker1.Month & "/" & DTPicker1.Year
Dim col
col = ""
Dim i
i = 10
For j = 1 To MSHFlexGrid1.Rows - 2
objsheet.Cells(i, 1) = MSHFlexGrid1.TextMatrix(j, 0)

pos = InStr(1, MSHFlexGrid1.TextMatrix(j, 1), "-", vbTextCompare)
num = Mid(MSHFlexGrid1.TextMatrix(j, 1), 1, pos - 1)

If (num = "C") Then
    objsheet.Cells(i, 2) = "Churi" & " (Dia. " & MSHFlexGrid1.TextMatrix(j, 2) & " Cts.@" & MSHFlexGrid1.TextMatrix(j, 3) & ")"
    ElseIf (num = "SL") Then
    objsheet.Cells(i, 2) = "Single Line" & " (Dia. " & MSHFlexGrid1.TextMatrix(j, 2) & " Cts.@" & MSHFlexGrid1.TextMatrix(j, 3) & ")"
    ElseIf (num = "S") Then
    objsheet.Cells(i, 2) = "Neklace & Tops" & " (Dia. " & MSHFlexGrid1.TextMatrix(j, 2) & " Cts.@" & MSHFlexGrid1.TextMatrix(j, 3) & ")"
    ElseIf (num = "P") Then
    objsheet.Cells(i, 2) = "Pendent" & " (Dia. " & MSHFlexGrid1.TextMatrix(j, 2) & " Cts.@" & MSHFlexGrid1.TextMatrix(j, 3) & ")"
    ElseIf (num = "K") Then
    objsheet.Cells(i, 2) = "Bracelet" & " (Dia. " & MSHFlexGrid1.TextMatrix(j, 2) & " Cts.@" & MSHFlexGrid1.TextMatrix(j, 3) & ")"
    ElseIf (num = "TP") Then
    objsheet.Cells(i, 2) = "Tops & Pendt." & " (Dia. " & MSHFlexGrid1.TextMatrix(j, 2) & " Cts.@" & MSHFlexGrid1.TextMatrix(j, 3) & ")"
    ElseIf (num = "O") Then
    objsheet.Cells(i, 2) = "Others" & " (Dia. " & MSHFlexGrid1.TextMatrix(j, 2) & " Cts.@" & MSHFlexGrid1.TextMatrix(j, 3) & ")"
    ElseIf (num = "GR") Then
    objsheet.Cells(i, 2) = "Gents Ring" & " (Dia. " & MSHFlexGrid1.TextMatrix(j, 2) & " Cts.@" & MSHFlexGrid1.TextMatrix(j, 3) & ")"
    ElseIf (num = "LR") Then
    objsheet.Cells(i, 2) = "Ladies Ring" & " (Dia. " & MSHFlexGrid1.TextMatrix(j, 2) & " Cts.@" & MSHFlexGrid1.TextMatrix(j, 3) & ")"
    ElseIf (num = "T") Then
    objsheet.Cells(i, 2) = "Tops" & " (Dia. " & MSHFlexGrid1.TextMatrix(j, 2) & " Cts.@" & MSHFlexGrid1.TextMatrix(j, 3) & ")"
    ElseIf (num = "W") Then
    objsheet.Cells(i, 2) = "Watch" & " (Dia. " & MSHFlexGrid1.TextMatrix(j, 2) & " Cts.@" & MSHFlexGrid1.TextMatrix(j, 3) & ")"
    Else
    objsheet.Cells(i, 2) = "" & " (Dia. " & MSHFlexGrid1.TextMatrix(j, 2) & " Cts.@" & MSHFlexGrid1.TextMatrix(j, 3) & ")"
End If
num = ""



''objsheet.Cells(i, 2) = MSHFlexGrid1.TextMatrix(j, 1)

If (MSHFlexGrid1.TextMatrix(j, 4) <> "") Then
col = " " & MSHFlexGrid1.TextMatrix(j, 4) & "-" & MSHFlexGrid1.TextMatrix(j, 5) & " Cts.@" & MSHFlexGrid1.TextMatrix(j, 6)
End If

If (MSHFlexGrid1.TextMatrix(j, 7) <> "") Then
col = col & " " & MSHFlexGrid1.TextMatrix(j, 7) & "-" & MSHFlexGrid1.TextMatrix(j, 8) & " Cts.@" & MSHFlexGrid1.TextMatrix(j, 9)
End If


col = col & " " & "Gold Wt." & MSHFlexGrid1.TextMatrix(j, 10) & " Gms.@" & MSHFlexGrid1.TextMatrix(j, 11) & " Mak.@" & MSHFlexGrid1.TextMatrix(j, 12)


objsheet.Cells(i + 1, 2) = col
objsheet.Cells(i, 8) = MSHFlexGrid1.TextMatrix(j, 13)
i = i + 2
col = ""
Next j

For i = i To 30
objsheet.Cells(i, 1) = ""
objsheet.Cells(i, 2) = ""
objsheet.Cells(i, 8) = ""
Next i


objBook.save
objExcel.Visible = True
objBook.PrintPreview
''objBook.CLOSE
' cell(1,1) means cell A1 ;
Set objsheet = Null
Set objBook = Null


''result = PrintMSFlexgrid("", vbPRORPortrait, MSHFlexGrid1, 1)
End Sub


''result = PrintMSFlexgrid1("", vbPRORPortrait, MSHFlexGrid1, 1)


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
str1 = "Select * from tblsale where nappno=" & Val(qut)
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

Private Sub edatabase_Click()
str1 = "select * from tblcosting where ncode in (Select ncode from tblqut where nappno=" & Val(qut) & ")"
res.Open str1, MDIForm1.con1, adOpenStatic, adLockOptimistic
If (res.EOF = True) Then
MsgBox ("Please Save The Record First")
Exit Sub
End If
End Sub

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
MSHFlexGrid1.ColWidth(7) = 800
MSHFlexGrid1.TextMatrix(0, 8) = "St2 Wt."
MSHFlexGrid1.ColWidth(8) = 650
MSHFlexGrid1.TextMatrix(0, 9) = "St2 Rt."
MSHFlexGrid1.ColWidth(9) = 750
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
MSHFlexGrid1.TextMatrix(0, 19) = ""
MSHFlexGrid1.ColWidth(19) = 0
MSHFlexGrid1.TextMatrix(0, 20) = "Dia MinRate"
MSHFlexGrid1.ColWidth(20) = 0

MSHFlexGrid1.TextMatrix(0, 21) = "St1. Minrate "
MSHFlexGrid1.ColWidth(21) = 0

MSHFlexGrid1.TextMatrix(0, 22) = "St2. Minrate "
MSHFlexGrid1.ColWidth(22) = 0

MSHFlexGrid1.TextMatrix(0, 23) = "Min Making "
MSHFlexGrid1.ColWidth(23) = 0

MSHFlexGrid1.TextMatrix(0, 24) = "Min Matel Rate "
MSHFlexGrid1.ColWidth(24) = 0

Text1.Visible = False
loaditem
'con1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\lahar\vb\db2.mdb;Persist Security Info=False"
'con1.Open

End Sub
Private Sub loaditem()
'str1 = "Select nappno,chissue from tblappmaster order by nappno"
''res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
''If (List2.ListCount > 0) Then
''List2.clear
''End If

''While (res1.EOF = False)
 ''List2.AddItem (res1!nappno & "(" & res1!chissue & ")")
'List2.AddItem (res1!nappo)

'appno.AddItem (res1!nappno)
''res1.MoveNext
''Wend
''res1.CLOSE
qut.Clear
party.Clear
res1.Open "select distinct(nappno) from tblsale order by nappno", MDIForm1.con1, adOpenDynamic, adLockOptimistic
While (res1.EOF = False)
qut.AddItem (res1!nappno)
res1.MoveNext
Wend
res1.close

res1.Open "select distinct(chissue) from tblsale order by chissue", MDIForm1.con1, adOpenDynamic, adLockOptimistic
While (res1.EOF = False)
party.AddItem (res1!chissue)
res1.MoveNext
Wend
res1.close
End Sub
Sub load_Click()
On Error Resume Next
str1 = "select * from tblsale where nappno=" & Val(qut.Text) & " order by chcategory,ncode"
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
party.Clear
If (IsNull(res!chissue) = False) Then
''party.List(0) = res!chissue
party.Text = res!chissue
Else
party.Text = ""
End If

If (IsNull(res!chthrough)) Then
through = ""
Else
through = res!chthrough
End If

If (IsNull(res!chaddress)) Then
address = ""
Else
address = res!chaddress
End If

If (IsNull(res!chcity)) Then
city = ""
Else
city = res!chcity
End If

If (IsNull(res!chstate)) Then
state = ""
Else
state = res!chstate
End If


'res.CLOSE
'str1 = "select * from tblqut where nappno= " & Val(qut.Text)
'res.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
'Clearing
i = 1
While (res.EOF = False)
    MSHFlexGrid1.TextMatrix(i, 0) = i
    MSHFlexGrid1.TextMatrix(i, 1) = res!chcategory & "-" & res!ncode
    If (IsNull(res!minrate1) = False) Then
      MSHFlexGrid1.TextMatrix(i, 3) = res!minrate1
    End If
    If (IsNull(res!minrate2) = False) Then
    MSHFlexGrid1.TextMatrix(i, 6) = res!minrate2
    End If
    If (IsNull(res!minrate3) = False) Then
    MSHFlexGrid1.TextMatrix(i, 9) = res!minrate3
    End If
    If (IsNull(res!minrate4) = False) Then
    MSHFlexGrid1.TextMatrix(i, 11) = res!minrate4
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
    
    If (IsNull(res!chowner) = False) Then
    MSHFlexGrid1.TextMatrix(i, 19) = res!chowner
    End If
    
    If (IsNull(res!opicture) = False) Then
    MSHFlexGrid1.TextMatrix(i, 16) = res!opicture
    End If
    
    If (IsNull(res!nate1) = False) Then
        MSHFlexGrid1.TextMatrix(i, 20) = res!nrate1
    End If
    If (IsNull(res!nrate2) = False) Then
        MSHFlexGrid1.TextMatrix(i, 21) = res!nrate2
    End If
    If (IsNull(res!nrate3) = False) Then
    MSHFlexGrid1.TextMatrix(i, 22) = res!nrate3
    End If
    If (IsNull(res!nmaking) = False) Then
    MSHFlexGrid1.TextMatrix(i, 23) = res!nmaking
    End If
    If (IsNull(res!nrate4) = False) Then
    MSHFlexGrid1.TextMatrix(i, 24) = res!nrate4
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

On Error Resume Next

Dim objExcel As Excel.Application, objBook As Excel.Workbook, objsheet As Excel.Worksheet

Set objExcel = GetObject(, "Excel.Application") 'if excel already open you can use GetObject
If Err.Number Then
Err.Clear
End If
Set objExcel = CreateObject("Excel.Application") 'or CreateObject to open new Excel Application
Set objBook = objExcel.Workbooks.Open("c:\\daksh\vb\Book1.xls")
Set objsheet = objBook.Worksheets(1)
''objsheet.Range(B9).Select
objsheet.Cells(3, 3) = ""
objsheet.Cells(3, 3) = party.Text
objsheet.Cells(4, 3) = ""
objsheet.Cells(4, 3) = address.Text & " " & city.Text
objsheet.Cells(5, 3) = ""
''objsheet.Cells(5, 3) = City.Text
objsheet.Cells(6, 3) = ""
objsheet.Cells(6, 3) = through.Text
objsheet.Cells(4, 8) = DTPicker1.Value

objsheet.Cells(5, 8) = qut.Text & "/" & DTPicker1.Month & "/" & DTPicker1.Year
Dim col
col = ""
Dim i
i = 10
For j = 1 To MSHFlexGrid1.Rows - 2
objsheet.Cells(i, 1) = MSHFlexGrid1.TextMatrix(j, 0)

pos = InStr(1, MSHFlexGrid1.TextMatrix(j, 1), "-", vbTextCompare)
num = Mid(MSHFlexGrid1.TextMatrix(j, 1), 1, pos - 1)

If (num = "C") Then
    objsheet.Cells(i, 2) = "Churi"
    ElseIf (num = "SL") Then
    objsheet.Cells(i, 2) = "Single Line"
    ElseIf (num = "S") Then
    objsheet.Cells(i, 2) = "Neklace & Tops"
    ElseIf (num = "P") Then
    objsheet.Cells(i, 2) = "Pendent"
    ElseIf (num = "K") Then
    objsheet.Cells(i, 2) = "Bracelet"
    ElseIf (num = "TP") Then
    objsheet.Cells(i, 2) = "Tops & Pendt."
    ElseIf (num = "O") Then
    objsheet.Cells(i, 2) = "Others"
    ElseIf (num = "GR") Then
    objsheet.Cells(i, 2) = "Gents Ring"
    ElseIf (num = "LR") Then
    objsheet.Cells(i, 2) = "Ladies Ring"
    ElseIf (num = "T") Then
    objsheet.Cells(i, 2) = "Tops"
    ElseIf (num = "W") Then
    objsheet.Cells(i, 2) = "Watch"
    Else
    objsheet.Cells(i, 2) = ""
End If
num = ""



''objsheet.Cells(i, 2) = MSHFlexGrid1.TextMatrix(j, 1)
If (MSHFlexGrid1.TextMatrix(j, 4) <> "") Then
col = " " & MSHFlexGrid1.TextMatrix(j, 4) & "-" & MSHFlexGrid1.TextMatrix(j, 5) & " Cts."
End If

If (MSHFlexGrid1.TextMatrix(j, 7) <> "") Then
col = col & " " & MSHFlexGrid1.TextMatrix(j, 7) & "-" & MSHFlexGrid1.TextMatrix(j, 8) & " Cts."
End If


col = col & " " & "Gold Wt." & MSHFlexGrid1.TextMatrix(j, 10) & " Gms."

objsheet.Cells(i + 1, 2) = "Dia. " & MSHFlexGrid1.TextMatrix(j, 2) & " Cts." & col
objsheet.Cells(i, 8) = MSHFlexGrid1.TextMatrix(j, 13)
i = i + 2
col = ""
Next j

For i = i To 30
objsheet.Cells(i, 1) = ""
objsheet.Cells(i, 2) = ""
objsheet.Cells(i, 8) = ""
Next i


objBook.save
objExcel.Visible = True
objBook.PrintPreview
''objBook.CLOSE
' cell(1,1) means cell A1 ;
Set objsheet = Null
Set objBook = Null


''result = PrintMSFlexgrid("", vbPRORPortrait, MSHFlexGrid1, 1)
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
Public Function PrintMSFlexgrid1(title As String, orientation As Integer, grdName As MSHFlexGrid, line_spaces As Integer)
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

str1 = "Select * from tblsale where ncode=" & Val(Code.Text)
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
On Error Resume Next
If (party.Text = Empty Or through = Empty) Then
MsgBox ("Party Name And Through Cannot Be Empty")
Exit Sub
End If
''If (res.state > -1) Then
''res.CLOSE
''End If

res.Open "tblsale", MDIForm1.con1, adOpenStatic, adLockOptimistic

For j = 1 To MSHFlexGrid1.Rows - 2
res.Addnew
res!nappno = Val(qut)
res!ddate = CDate(DTPicker1)
res!chissue = party.Text
res!chthrough = through
res!chaddress = address
res!chcity = city
res!chstate = state
pos = InStr(1, MSHFlexGrid1.TextMatrix(j, 1), "-", vbTextCompare)
res!ncode = Val(Mid(MSHFlexGrid1.TextMatrix(j, 1), pos + 1))
pos = InStr(1, MSHFlexGrid1.TextMatrix(j, 1), "-", vbTextCompare)
res!chcategory = Mid(MSHFlexGrid1.TextMatrix(j, 1), 1, pos - 1)
' MsgBox (Mid(MSHFlexgrid1.TextMatrix(j, 1), 1, pos - 1))
res!nweight1 = Val(MSHFlexGrid1.TextMatrix(j, 2))
res!minrate1 = Val(MSHFlexGrid1.TextMatrix(j, 3))
res!chcontent1 = MSHFlexGrid1.TextMatrix(j, 4)
res!nweight2 = Val(MSHFlexGrid1.TextMatrix(j, 5))
res!minrate2 = Val(MSHFlexGrid1.TextMatrix(j, 6))
res!chcontent2 = MSHFlexGrid1.TextMatrix(j, 7)
res!nweight3 = Val(MSHFlexGrid1.TextMatrix(j, 8))
res!minrate3 = Val(MSHFlexGrid1.TextMatrix(j, 9))
res!nweight4 = Val(MSHFlexGrid1.TextMatrix(j, 10))
res!minrate4 = Val(MSHFlexGrid1.TextMatrix(j, 11))
res!nmaking1 = Val(MSHFlexGrid1.TextMatrix(j, 12))
res!npno = Val(MSHFlexGrid1.TextMatrix(j, 14))
res!opicture = MSHFlexGrid1.TextMatrix(j, 16)

If (MSHFlexGrid1.TextMatrix(j, 17) = Empty) Then
res!chpcode = ""
Else
res!chpcode = MSHFlexGrid1.TextMatrix(j, 17)
End If

If (MSHFlexGrid1.TextMatrix(j, 18) = Empty) Then
res!ndno = ""
Else
res!ndno = MSHFlexGrid1.TextMatrix(j, 18)
End If

res!chowner = MSHFlexGrid1.TextMatrix(j, 19)
res!nrate1 = Val(MSHFlexGrid1.TextMatrix(j, 20))
res!nrate2 = Val(MSHFlexGrid1.TextMatrix(j, 21))
res!nrate3 = Val(MSHFlexGrid1.TextMatrix(j, 22))
res!nmaking = Val(MSHFlexGrid1.TextMatrix(j, 23))
res!nrate4 = Val(MSHFlexGrid1.TextMatrix(j, 24))

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

