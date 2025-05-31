VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Approvel 
   Caption         =   "Delivery chalaan for Approval"
   ClientHeight    =   8340
   ClientLeft      =   1560
   ClientTop       =   1980
   ClientWidth     =   11775
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   8340
   ScaleWidth      =   11775
   WindowState     =   2  'Maximized
   Begin VB.OptionButton Option1 
      Caption         =   "St.App."
      Height          =   195
      Index           =   2
      Left            =   6720
      TabIndex        =   25
      Top             =   1440
      Width           =   975
   End
   Begin VB.ComboBox dis 
      Height          =   315
      ItemData        =   "Approvel.frx":0000
      Left            =   9240
      List            =   "Approvel.frx":001C
      TabIndex        =   24
      Text            =   "25%"
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton rmv 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   3720
      TabIndex        =   23
      Top             =   1440
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "With Image."
      Height          =   195
      Index           =   1
      Left            =   5160
      TabIndex        =   22
      Top             =   1680
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Without Image"
      Height          =   375
      Index           =   0
      Left            =   5160
      TabIndex        =   21
      Top             =   1320
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   120
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   7320
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   4080
      TabIndex        =   19
      Top             =   840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yy"
      Format          =   43384833
      CurrentDate     =   42461
   End
   Begin VB.TextBox through 
      Height          =   435
      Left            =   9600
      TabIndex        =   17
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton delete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   5400
      TabIndex        =   16
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton clear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   6720
      TabIndex        =   15
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton CLOSE 
      Caption         =   "C&lose"
      Height          =   375
      Left            =   8040
      TabIndex        =   14
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton load 
      Caption         =   "&Load"
      Height          =   375
      Left            =   2640
      TabIndex        =   13
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton addnew 
      Caption         =   "A&dd New"
      Height          =   375
      Left            =   1320
      TabIndex        =   12
      Top             =   7320
      Width           =   1215
   End
   Begin VB.TextBox ano 
      Height          =   315
      Left            =   1440
      TabIndex        =   9
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton save 
      Caption         =   "&Save"
      Height          =   375
      Left            =   3960
      TabIndex        =   8
      Top             =   7320
      Width           =   1215
   End
   Begin VB.TextBox issue 
      Height          =   435
      Left            =   6840
      TabIndex        =   6
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton Print 
      Caption         =   "&Print"
      Height          =   375
      Left            =   9480
      TabIndex        =   4
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton add 
      Caption         =   "&Add"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox code 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Top             =   1440
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5055
      Left            =   0
      TabIndex        =   3
      Top             =   2160
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   8916
      _Version        =   393216
      Cols            =   15
      Redraw          =   -1  'True
      ScrollTrack     =   -1  'True
      HighLight       =   2
      AllowUserResizing=   1
      MousePointer    =   1
   End
   Begin VB.Image Image1 
      Height          =   2775
      Left            =   9120
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "Through:"
      Height          =   255
      Left            =   8760
      TabIndex        =   18
      Top             =   840
      Width           =   975
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   11880
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "Daksh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   11
      Top             =   0
      Width           =   8895
   End
   Begin VB.Line Line3 
      X1              =   -120
      X2              =   11760
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   11880
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label3 
      Caption         =   "App. No."
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Issue To/GSTIN"
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Date:"
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   11880
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label icode 
      Caption         =   "S. NO.:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   615
   End
End
Attribute VB_Name = "Approvel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con1 As New ADODB.Connection
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

Private Sub add_Click()
Dim num As Integer
For j = 0 To i - 1
pos = InStr(1, MSFlexGrid1.TextMatrix(j, 1), "-", vbTextCompare)
num = Val(Mid(MSFlexGrid1.TextMatrix(j, 1), pos + 1))
If (num = Val(Code.Text)) Then
MsgBox ("Item All Ready Added In The List.")
Exit Sub
End If
Next
str1 = "select Tblcosting.*, Tblcosting.minrate1 * tblcosting.nweight1 as amt1,Tblcosting.minrate2 * tblcosting.nweight2 as amt2,Tblcosting.minrate3 * tblcosting.nweight3 as amt3,"
str1 = str1 & " Tblcosting.minrate4 * tblcosting.nweight4 as amt4,amt1+amt2+amt3+amt4+nmaking1 AS totamt,totamt * 1.333333 as priceno,priceno-((priceno)/100 * " & Val(dis.Text)
str1 = str1 & ") as ftotamt,(ftotamt-(amt2+amt3+amt4+nmaking1))/nweight1 as dprice,dprice*Tblcosting.nweight1 as damt "
str1 = str1 & " from tblcosting  where ncode= "

'str1 = "SELECT * from tblcosting Where ncode ="
If (Option1(2).Value = False) Then
If (i = 22) Then
    MsgBox ("You Can't Add More Than 21 Items In The One Form")
Exit Sub
End If
End If

If (Val(Code.Text) <= 0) Then
    MsgBox ("Enter the Valid S.No")
Else
str1 = str1 & Val(Code.Text) & " and chstate='Avi.' and ncode not in (Select ncode from tblappdetail order by ncode) "
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
    Exit Sub
    
  End If
Else
Code.Text = ""
Code.SetFocus
   
While (res1.EOF = False)
    
    MSFlexGrid1.TextMatrix(i, 0) = i
    MSFlexGrid1.TextMatrix(i, 1) = res1!chcategory & "-" & res1!ncode
    dweight = res1!nweight1
    MSFlexGrid1.TextMatrix(i, 2) = dweight
'    MSFlexGrid1.TextMatrix(i, 3) = res1!minrate1
     MSFlexGrid1.TextMatrix(i, 3) = Round(res1!dprice)
     
    MSFlexGrid1.TextMatrix(i, 4) = res1!chcontent2
    swt1 = res1!nweight2
    MSFlexGrid1.TextMatrix(i, 5) = swt1
    MSFlexGrid1.TextMatrix(i, 6) = res1!minrate2
    MSFlexGrid1.TextMatrix(i, 7) = res1!chcontent3
    swt2 = res1!nweight3
    MSFlexGrid1.TextMatrix(i, 8) = swt2
    MSFlexGrid1.TextMatrix(i, 9) = res1!minrate3
    mwt = res1!nweight4
    MSFlexGrid1.TextMatrix(i, 10) = mwt
    MSFlexGrid1.TextMatrix(i, 11) = res1!minrate4
    making = res1!nmaking1
    MSFlexGrid1.TextMatrix(i, 12) = making
    If (IsNull(res1!opicture) = False) Then
    Dim picturename As String
    pos = InStrRev(res1!opicture, "\")
    If (pos <> 0) Then
    picturename = Mid(res1!opicture, pos + 1)
    Else
    picturename = res1!opicture
    End If
     MSFlexGrid1.TextMatrix(i, 14) = picturename
    End If
    'amt1 = Val(res1!minrate1) * Val(res1!nweight1)
    'amt2 = Val(res1!minrate2) * Val(res1!nweight2)
    'amt3 = Val(res1!minrate3) * Val(res1!nweight3)
    'amt4 = Val(res1!minrate4) * Val(res1!nweight4)
    metalamt = metalamt + res1!amt4
    stonetotamt1 = stonetotamt1 + res1!amt2
    stonetotamt2 = stonetotamt2 + res1!amt3
    diatotrt = diatotrt + res1!damt
    totamt = Val(res1!damt) + Val(res1!amt2) + Val(res1!amt3) + Val(res1!amt4) + Val(res1!nmaking1)
    MSFlexGrid1.TextMatrix(i, 13) = Round(totamt)
    
    gtotal = gtotal + totamt
    dwt = dwt + dweight
    metal = metal + mwt
    stonewt1 = stonewt1 + swt1
    stonewt2 = stonewt2 + swt2
    
    makingtot = makingtot + making
    res1.MoveNext
    i = i + 1
    MSFlexGrid1.Rows = i + 1
    Wend

    MSFlexGrid1.TextMatrix(i, 1) = "Grand Total:"
    MSFlexGrid1.TextMatrix(i, 2) = Round(dwt, 2)
    MSFlexGrid1.TextMatrix(i, 3) = Round(diatotrt)
    MSFlexGrid1.TextMatrix(i, 5) = Round(stonewt1, 2)
    MSFlexGrid1.TextMatrix(i, 6) = Round(stonetotamt1)
    MSFlexGrid1.TextMatrix(i, 8) = Round(stonewt2, 2)
    
    MSFlexGrid1.TextMatrix(i, 9) = Round(stonetotamt2)
    MSFlexGrid1.TextMatrix(i, 10) = Round(metal, 2)
    MSFlexGrid1.TextMatrix(i, 11) = Round(metalamt)
    
    MSFlexGrid1.TextMatrix(i, 12) = Round(makingtot)
    MSFlexGrid1.TextMatrix(i, 13) = Round(gtotal)
   res1.close
End If
End If

End Sub
Private Sub Addnew_Click()
clearing
res1.Open "select max(nappno) from tblappmaster ", MDIForm1.con1, adOpenDynamic, adLockOptimistic
If (IsNull(res1.Fields(0))) Then
ano = 1
Else
ano = res1.Fields(0) + 1
End If
res1.close
DTPicker1 = Format(date, "dd-mmm-yy")
'apdate = Format(Date, "dd-mmm-yy")
save.Enabled = True
End Sub



Private Sub ano_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
load_Click
End If
End Sub

Private Sub Clear_Click()
clearing
save.Enabled = False
delete.Enabled = True
End Sub

Private Sub close_Click()
'con1.close
Unload Me
End Sub


Private Sub code_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
add_Click
End If
End Sub

Private Sub Command1_Click()
'parserText ("adsdsd'fsfd'sdfdfsd'' fdsfs")
End Sub

Private Sub delete_Click()
str1 = "Select * from tblappmaster where nappno=" & Val(ano)
res.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
If (res.EOF = True) Then
MsgBox ("NO RECORD FOUND!")
Exit Sub
End If

res.delete
res.close

str1 = "Select * from tblappdetail where nappno=" & Val(ano)
'MsgBox (str1)
res.Open str1, MDIForm1.con1, adOpenStatic, adLockOptimistic
'MsgBox (res.RecordCount)
If (res.RecordCount > 0) Then
res.MoveFirst
End If
Do While Not res.EOF
    res.delete
    res.MoveNext
Loop
   MsgBox ("Record Deleted")
clearing
delete.Enabled = False
res.close
End Sub

Private Sub Form_Activate()
ano.SetFocus
Text1.Visible = False
End Sub

Private Sub Form_Load()
i = 1
DTPicker1.Format = dtpCustom
'apdate = Format(Date, "dd-mmm-yy")
save.Enabled = False
MSFlexGrid1.TextMatrix(0, 0) = "S.No."
MSFlexGrid1.ColWidth(0) = 650
MSFlexGrid1.TextMatrix(0, 1) = "Item Code"
MSFlexGrid1.ColWidth(1) = 850
MSFlexGrid1.TextMatrix(0, 2) = "Dia.Wt."
MSFlexGrid1.ColWidth(2) = 650
MSFlexGrid1.TextMatrix(0, 3) = "Dia.Rt."
MSFlexGrid1.ColWidth(3) = 850
MSFlexGrid1.TextMatrix(0, 4) = "Stone1."
MSFlexGrid1.ColWidth(4) = 900
MSFlexGrid1.TextMatrix(0, 5) = "St1 Wt."
MSFlexGrid1.ColWidth(5) = 650
MSFlexGrid1.TextMatrix(0, 6) = "St1 Rt."
MSFlexGrid1.ColWidth(6) = 850
MSFlexGrid1.TextMatrix(0, 7) = "Stone2"
MSFlexGrid1.ColWidth(7) = 900
MSFlexGrid1.TextMatrix(0, 8) = "St2 Wt."
MSFlexGrid1.ColWidth(8) = 650
MSFlexGrid1.TextMatrix(0, 9) = "St2 Rt."
MSFlexGrid1.ColWidth(9) = 850
MSFlexGrid1.TextMatrix(0, 10) = "MetalWt."
MSFlexGrid1.ColWidth(10) = 700
MSFlexGrid1.TextMatrix(0, 11) = "MetalRt."
MSFlexGrid1.ColWidth(11) = 650
MSFlexGrid1.TextMatrix(0, 12) = "Making."
MSFlexGrid1.ColWidth(12) = 900
MSFlexGrid1.TextMatrix(0, 13) = "Total."
MSFlexGrid1.ColWidth(13) = 1200
MSFlexGrid1.TextMatrix(0, 14) = ""
MSFlexGrid1.ColWidth(14) = 50
'con1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\lahar\vb\db2.mdb;Persist Security Info=False"
'con1.Open

End Sub
Private Sub clearing()
'apdate = Format(Date, "dd-mmm-yy")
'ano = 0
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
issue = ""
Code = ""
    
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
'con1.CLOSE
End Sub

Private Sub load_Click()
str1 = "select * from tblappmaster where nappno=" & Val(ano.Text)
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
DTPicker1 = Format(res!dappdate, "dd-mmm-yy")
issue = res!chissue
If (IsNull(res!chthrough)) Then
through = ""
Else
through = res!chthrough
End If
res.close
str1 = "select A.* from tblappdetail As A Where A.nappno = " & Val(ano.Text)
res.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
'Clearing
i = 1
While (res.EOF = False)
    MSFlexGrid1.TextMatrix(i, 0) = i
    MSFlexGrid1.TextMatrix(i, 1) = res!chcategory & "-" & res!ncode
    If (IsNull(res!nrate1) = False) Then
      MSFlexGrid1.TextMatrix(i, 3) = res!nrate1
    End If
    If (IsNull(res!nrate2) = False) Then
    MSFlexGrid1.TextMatrix(i, 6) = res!nrate2
    End If
    If (IsNull(res!nrate3) = False) Then
    MSFlexGrid1.TextMatrix(i, 9) = res!nrate3
    End If
    If (IsNull(res!nrate4) = False) Then
    MSFlexGrid1.TextMatrix(i, 11) = res!nrate4
    End If
    dweight = res!nweight1
    MSFlexGrid1.TextMatrix(i, 2) = dweight
    swt1 = res!nweight2
    MSFlexGrid1.TextMatrix(i, 5) = swt1
    swt2 = res!nweight3
    MSFlexGrid1.TextMatrix(i, 8) = swt2
    mwt = res!nweight4
    MSFlexGrid1.TextMatrix(i, 10) = mwt
    
    MSFlexGrid1.TextMatrix(i, 4) = res!chcontent1
    MSFlexGrid1.TextMatrix(i, 7) = res!chcontent2
    If (IsNull(res!opicture) = False) Then
    MSFlexGrid1.TextMatrix(i, 14) = res!opicture
    End If
    If (IsNull(res!nmaking1) = False) Then
    making = res!nmaking1
    End If
    MSFlexGrid1.TextMatrix(i, 12) = making
    amt1 = Val(MSFlexGrid1.TextMatrix(i, 3)) * Val(res!nweight1)
    amt2 = Val(MSFlexGrid1.TextMatrix(i, 6)) * Val(res!nweight2)
    amt3 = Val(MSFlexGrid1.TextMatrix(i, 9)) * Val(res!nweight3)
    amt4 = Val(MSFlexGrid1.TextMatrix(i, 11)) * Val(res!nweight4)
    metalamt = metalamt + amt4
    stonetotamt1 = stonetotamt1 + amt2
    stonetotamt2 = stonetotamt2 + amt3
    diatotrt = diatotrt + amt1
    totamt = Val(amt1) + Val(amt2) + Val(amt3) + Val(amt4) + Val(MSFlexGrid1.TextMatrix(i, 12))
    MSFlexGrid1.TextMatrix(i, 13) = totamt
    gtotal = gtotal + totamt
    dwt = dwt + dweight
    metal = metal + mwt
    stonewt1 = stonewt1 + swt1
    stonewt2 = stonewt2 + swt2
    makingtot = makingtot + making
    res.MoveNext
    i = i + 1
    MSFlexGrid1.Rows = i + 1
    Wend
    MSFlexGrid1.TextMatrix(i, 1) = "Grand Total:"
    MSFlexGrid1.TextMatrix(i, 2) = Round(dwt, 2)
    MSFlexGrid1.TextMatrix(i, 3) = Round(diatotrt)
    MSFlexGrid1.TextMatrix(i, 5) = Round(stonewt1, 2)
    MSFlexGrid1.TextMatrix(i, 6) = Round(stonetotamt1)
    MSFlexGrid1.TextMatrix(i, 8) = Round(stonewt2, 2)
    
    MSFlexGrid1.TextMatrix(i, 9) = Round(stonetotamt2)
    MSFlexGrid1.TextMatrix(i, 10) = Round(metal, 2)
    MSFlexGrid1.TextMatrix(i, 11) = Round(metalamt)
    
    MSFlexGrid1.TextMatrix(i, 12) = Round(makingtot)
    MSFlexGrid1.TextMatrix(i, 13) = Round(gtotal)
  
  res.close
End Sub

Private Sub MSFlexGrid1_DblClick()
On Error Resume Next
If (MSFlexGrid1.col <> 1 And MSFlexGrid1.col <> 2 And MSFlexGrid1.col <> 4 And MSFlexGrid1.col <> 5 And MSFlexGrid1.col <> 7 And MSFlexGrid1.col <> 8 And MSFlexGrid1.col <> 10 And MSFlexGrid1.col <> 13) Then
If (Option1(1).Value = True) Then
    Text1.Visible = True
    Text1.Width = MSFlexGrid1.CellWidth
    Text1.Height = MSFlexGrid1.CellHeight
    Text1.top = MSFlexGrid1.CellTop + MSFlexGrid1.top
    Text1.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
    Text1.Text = MSFlexGrid1.Text
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
    Text1.ZOrder
    Text1.SetFocus
End If
End If
End Sub



Private Sub MSFlexGrid1_SelChange()
On Error Resume Next
If (MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 14) <> "" And MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 14) <> Empty) Then
     Dim picturename As String
     pos = InStrRev(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 14), "\")
     If (pos <> 0) Then
     picturename = Mid(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 14), pos + 1)
     Else
     picturename = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 14)
     End If
     Image1.Visible = True
     Image1.Picture = LoadPicture(MDIForm1.picturepath & picturename)
     'Image1.Picture = LoadPicture(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 17))
Else
    Image1.Picture = LoadPicture()
End If
End Sub


Private Sub print_Click()
If (Option1(0).Value = True) Then

If DataEnvironment2.rsCommand3_Grouping.state = adStateOpen Then
DataEnvironment2.rsCommand3_Grouping.close
End If
DataEnvironment2.Commands(3).Parameters(0).Value = Val(ano.Text)
DataEnvironment2.Commands(3).Parameters(1).Value = Val(ano.Text)

If (DataEnvironment2.Commands(3).Execute.EOF = True) Then
   MsgBox ("Please Save The Record First Then Print The Approval")
Else
str1 = " SELECT sum(((B.nweight1 * B.minrate1 + B.nweight2 * B.minrate2 + B.nweight3 * B.minrate3 + B.nweight4 * B.minrate4) + B.nmaking1) * 00.01333333) AS total,"
str1 = str1 & " sum(A.nweight1) AS nwgt, sum(A.nweight2+A.nweight3) AS stwt, sum(((B.nweight1 * B.minrate1 + B.nweight2 * B.minrate2 + B.nweight3 * B.minrate3 + B.nweight4 * B.minrate4) + B.nmaking1) * 00.01333333) *100/68 AS usd "
str1 = str1 & " FROM tblcosting AS B, tblappdetail AS A Where A.nappno = " & ano.Text & " And B.ncode = A.ncode GROUP BY A.nappno"
res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
Debug.Print (str1)
DataReport2.Sections("Section3").Controls("label29").Caption = "" & res1.Fields("nwgt")
DataReport2.Sections("Section3").Controls("label31").Caption = "" & res1.Fields("stwt")
DataReport2.Sections("Section3").Controls("label3").Caption = "" & Round(res1.Fields("total"))
'DataReport2.Sections("Section3").Controls("label30").Caption = "" & Round(res1.Fields("usd"))
res1.close
DataReport2.Show
'DataReport2.PrintReport True
End If
End If
'Else
'If DataEnvironment2.rsCommand9_Grouping.State = adStateOpen Then
'DataEnvironment2.rsCommand9_Grouping.CLOSE
'End If
'DataEnvironment2.Commands(9).Parameters(0).Value = Val(ano.Text)
'DataEnvironment2.Commands(9).Parameters(1).Value = Val(ano.Text)
'
'If (DataEnvironment2.Commands(9).Execute.EOF = True) Then
'   MsgBox ("Please Save The Record First Then Print The Approval")
'Else
'With DataReport7.Sections("Section1")
''DataEnvironment2.rsCommand10.Open
''If (IsNull(DataEnvironment2.rsCommand10.Fields("opicture")) = False) Then
' '  .Controls("image1").Picture = LoadPicture(DataEnvironment2.rsCommand10.Fields("opicture"))
''End If
'End With
'
'DataReport7.Show
'End If
'End If
End Sub
Private Sub rmv_Click()
Dim num As Integer
Dim flag As Boolean
For j = 0 To i - 1
pos = InStr(1, MSFlexGrid1.TextMatrix(j, 1), "-", vbTextCompare)
num = Val(Mid(MSFlexGrid1.TextMatrix(j, 1), pos + 1))
If (num = Val(Code.Text)) Then
flag = True
str1 = "Select * from tblappdetail where ncode=" & Val(Code.Text)
If (res.state <> 0) Then
res.close
End If
res.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
If (res.EOF = False) Then
res.delete
load_Click
Exit Sub
End If

MSFlexGrid1.RemoveItem (j)
'MSFlexGrid1.Refresh
i = i - 1

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
If (issue.Text = Empty Or through = Empty) Then
MsgBox ("'Issue To' And 'Through' Cannot Be Empty")
Exit Sub
End If

res.Open "tblappmaster", MDIForm1.con1, adOpenDynamic, adLockOptimistic
res.Addnew
res!nappno = Val(ano)
res!dappdate = CDate(DTPicker1)
'res!chissue = parserText(issue)
res!chissue = issue
'MsgBox (parserText(issue))
res!chthrough = through
res.update
res.close

res.Open "tblappdetail", MDIForm1.con1, adOpenDynamic, adLockOptimistic

For j = 1 To i - 1
    res.Addnew
    res!nappno = Val(ano.Text)
    pos = InStr(1, MSFlexGrid1.TextMatrix(j, 1), "-", vbTextCompare)
    res!ncode = Val(Mid(MSFlexGrid1.TextMatrix(j, 1), pos + 1))
    
    pos = InStr(1, MSFlexGrid1.TextMatrix(j, 1), "-", vbTextCompare)
    res!chcategory = Mid(MSFlexGrid1.TextMatrix(j, 1), 1, pos - 1)
   ' MsgBox (Mid(MSFlexGrid1.TextMatrix(j, 1), 1, pos - 1))
    res!nweight1 = Val(MSFlexGrid1.TextMatrix(j, 2))
    res!nrate1 = Val(MSFlexGrid1.TextMatrix(j, 3))
    res!chcontent1 = MSFlexGrid1.TextMatrix(j, 4)
    res!nweight2 = Val(MSFlexGrid1.TextMatrix(j, 5))
    res!nrate2 = Val(MSFlexGrid1.TextMatrix(j, 6))
    res!chcontent2 = MSFlexGrid1.TextMatrix(j, 7)
    res!nweight3 = Val(MSFlexGrid1.TextMatrix(j, 8))
    res!nrate3 = Val(MSFlexGrid1.TextMatrix(j, 9))
    res!nweight4 = Val(MSFlexGrid1.TextMatrix(j, 10))
    res!nrate4 = Val(MSFlexGrid1.TextMatrix(j, 11))
    res!nmaking1 = Val(MSFlexGrid1.TextMatrix(j, 12))
    res!opicture = MSFlexGrid1.TextMatrix(j, 14)
    res.update
Next j
res.close
save.Enabled = False
MsgBox ("Record Saved")
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = 13) Then
MSFlexGrid1.TextMatrix(MSFlexGrid1.row, MSFlexGrid1.col) = Text1.Text
weight1 = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 2))
rate1 = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 3))
weight2 = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 5))
rate2 = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 6))
weight3 = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 8))
rate3 = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 9))
weight4 = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 10))
rate4 = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 11))
making = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 12))
MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 13) = Val(weight1) * Val(rate1) + Val(weight2) * Val(rate2) + Val(rate3) * Val(weight3) + Val(rate4) * Val(weight4) + Val(making)
Text1.Visible = False
End If
End Sub

