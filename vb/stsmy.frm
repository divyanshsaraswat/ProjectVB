VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form stsmy 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Stock Summary"
   ClientHeight    =   6585
   ClientLeft      =   1995
   ClientTop       =   1875
   ClientWidth     =   7560
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form8"
   MDIChild        =   -1  'True
   ScaleHeight     =   6585
   ScaleWidth      =   7560
   WindowState     =   2  'Maximized
   Begin VB.ListBox sowner 
      Height          =   510
      ItemData        =   "stsmy.frx":0000
      Left            =   1200
      List            =   "stsmy.frx":0013
      Style           =   1  'Checkbox
      TabIndex        =   8
      Top             =   5160
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3855
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   6800
      _Version        =   393216
      Rows            =   14
      Cols            =   12
      _NumberOfBands  =   1
      _Band(0).Cols   =   12
   End
   Begin VB.CommandButton print1 
      Caption         =   "&Print "
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton close1 
      Caption         =   "&Close"
      Height          =   375
      Left            =   7080
      TabIndex        =   2
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton Search 
      Caption         =   "&Search"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton Clear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label date 
      BackColor       =   &H8000000E&
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Date:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Owner Type.:"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   5160
      Width           =   975
   End
End
Attribute VB_Name = "stsmy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim con1 As New ADODB.Connection
Dim res1 As New ADODB.Recordset
Private Sub Clear_Click()
For i = 1 To 12
    MSHFlexGrid1.TextMatrix(i, 1) = ""
    MSHFlexGrid1.TextMatrix(i, 2) = ""
    MSHFlexGrid1.TextMatrix(i, 3) = ""
    MSHFlexGrid1.TextMatrix(i, 4) = ""
    MSHFlexGrid1.TextMatrix(i, 5) = ""
    MSHFlexGrid1.TextMatrix(i, 6) = ""
    MSHFlexGrid1.TextMatrix(i, 7) = ""
    MSHFlexGrid1.TextMatrix(i, 8) = ""
    MSHFlexGrid1.TextMatrix(i, 9) = ""
    MSHFlexGrid1.TextMatrix(i, 10) = ""
    MSHFlexGrid1.TextMatrix(i, 11) = ""
    MSHFlexGrid1.TextMatrix(i, 12) = ""
  '  MSHFlexGrid1.TextMatrix(i, 13) = ""
   Next i
End Sub

Private Sub close_Click()
Unload Me
End Sub


Private Sub close1_Click()
Unload Me
End Sub

Private Sub Form_Load()
date.Caption = Now()
MSHFlexGrid1.TextMatrix(0, 0) = "Item"
MSHFlexGrid1.ColWidth(0) = 950
MSHFlexGrid1.TextMatrix(0, 1) = "0-25"
MSHFlexGrid1.ColWidth(1) = 700
MSHFlexGrid1.TextMatrix(0, 2) = "25-50"
MSHFlexGrid1.ColWidth(2) = 900
MSHFlexGrid1.TextMatrix(0, 3) = "50-100"
MSHFlexGrid1.ColWidth(3) = 900
MSHFlexGrid1.TextMatrix(0, 4) = "100-300"
MSHFlexGrid1.ColWidth(4) = 900
MSHFlexGrid1.TextMatrix(0, 5) = "300-500"
MSHFlexGrid1.ColWidth(5) = 900
MSHFlexGrid1.TextMatrix(0, 6) = "500-1000"
MSHFlexGrid1.ColWidth(6) = 900
MSHFlexGrid1.TextMatrix(0, 7) = "1000-1500"
MSHFlexGrid1.ColWidth(7) = 900
MSHFlexGrid1.TextMatrix(0, 8) = "1500-2000"
MSHFlexGrid1.ColWidth(8) = 900
MSHFlexGrid1.TextMatrix(0, 9) = ">2000"
MSHFlexGrid1.ColWidth(9) = 900
MSHFlexGrid1.TextMatrix(0, 10) = "I.Total"
MSHFlexGrid1.ColWidth(10) = 900
MSHFlexGrid1.TextMatrix(0, 11) = "Tot.Amt."
MSHFlexGrid1.ColWidth(11) = 950
MSHFlexGrid1.TextMatrix(1, 0) = "Churi"
MSHFlexGrid1.TextMatrix(2, 0) = "GRing"
MSHFlexGrid1.TextMatrix(3, 0) = "Kara"
MSHFlexGrid1.TextMatrix(4, 0) = "LRing"
MSHFlexGrid1.TextMatrix(5, 0) = "Others"
MSHFlexGrid1.TextMatrix(6, 0) = "Pendent"
MSHFlexGrid1.TextMatrix(7, 0) = "Set"
MSHFlexGrid1.TextMatrix(8, 0) = "S.Line"
MSHFlexGrid1.TextMatrix(9, 0) = "Tops"
MSHFlexGrid1.TextMatrix(10, 0) = "TP"
MSHFlexGrid1.TextMatrix(11, 0) = "LS"
MSHFlexGrid1.TextMatrix(12, 0) = "W"
MSHFlexGrid1.TextMatrix(13, 0) = "Total"
End Sub


Private Sub print1_Click()
Search.Visible = False
Clear.Visible = False
close1.Visible = False
sowner.Visible = False
Label1.Visible = False
print1.Visible = False
stsmy.PrintForm
print1.Visible = True
Label1.Visible = True
Search.Visible = True
Clear.Visible = True
close1.Visible = True
sowner.Visible = True
End Sub

Private Sub Search_Click()
On Error Resume Next
Clear_Click
Dim j As Double
Dim k As Double
Dim lstotal, lstotamt, stotal, ctotal, ctotamt, grtotal, grtotamt, lrtotal, ktotal, ktotamt, sltotal, sltotamt, ototal, ototamt, ptotal, pttotamt, ttotal, ttotamt, tptotal, tptotamt, wtotal, wtotalamt, amttotal As Double
Dim listapp As String
listapp = ""
For Id = 0 To sowner.ListCount - 1
If (sowner.Selected(Id) = True) Then
       listapp = listapp & "'" & sowner.List(Id) & "',"
 End If
Next

    
For i = 1 To 9
If (i = 1) Then
j = 0
k = 25000
ElseIf (i = 2) Then
j = k
k = 50000
ElseIf (i = 3) Then
j = k
k = 100000
ElseIf (i = 4) Then
j = k
k = 300000
ElseIf (i = 5) Then
j = k
k = 500000
ElseIf (i = 6) Then
j = k
k = 1000000
ElseIf (i = 7) Then
j = k
k = 1500000
ElseIf (i = 8) Then
j = k
k = 2000000
ElseIf (i = 9) Then
j = k
k = 25000000
End If

str1 = " SELECT count(ncode) as no1 , sum(tblcosting.nweight1 * tblcosting.minrate1 + "
str1 = str1 & " tblcosting.nweight2 * tblcosting.minrate2+ tblcosting.nweight3 * tblcosting.minrate3 + "
str1 = str1 & " + tblcosting.nweight4 * tblcosting.minrate4 +nmaking1) "
str1 = str1 & " as amt, chcategory From tblcosting WHERE chstate=" & "'" & "Avi." & "'"
str1 = str1 & " and (tblcosting.nweight1 * tblcosting.minrate1 + "
str1 = str1 & " tblcosting.nweight2 * tblcosting.minrate2+ tblcosting.nweight3 * tblcosting.minrate3 "
str1 = str1 & " + tblcosting.nweight4 * tblcosting.minrate4 +nmaking1) >" & j
str1 = str1 & " and (tblcosting.nweight1 * tblcosting.minrate1 + tblcosting.nweight2 * tblcosting.minrate2 +"
str1 = str1 & " tblcosting.nweight3 * tblcosting.minrate3+ tblcosting.nweight4 * tblcosting.minrate4 +nmaking1)<=" & k



If (listapp <> "") Then
     pos = InStrRev(listapp, ",")
     issuein = Mid(listapp, 1, pos - 1)
     'listdt = MDIForm1.parserText(listdt)
     'issuein = MDIForm1.parserText(issuein)
    ' MsgBox (issuein)
     If (issuein = "'OT'") Then
     str1 = str1 & " And chowner not in('LE','NJ','DJ','JP')"
     Else
     str1 = str1 & " And chowner in(" & issuein & ")"
     End If
End If
'MsgBox (str1)
str1 = str1 & " GROUP BY chcategory ORDER BY chcategory"
res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
 
 Dim total As Double
 total = 0
 While (res1.EOF = False)
    If (res1!chcategory = "C") Then
    MSHFlexGrid1.TextMatrix(1, i) = res1!no1
    total = total + res1!no1
    ctotal = ctotal + res1!no1
    ctotamt = ctotamt + res1!amt
    ElseIf (res1!chcategory = "GR") Then
    MSHFlexGrid1.TextMatrix(2, i) = res1!no1
    total = total + res1!no1
    grtotal = grtotal + res1!no1
    grtotamt = grtotamt + res1!amt
    ElseIf (res1!chcategory = "K") Then
    MSHFlexGrid1.TextMatrix(3, i) = res1!no1
    total = total + res1!no1
    ktotal = ktotal + res1!no1
    ktotamt = ktotamt + res1!amt
    ElseIf (res1!chcategory = "LR") Then
    MSHFlexGrid1.TextMatrix(4, i) = res1!no1
    total = total + res1!no1
    lrtotal = lrtotal + res1!no1
    lrtotamt = lrtotamt + res1!amt
    ElseIf (res1!chcategory = "O") Then
    MSHFlexGrid1.TextMatrix(5, i) = res1!no1
    total = total + res1!no1
    ototal = ototal + res1!no1
    ototamt = ototamt + res1!amt
    ElseIf (res1!chcategory = "P") Then
    MSHFlexGrid1.TextMatrix(6, i) = res1!no1
    total = total + res1!no1
    ptotal = ptotal + res1!no1
    ptotamt = ptotamt + res1!amt
    ElseIf (res1!chcategory = "S") Then
    MSHFlexGrid1.TextMatrix(7, i) = res1!no1
    total = total + res1!no1
    stotal = stotal + res1!no1
    stotamt = stotamt + res1!amt
    ElseIf (res1!chcategory = "SL") Then
    MSHFlexGrid1.TextMatrix(8, i) = res1!no1
    total = total + res1!no1
    sltotal = sltotal + res1!no1
    sltotamt = sltotamt + res1!amt
    ElseIf (res1!chcategory = "T") Then
    MSHFlexGrid1.TextMatrix(9, i) = res1!no1
    total = total + res1!no1
    ttotal = ttotal + res1!no1
    ttotamt = ttotamt + res1!amt
    ElseIf (res1!chcategory = "TP") Then
    MSHFlexGrid1.TextMatrix(10, i) = res1!no1
    total = total + res1!no1
    tptotal = tptotal + res1!no1
    tptotamt = tptotamt + res1!amt
    ElseIf (res1!chcategory = "LS") Then
    MSHFlexGrid1.TextMatrix(11, i) = res1!no1
    total = total + res1!no1
    lstotal = lstotal + res1!no1
    lstotamt = lstotamt + res1!amt
    ElseIf (res1!chcategory = "W") Then
    MSHFlexGrid1.TextMatrix(12, i) = res1!no1
    total = total + res1!no1
    wtotal = wtotal + res1!no1
    wtotamt = wtotamt + res1!amt
    End If

    res1.MoveNext
 Wend
    MSHFlexGrid1.TextMatrix(13, i) = total
    MSHFlexGrid1.row = 13
    MSHFlexGrid1.col = i
    MSHFlexGrid1.CellFontBold = True
    MSHFlexGrid1.CellForeColor = vbBlue
 res1.close
Next i
Dim gttotal As Double
gttotal = 0
For l = 1 To 12
If (l = 1) Then
MSHFlexGrid1.row = l
MSHFlexGrid1.col = 10
MSHFlexGrid1.CellFontBold = True
MSHFlexGrid1.CellForeColor = vbRed
MSHFlexGrid1.TextMatrix(l, 10) = ctotal
MSHFlexGrid1.TextMatrix(l, 11) = Round(ctotamt)
gttotal = gttotal + ctotal
ElseIf (l = 2) Then
MSHFlexGrid1.row = l
MSHFlexGrid1.col = 10
MSHFlexGrid1.CellFontBold = True
MSHFlexGrid1.CellForeColor = vbRed
MSHFlexGrid1.TextMatrix(l, 10) = grtotal
MSHFlexGrid1.TextMatrix(l, 11) = Round(grtotamt)
gttotal = gttotal + grtotal
ElseIf (l = 3) Then
MSHFlexGrid1.row = l
MSHFlexGrid1.col = 10
MSHFlexGrid1.CellFontBold = True
MSHFlexGrid1.CellForeColor = vbRed
MSHFlexGrid1.TextMatrix(l, 10) = ktotal
MSHFlexGrid1.TextMatrix(l, 11) = Round(ktotamt)
gttotal = gttotal + ktotal
ElseIf (l = 4) Then
MSHFlexGrid1.row = l
MSHFlexGrid1.col = 10
MSHFlexGrid1.CellFontBold = True
MSHFlexGrid1.CellForeColor = vbRed
MSHFlexGrid1.TextMatrix(l, 10) = lrtotal
MSHFlexGrid1.TextMatrix(l, 11) = Round(lrtotamt)
gttotal = gttotal + lrtotal
ElseIf (l = 5) Then
MSHFlexGrid1.row = l
MSHFlexGrid1.col = 10
MSHFlexGrid1.CellFontBold = True
MSHFlexGrid1.CellForeColor = vbRed
MSHFlexGrid1.TextMatrix(l, 10) = ototal
MSHFlexGrid1.TextMatrix(l, 11) = Round(ototamt)
gttotal = gttotal + ototal
ElseIf (l = 6) Then
MSHFlexGrid1.row = l
MSHFlexGrid1.col = 10
MSHFlexGrid1.CellFontBold = True
MSHFlexGrid1.CellForeColor = vbRed
MSHFlexGrid1.TextMatrix(l, 10) = ptotal
MSHFlexGrid1.TextMatrix(l, 11) = Round(ptotamt)
gttotal = gttotal + ptotal
ElseIf (l = 7) Then
MSHFlexGrid1.row = l
MSHFlexGrid1.col = 10
MSHFlexGrid1.CellFontBold = True
MSHFlexGrid1.CellForeColor = vbRed
MSHFlexGrid1.TextMatrix(l, 10) = stotal
MSHFlexGrid1.TextMatrix(l, 11) = Round(stotamt)
gttotal = gttotal + stotal
ElseIf (l = 8) Then
MSHFlexGrid1.row = l
MSHFlexGrid1.col = 10
MSHFlexGrid1.CellFontBold = True
MSHFlexGrid1.CellForeColor = vbRed
MSHFlexGrid1.TextMatrix(l, 10) = sltotal
MSHFlexGrid1.TextMatrix(l, 11) = Round(sltotamt)
gttotal = gttotal + sltotal
ElseIf (l = 9) Then
MSHFlexGrid1.row = l
MSHFlexGrid1.col = 10
MSHFlexGrid1.CellFontBold = True
MSHFlexGrid1.CellForeColor = vbRed
MSHFlexGrid1.TextMatrix(l, 10) = ttotal
MSHFlexGrid1.TextMatrix(l, 11) = Round(ttotamt)
gttotal = gttotal + ttotal
ElseIf (l = 10) Then
MSHFlexGrid1.row = l
MSHFlexGrid1.col = 10
MSHFlexGrid1.CellFontBold = True
MSHFlexGrid1.CellForeColor = vbRed
MSHFlexGrid1.TextMatrix(l, 10) = tptotal
MSHFlexGrid1.TextMatrix(l, 11) = Round(tptotamt)
gttotal = gttotal + tptotal
ElseIf (l = 11) Then
MSHFlexGrid1.row = l
MSHFlexGrid1.col = 10
MSHFlexGrid1.CellFontBold = True
MSHFlexGrid1.CellForeColor = vbRed
MSHFlexGrid1.TextMatrix(l, 10) = lstotal
MSHFlexGrid1.TextMatrix(l, 11) = Round(lstotamt)
gttotal = gttotal + lstotal
ElseIf (l = 12) Then
MSHFlexGrid1.row = l
MSHFlexGrid1.col = 10
MSHFlexGrid1.CellFontBold = True
MSHFlexGrid1.CellForeColor = vbRed
MSHFlexGrid1.TextMatrix(l, 10) = wtotal
MSHFlexGrid1.TextMatrix(l, 11) = Round(wtotamt)
gttotal = gttotal + wtotal
End If
Next l
MSHFlexGrid1.row = 13
MSHFlexGrid1.col = 10
MSHFlexGrid1.CellFontBold = True
MSHFlexGrid1.CellForeColor = vbRed
MSHFlexGrid1.TextMatrix(13, 10) = gttotal
amttotal = tptotamt + ttotamt + sltotamt + grtotamt + stotamt + lrtotamt + ototamt + ptotamt + ktotamt + ctotamt + lstotamt + wtotamt
MSHFlexGrid1.TextMatrix(13, 11) = Round(amttotal)
End Sub
