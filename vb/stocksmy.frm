VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form stocksmy 
   BackColor       =   &H8000000E&
   Caption         =   "Stock Summary"
   ClientHeight    =   6465
   ClientLeft      =   3540
   ClientTop       =   1830
   ClientWidth     =   9435
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form8"
   MDIChild        =   -1  'True
   ScaleHeight     =   6465
   ScaleWidth      =   9435
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Clear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton Search 
      Caption         =   "&Search"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton close1 
      Caption         =   "&Close"
      Height          =   375
      Left            =   7680
      TabIndex        =   5
      Top             =   5760
      Width           =   1335
   End
   Begin VB.ComboBox sowner 
      Height          =   315
      ItemData        =   "stocksmy.frx":0000
      Left            =   2040
      List            =   "stocksmy.frx":000D
      TabIndex        =   4
      Top             =   5520
      Width           =   615
   End
   Begin VB.CommandButton print12 
      Caption         =   "&Print "
      Height          =   375
      Left            =   6120
      TabIndex        =   3
      Top             =   5760
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   7435
      _Version        =   393216
      Rows            =   13
      Cols            =   12
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   -2147483624
      ForeColorFixed  =   -2147483640
      BackColorBkg    =   16777215
      GridColor       =   8421504
      GridColorFixed  =   12632256
      FocusRect       =   0
      HighLight       =   0
      PictureType     =   1
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Owner Type.:"
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Date:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   615
   End
   Begin VB.Label date 
      BackColor       =   &H8000000E&
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "stocksmy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim con1 As New ADODB.Connection
Dim res1 As New ADODB.Recordset
Private Sub Clear_Click()
For i = 1 To 12
    MSFlexGrid1.TextMatrix(i, 1) = ""
    MSFlexGrid1.TextMatrix(i, 2) = ""
    MSFlexGrid1.TextMatrix(i, 3) = ""
    MSFlexGrid1.TextMatrix(i, 4) = ""
    MSFlexGrid1.TextMatrix(i, 5) = ""
    MSFlexGrid1.TextMatrix(i, 6) = ""
    MSFlexGrid1.TextMatrix(i, 7) = ""
    MSFlexGrid1.TextMatrix(i, 8) = ""
    MSFlexGrid1.TextMatrix(i, 9) = ""
    MSFlexGrid1.TextMatrix(i, 10) = ""
    MSFlexGrid1.TextMatrix(i, 11) = ""
   ' MSFlexGrid1.TextMatrix(I, 12) = ""
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
MSFlexGrid1.TextMatrix(0, 0) = "Item"
MSFlexGrid1.ColWidth(0) = 950
MSFlexGrid1.TextMatrix(0, 1) = "0-25"
MSFlexGrid1.ColWidth(1) = 700
MSFlexGrid1.TextMatrix(0, 2) = "25-100"
MSFlexGrid1.ColWidth(2) = 900
MSFlexGrid1.TextMatrix(0, 3) = "100-300"
MSFlexGrid1.ColWidth(3) = 900
MSFlexGrid1.TextMatrix(0, 4) = "300-500"
MSFlexGrid1.ColWidth(4) = 900
MSFlexGrid1.TextMatrix(0, 5) = "500-1000"
MSFlexGrid1.ColWidth(5) = 900
MSFlexGrid1.TextMatrix(0, 6) = "1000-1500"
MSFlexGrid1.ColWidth(6) = 900
MSFlexGrid1.TextMatrix(0, 7) = "1500-2000"
MSFlexGrid1.ColWidth(7) = 900
MSFlexGrid1.TextMatrix(0, 8) = "2000-2500"
MSFlexGrid1.ColWidth(8) = 900
MSFlexGrid1.TextMatrix(0, 9) = ">2500"
MSFlexGrid1.ColWidth(9) = 900
MSFlexGrid1.TextMatrix(0, 10) = "I.Total"
MSFlexGrid1.ColWidth(10) = 900
MSFlexGrid1.TextMatrix(0, 11) = "Tot.Amt."
MSFlexGrid1.ColWidth(11) = 950
MSFlexGrid1.TextMatrix(1, 0) = "Churi"
MSFlexGrid1.TextMatrix(2, 0) = "GRing"
MSFlexGrid1.TextMatrix(3, 0) = "Kara"
MSFlexGrid1.TextMatrix(4, 0) = "LRing"
MSFlexGrid1.TextMatrix(5, 0) = "Others"
MSFlexGrid1.TextMatrix(6, 0) = "Pendent"
MSFlexGrid1.TextMatrix(7, 0) = "Set"
MSFlexGrid1.TextMatrix(8, 0) = "S.Line"
MSFlexGrid1.TextMatrix(9, 0) = "Tops"
MSFlexGrid1.TextMatrix(10, 0) = "TP"
MSFlexGrid1.TextMatrix(11, 0) = "LS"
MSFlexGrid1.TextMatrix(12, 0) = "Total"
End Sub



Private Sub print12_Click()
Search.Visible = False
Clear.Visible = False
close1.Visible = False
sowner.Visible = False
Label1.Visible = False
print12.Visible = False
stocksmy.PrintForm
print12.Visible = True
Label1.Visible = True
Search.Visible = True
Clear.Visible = True
close1.Visible = True
sowner.Visible = True
End Sub

Private Sub Search_Click()
Clear_Click
Dim j As Double
Dim k As Double
Dim lstotal, lstotamt, stotal, ctotal, ctotamt, grtotal, grtotamt, lrtotal, ktotal, ktotamt, sltotal, sltotamt, ototal, ototamt, ptotal, pttotamt, ttotal, ttotamt, tptotal, tptotamt, amttotal As Double
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
If (sowner.Text <> "" And sowner.Text <> "ALL") Then
str1 = str1 & " and chowner=" & "'" & sowner.Text & "'"
End If
str1 = str1 & " GROUP BY chcategory ORDER BY chcategory"
res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
 
 Dim total As Double
 total = 0
 While (res1.EOF = False)
    If (res1!chcategory = "C") Then
    MSFlexGrid1.TextMatrix(1, i) = res1!no1
    total = total + res1!no1
    ctotal = ctotal + res1!no1
    ctotamt = ctotamt + res1!amt
    ElseIf (res1!chcategory = "GR") Then
    MSFlexGrid1.TextMatrix(2, i) = res1!no1
    total = total + res1!no1
    grtotal = grtotal + res1!no1
    grtotamt = grtotamt + res1!amt
    ElseIf (res1!chcategory = "K") Then
    MSFlexGrid1.TextMatrix(3, i) = res1!no1
    total = total + res1!no1
    ktotal = ktotal + res1!no1
    ktotamt = ktotamt + res1!amt
    ElseIf (res1!chcategory = "LR") Then
    MSFlexGrid1.TextMatrix(4, i) = res1!no1
    total = total + res1!no1
    lrtotal = lrtotal + res1!no1
    lrtotamt = lrtotamt + res1!amt
    ElseIf (res1!chcategory = "O") Then
    MSFlexGrid1.TextMatrix(5, i) = res1!no1
    total = total + res1!no1
    ototal = ototal + res1!no1
    ototamt = ototamt + res1!amt
    ElseIf (res1!chcategory = "P") Then
    MSFlexGrid1.TextMatrix(6, i) = res1!no1
    total = total + res1!no1
    ptotal = ptotal + res1!no1
    ptotamt = ptotamt + res1!amt
    ElseIf (res1!chcategory = "S") Then
    MSFlexGrid1.TextMatrix(7, i) = res1!no1
    total = total + res1!no1
    stotal = stotal + res1!no1
    stotamt = stotamt + res1!amt
    ElseIf (res1!chcategory = "SL") Then
    MSFlexGrid1.TextMatrix(8, i) = res1!no1
    total = total + res1!no1
    sltotal = sltotal + res1!no1
    sltotamt = sltotamt + res1!amt
    ElseIf (res1!chcategory = "T") Then
    MSFlexGrid1.TextMatrix(9, i) = res1!no1
    total = total + res1!no1
    ttotal = ttotal + res1!no1
    ttotamt = ttotamt + res1!amt
    ElseIf (res1!chcategory = "TP") Then
    MSFlexGrid1.TextMatrix(10, i) = res1!no1
    total = total + res1!no1
    tptotal = tptotal + res1!no1
    tptotamt = tptotamt + res1!amt
    ElseIf (res1!chcategory = "LS") Then
    MSFlexGrid1.TextMatrix(11, i) = res1!no1
    total = total + res1!no1
    lstotal = lstotal + res1!no1
    lstotamt = lstotamt + res1!amt
    End If

    res1.MoveNext
 Wend
    MSFlexGrid1.TextMatrix(12, i) = total
    MSFlexGrid1.row = 12
    MSFlexGrid1.col = i
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.CellForeColor = vbBlue
 res1.close
Next i
Dim gttotal As Double
gttotal = 0
For l = 1 To 11
If (l = 1) Then
MSFlexGrid1.row = l
MSFlexGrid1.col = 10
MSFlexGrid1.CellFontBold = True
MSFlexGrid1.CellForeColor = vbRed
MSFlexGrid1.TextMatrix(l, 10) = ctotal
MSFlexGrid1.TextMatrix(l, 11) = Round(ctotamt)
gttotal = gttotal + ctotal
ElseIf (l = 2) Then
MSFlexGrid1.row = l
MSFlexGrid1.col = 10
MSFlexGrid1.CellFontBold = True
MSFlexGrid1.CellForeColor = vbRed
MSFlexGrid1.TextMatrix(l, 10) = grtotal
MSFlexGrid1.TextMatrix(l, 11) = Round(grtotamt)
gttotal = gttotal + grtotal
ElseIf (l = 3) Then
MSFlexGrid1.row = l
MSFlexGrid1.col = 10
MSFlexGrid1.CellFontBold = True
MSFlexGrid1.CellForeColor = vbRed
MSFlexGrid1.TextMatrix(l, 10) = ktotal
MSFlexGrid1.TextMatrix(l, 11) = Round(ktotamt)
gttotal = gttotal + ktotal
ElseIf (l = 4) Then
MSFlexGrid1.row = l
MSFlexGrid1.col = 10
MSFlexGrid1.CellFontBold = True
MSFlexGrid1.CellForeColor = vbRed
MSFlexGrid1.TextMatrix(l, 10) = lrtotal
MSFlexGrid1.TextMatrix(l, 11) = Round(lrtotamt)
gttotal = gttotal + lrtotal
ElseIf (l = 5) Then
MSFlexGrid1.row = l
MSFlexGrid1.col = 10
MSFlexGrid1.CellFontBold = True
MSFlexGrid1.CellForeColor = vbRed
MSFlexGrid1.TextMatrix(l, 10) = ototal
MSFlexGrid1.TextMatrix(l, 11) = Round(ototamt)
gttotal = gttotal + ototal
ElseIf (l = 6) Then
MSFlexGrid1.row = l
MSFlexGrid1.col = 10
MSFlexGrid1.CellFontBold = True
MSFlexGrid1.CellForeColor = vbRed
MSFlexGrid1.TextMatrix(l, 10) = ptotal
MSFlexGrid1.TextMatrix(l, 11) = Round(ptotamt)
gttotal = gttotal + ptotal
ElseIf (l = 7) Then
MSFlexGrid1.row = l
MSFlexGrid1.col = 10
MSFlexGrid1.CellFontBold = True
MSFlexGrid1.CellForeColor = vbRed
MSFlexGrid1.TextMatrix(l, 10) = stotal
MSFlexGrid1.TextMatrix(l, 11) = Round(stotamt)
gttotal = gttotal + stotal
ElseIf (l = 8) Then
MSFlexGrid1.row = l
MSFlexGrid1.col = 10
MSFlexGrid1.CellFontBold = True
MSFlexGrid1.CellForeColor = vbRed
MSFlexGrid1.TextMatrix(l, 10) = sltotal
MSFlexGrid1.TextMatrix(l, 11) = Round(sltotamt)
gttotal = gttotal + sltotal
ElseIf (l = 9) Then
MSFlexGrid1.row = l
MSFlexGrid1.col = 10
MSFlexGrid1.CellFontBold = True
MSFlexGrid1.CellForeColor = vbRed
MSFlexGrid1.TextMatrix(l, 10) = ttotal
MSFlexGrid1.TextMatrix(l, 11) = Round(ttotamt)
gttotal = gttotal + ttotal
ElseIf (l = 10) Then
MSFlexGrid1.row = l
MSFlexGrid1.col = 10
MSFlexGrid1.CellFontBold = True
MSFlexGrid1.CellForeColor = vbRed
MSFlexGrid1.TextMatrix(l, 10) = tptotal
MSFlexGrid1.TextMatrix(l, 11) = Round(tptotamt)
gttotal = gttotal + tptotal
ElseIf (l = 11) Then
MSFlexGrid1.row = l
MSFlexGrid1.col = 10
MSFlexGrid1.CellFontBold = True
MSFlexGrid1.CellForeColor = vbRed
MSFlexGrid1.TextMatrix(l, 10) = lstotal
MSFlexGrid1.TextMatrix(l, 11) = Round(lstotamt)
gttotal = gttotal + lstotal
End If
Next l
MSFlexGrid1.row = 12
MSFlexGrid1.col = 10
MSFlexGrid1.CellFontBold = True
MSFlexGrid1.CellForeColor = vbRed
MSFlexGrid1.TextMatrix(12, 10) = gttotal
amttotal = tptotamt + ttotamt + sltotamt + grtotamt + stotamt + lrtotamt + ototamt + ptotamt + ktotamt + ctotamt + lstotamt
MSFlexGrid1.TextMatrix(12, 11) = Round(amttotal)
End Sub

