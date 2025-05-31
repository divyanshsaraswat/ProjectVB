VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form2 
   Caption         =   "salesheet"
   ClientHeight    =   5715
   ClientLeft      =   135
   ClientTop       =   1305
   ClientWidth     =   11730
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5715
   ScaleWidth      =   11730
   WindowState     =   2  'Maximized
   Begin VB.CommandButton report 
      Caption         =   "Report"
      Height          =   375
      Left            =   10440
      TabIndex        =   25
      Top             =   480
      Width           =   1335
   End
   Begin VB.ComboBox content2 
      Height          =   315
      ItemData        =   "DetailsCost.frx":0000
      Left            =   4800
      List            =   "DetailsCost.frx":0043
      TabIndex        =   24
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton close 
      Caption         =   "&Close"
      Height          =   375
      Left            =   10440
      TabIndex        =   23
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Clear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   10440
      TabIndex        =   22
      Top             =   840
      Width           =   1335
   End
   Begin VB.ComboBox content4 
      Height          =   315
      ItemData        =   "DetailsCost.frx":00C2
      Left            =   6720
      List            =   "DetailsCost.frx":00D5
      TabIndex        =   21
      Top             =   720
      Width           =   1095
   End
   Begin VB.ComboBox state 
      Height          =   315
      ItemData        =   "DetailsCost.frx":0104
      Left            =   9120
      List            =   "DetailsCost.frx":010E
      TabIndex        =   18
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox issueto 
      Height          =   315
      Left            =   3960
      TabIndex        =   17
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox maker 
      Height          =   315
      Left            =   1320
      TabIndex        =   16
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox edate 
      Height          =   315
      Left            =   6720
      TabIndex        =   15
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox sdate 
      Height          =   315
      Left            =   3960
      TabIndex        =   14
      Top             =   120
      Width           =   1095
   End
   Begin VB.ComboBox gpur1 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "DetailsCost.frx":0122
      Left            =   9120
      List            =   "DetailsCost.frx":0145
      TabIndex        =   10
      Top             =   720
      Width           =   735
   End
   Begin VB.ComboBox content1 
      Height          =   315
      ItemData        =   "DetailsCost.frx":0170
      Left            =   3960
      List            =   "DetailsCost.frx":01B3
      TabIndex        =   8
      Top             =   720
      Width           =   735
   End
   Begin VB.ComboBox Category 
      Height          =   315
      ItemData        =   "DetailsCost.frx":0232
      Left            =   1320
      List            =   "DetailsCost.frx":0254
      TabIndex        =   5
      Top             =   720
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5535
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9763
      _Version        =   393216
      Cols            =   16
      AllowUserResizing=   1
      MousePointer    =   1
   End
   Begin VB.CommandButton Search 
      Caption         =   "&Search"
      Height          =   375
      Left            =   10440
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox code 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Metal:"
      Height          =   375
      Left            =   5640
      TabIndex        =   20
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lState 
      Alignment       =   2  'Center
      Caption         =   "St."
      Height          =   375
      Left            =   8040
      TabIndex        =   19
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label13 
      Caption         =   "Issue To:"
      Height          =   255
      Left            =   2640
      TabIndex        =   13
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Maker:"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Metal Qly.:"
      Height          =   375
      Left            =   8040
      TabIndex        =   11
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Stone:"
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "End Date:"
      Height          =   375
      Left            =   5640
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "St. Date:"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Category:"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Item Code:"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   145
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   11880
      Y1              =   1920
      Y2              =   1920
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

Private Sub clear_Click()
Clearing
End Sub
Private Sub CLOSE_Click()
Unload Me
con2.CLOSE
End Sub

Private Sub Form_Load()
con2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\lahar\vb\db2.mdb;Persist Security Info=False"
con2.Open
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
MSFlexGrid1.TextMatrix(0, 15) = "Total Amt."
MSFlexGrid1.ColWidth(15) = 950
End Sub
Private Sub Clearing()
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

End Sub



Private Sub report_Click()
If (str1 = Empty) Then
str1 = "SELECT tblcosting.*, Tblcosting.minrate1 * tblcosting.nweight1 + Tblcosting.minrate2 * tblcosting.nweight2 + Tblcosting.minrate3 * tblcosting.nweight3+Tblcosting.minrate4 * tblcosting.nweight4+nmaking1 AS total From tblcosting ORDER BY chcategory,ncode"
End If
DataEnvironment2.Commands(1).CommandText = str1
DataReport1.PrintReport True
End Sub

Private Sub Search_Click()
Clearing
Dim str2 As String
str2 = ""

str1 = "select Tblcosting.*, Tblcosting.minrate1 * tblcosting.nweight1 + Tblcosting.minrate2 * tblcosting.nweight2 + Tblcosting.minrate3 * tblcosting.nweight3+Tblcosting.minrate4 * tblcosting.nweight4+nmaking1 AS total from tblcosting"
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
    str2 = str2 & " And dinvoicedate<=#" & Format(sdate.Text, "dd-mmm-yy") & "# "
    Else
    str2 = str2 & " dinvoicedate<=#" & Format(sdate.Text, "dd-mmm-yy") & "# "
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

If (issueto.Text <> Empty) Then
    If (str2 <> Empty) Then
    str2 = str2 & " And chissueto=" & "'" & issueto & "'"
    Else
    str2 = str2 & "  chissueto=" & "'" & issueto & "'"
    End If
End If

If (str2 <> Empty) Then
   str1 = str1 & " where " & str2 & " order by chcategory,ncode"
   Else
   str1 = str1 & " order by chcategory,ncode"
End If

'Debug.Print (str1)

res1.Open str1, con2, adOpenDynamic, adLockOptimistic
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
    MSFlexGrid1.Row = i
    MSFlexGrid1.Col = 4
    MSFlexGrid1.CellBackColor = RGB(231, 228, 218)
    MSFlexGrid1.TextMatrix(i, 5) = res1!minrate1
    MSFlexGrid1.Row = i
    MSFlexGrid1.Col = 5
    MSFlexGrid1.CellBackColor = RGB(231, 228, 218)
    
    MSFlexGrid1.TextMatrix(i, 6) = res1!chcontent2
    swt1 = res1!nweight2
    MSFlexGrid1.TextMatrix(i, 7) = swt1
    MSFlexGrid1.TextMatrix(i, 8) = res1!minrate2
    
    MSFlexGrid1.Row = i
    MSFlexGrid1.Col = 7
    MSFlexGrid1.CellBackColor = RGB(231, 228, 218)
    
    MSFlexGrid1.Row = i
    MSFlexGrid1.Col = 8
    MSFlexGrid1.CellBackColor = RGB(231, 228, 218)
    
    MSFlexGrid1.TextMatrix(i, 9) = res1!chcontent3
    swt2 = res1!nweight3
    MSFlexGrid1.TextMatrix(i, 10) = swt2
    MSFlexGrid1.TextMatrix(i, 11) = res1!minrate3
    MSFlexGrid1.Row = i
    MSFlexGrid1.Col = 10
    MSFlexGrid1.CellBackColor = RGB(231, 228, 218)
    MSFlexGrid1.Row = i
    MSFlexGrid1.Col = 11
    MSFlexGrid1.CellBackColor = RGB(231, 228, 218)
    
    mwt = res1!nweight4
    MSFlexGrid1.TextMatrix(i, 12) = mwt
    MSFlexGrid1.TextMatrix(i, 13) = res1!minrate4
    
    MSFlexGrid1.Row = i
    MSFlexGrid1.Col = 12
    MSFlexGrid1.CellBackColor = RGB(234, 248, 222)
    MSFlexGrid1.Row = i
    MSFlexGrid1.Col = 13
    MSFlexGrid1.CellBackColor = RGB(234, 248, 222)
    
    making = res1!nmaking1
    MSFlexGrid1.TextMatrix(i, 14) = making
    
    MSFlexGrid1.Row = i
    MSFlexGrid1.Col = 14
    MSFlexGrid1.CellBackColor = RGB(255, 255, 236)
      
    amt1 = Val(res1!minrate1) * Val(res1!nweight1)
    amt2 = Val(res1!minrate2) * Val(res1!nweight2)
    amt3 = Val(res1!minrate3) * Val(res1!nweight3)
    amt4 = Val(res1!minrate4) * Val(res1!nweight4)
    metalamt = metalamt + amt4
    stonetotamt1 = stonetotamt1 + amt2
    stonetotamt2 = stonetotamt2 + amt3
    diatotrt = diatotrt + amt1
    totamt = Val(amt1) + Val(amt2) + Val(amt3) + Val(amt4) + Val(res1!nmaking1)
    MSFlexGrid1.TextMatrix(i, 15) = totamt
    
    MSFlexGrid1.Row = i
    MSFlexGrid1.Col = 15
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.CellForeColor = vbBlue
    
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
    MSFlexGrid1.Row = i
    MSFlexGrid1.Col = 1
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
       
    MSFlexGrid1.TextMatrix(i, 4) = Round(dwt, 2)
    MSFlexGrid1.Row = i
    MSFlexGrid1.Col = 4
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
        
    MSFlexGrid1.TextMatrix(i, 5) = Round(diatotrt)
    
    MSFlexGrid1.Row = i
    MSFlexGrid1.Col = 5
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
        
    MSFlexGrid1.TextMatrix(i, 7) = Round(stonewt1, 2)
    
    MSFlexGrid1.Row = i
    MSFlexGrid1.Col = 7
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
    
    MSFlexGrid1.TextMatrix(i, 8) = Round(stonetotamt1)
    
    MSFlexGrid1.Row = i
    MSFlexGrid1.Col = 8
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
    
    MSFlexGrid1.TextMatrix(i, 10) = Round(stonewt2, 2)
    MSFlexGrid1.Row = i
    MSFlexGrid1.Col = 10
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
    
    MSFlexGrid1.TextMatrix(i, 11) = Round(stonetotamt2)
    MSFlexGrid1.Row = i
    MSFlexGrid1.Col = 11
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
    
    MSFlexGrid1.TextMatrix(i, 12) = Round(metal, 2)
    MSFlexGrid1.Row = i
    MSFlexGrid1.Col = 12
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
    
    MSFlexGrid1.TextMatrix(i, 13) = Round(metalamt)
    MSFlexGrid1.Row = i
    MSFlexGrid1.Col = 13
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
    
    MSFlexGrid1.TextMatrix(i, 14) = Round(makingtot)
    MSFlexGrid1.Row = i
    MSFlexGrid1.Col = 14
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
    
    MSFlexGrid1.TextMatrix(i, 15) = Round(gtotal)
    MSFlexGrid1.Row = i
    MSFlexGrid1.Col = 15
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
    
    End If
    res1.CLOSE
 End Sub

