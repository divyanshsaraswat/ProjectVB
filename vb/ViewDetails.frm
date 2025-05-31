VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form ViewDetails 
   Caption         =   "View Details"
   ClientHeight    =   5880
   ClientLeft      =   630
   ClientTop       =   2340
   ClientWidth     =   10710
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   10710
   WindowState     =   2  'Maximized
   Begin VB.ComboBox sowner 
      Height          =   315
      ItemData        =   "ViewDetails.frx":0000
      Left            =   10320
      List            =   "ViewDetails.frx":000D
      TabIndex        =   14
      Top             =   240
      Width           =   615
   End
   Begin VB.ComboBox Category 
      Height          =   315
      ItemData        =   "ViewDetails.frx":001D
      Left            =   11040
      List            =   "ViewDetails.frx":003F
      TabIndex        =   12
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox weight2 
      Height          =   315
      Left            =   8640
      TabIndex        =   11
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox weight1 
      Height          =   315
      Left            =   6720
      TabIndex        =   10
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton Clear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Search 
      Caption         =   "&Submit"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton close 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox drate2 
      Height          =   315
      Left            =   4320
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox drate 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6375
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   11245
      _Version        =   393216
      Cols            =   17
      AllowUserResizing=   3
   End
   Begin VB.Image Image1 
      Height          =   2775
      Left            =   9240
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Item Type."
      Height          =   375
      Left            =   9480
      TabIndex        =   13
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "Wt2.(Ct.)"
      Height          =   375
      Left            =   7440
      TabIndex        =   9
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "Wt1.(Ct.)"
      Height          =   375
      Left            =   5400
      TabIndex        =   8
      Top             =   240
      Width           =   1095
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   11880
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   11880
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label drate1 
      Caption         =   "Dia. Rate2.:"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.Label drte 
      Caption         =   "Dia. Rate1.:"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "ViewDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con1 As New ADODB.Connection
Dim res1 As New ADODB.Recordset
Private Sub Clear_Click()
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
End Sub

Private Sub close_Click()
'con1.close
Unload Me
End Sub
Private Sub Form_Load()
'con1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\lahar\vb\db2.mdb;Persist Security Info=False"
'con1.Open
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
MSFlexGrid1.ColWidth(15) = 1200
MSFlexGrid1.TextMatrix(0, 16) = "Image."
MSFlexGrid1.ColWidth(16) = 1200
End Sub
Private Sub MSFlexGrid1_SelChange()

If (MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 16) <> "" And MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 16) <> Empty) Then
    Dim picturename As String
     pos = InStrRev(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 16), "\")
     If (pos <> 0) Then
     picturename = Mid(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 16), pos + 1)
     Else
     picturename = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 16)
     End If
     Image1.Visible = True
     Image1.Picture = LoadPicture(MDIForm1.picturepath & picturename)
Else
    Image1.Picture = LoadPicture()
End If
End Sub

Private Sub Search_Click()
clearing
str1 = "select * from tblcosting where "
Dim str2 As String

If (Val(weight1.Text) >= Val(weight2.Text) And weight1.Text <> Empty And weight2.Text <> Empty) Then
    MsgBox ("Enter The Valid Dia. Wt. Range")
Else

If (Val(drate.Text) >= Val(drate2.Text) And drate.Text <> Empty And drate2.Text <> Empty) Then
    MsgBox ("Enter The Valid Dia. Range")
 
Else
    If (drate.Text <> Empty) Then
       str2 = str2 & "  minrate1>=" & Val(drate.Text)
    End If
    
    If (drate2.Text <> Empty) Then
        If (str2 <> Empty) Then
            str2 = str2 & "  And minrate1<=" & Val(drate2.Text)
        Else
            str2 = str2 & "  minrate1<=" & Val(drate2.Text)
        End If
    End If
    
    If (weight1.Text <> Empty) Then
        If (str2 <> Empty) Then
            str2 = str2 & "  And nweight1>=" & Val(weight1.Text)
        Else
            str2 = str2 & "  nweight1>=" & Val(weight1.Text)
        End If
    End If
    
    If (weight2.Text <> Empty) Then
        If (str2 <> Empty) Then
            str2 = str2 & "  and nweight1<=" & Val(weight2.Text)
        Else
            str2 = str2 & "  nweight1<=" & Val(weight2.Text)
        End If
    End If
    
   If (sowner.Text <> Empty) Then
        If (str2 <> Empty) Then
            str2 = str2 & "  and chowner =" & "'" & sowner.Text & "'"
        Else
            str2 = str2 & "  chowner =" & "'" & sowner.Text & "'"
        End If

   End If
   
   If (Category.Text <> Empty) Then
        If (str2 <> Empty) Then
            str2 = str2 & "  and chcategory =" & "'" & Category.Text & "'"
        Else
            str2 = str2 & "  chcategory =" & "'" & Category.Text & "'"
        End If

   End If
   
    str1 = str1 & str2 & " order by chcategory,ncode"

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
    MSFlexGrid1.TextMatrix(i, 5) = res1!minrate1
    MSFlexGrid1.TextMatrix(i, 6) = res1!chcontent2
    swt1 = res1!nweight2
    MSFlexGrid1.TextMatrix(i, 7) = swt1
    MSFlexGrid1.TextMatrix(i, 8) = res1!minrate2
    MSFlexGrid1.TextMatrix(i, 9) = res1!chcontent3
    swt2 = res1!nweight3
    MSFlexGrid1.TextMatrix(i, 10) = swt2
    MSFlexGrid1.TextMatrix(i, 11) = res1!minrate3
    mwt = res1!nweight4
    MSFlexGrid1.TextMatrix(i, 12) = mwt
    
    MSFlexGrid1.TextMatrix(i, 13) = res1!minrate4
    making = res1!nmaking1
    MSFlexGrid1.TextMatrix(i, 14) = making
      
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
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 16

    
    If (IsNull(res1!opicture)) Then
        MSFlexGrid1.TextMatrix(i, 16) = ""
    ElseIf (res1!opicture <> "" And res1!opicture <> Empty) Then
        MSFlexGrid1.TextMatrix(i, 16) = res1!opicture
        'Set MSFlexGrid1.CellPicture = LoadPicture(res1!opicture)
       
    End If
    
    
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
    MSFlexGrid1.TextMatrix(i, 4) = Round(dwt, 2)
    MSFlexGrid1.TextMatrix(i, 5) = Round(diatotrt)
    MSFlexGrid1.TextMatrix(i, 7) = Round(stonewt1, 2)
    MSFlexGrid1.TextMatrix(i, 8) = Round(stonetotamt1)
    MSFlexGrid1.TextMatrix(i, 10) = Round(stonewt2, 2)
    MSFlexGrid1.TextMatrix(i, 11) = Round(stonetotamt2)
    MSFlexGrid1.TextMatrix(i, 12) = Round(metal, 2)
    MSFlexGrid1.TextMatrix(i, 13) = Round(metalamt)
    MSFlexGrid1.TextMatrix(i, 14) = Round(makingtot)
    MSFlexGrid1.TextMatrix(i, 15) = Round(gtotal)
    
    
End If
    res1.close
  
End If
End If


End Sub
