VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form7 
   Caption         =   "Tag Printing"
   ClientHeight    =   5715
   ClientLeft      =   3405
   ClientTop       =   1545
   ClientWidth     =   8175
   LinkTopic       =   "Form7"
   MDIChild        =   -1  'True
   ScaleHeight     =   5715
   ScaleWidth      =   8175
   WindowState     =   2  'Maximized
   Begin VB.PictureBox ListView1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5055
      Left            =   480
      ScaleHeight     =   4995
      ScaleWidth      =   8955
      TabIndex        =   22
      Top             =   2640
      Width           =   9015
   End
   Begin VB.ComboBox slabel 
      Height          =   315
      ItemData        =   "Tagform.frx":0000
      Left            =   1200
      List            =   "Tagform.frx":001F
      TabIndex        =   20
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton TEST 
      Caption         =   "&TEST"
      Height          =   375
      Left            =   8880
      TabIndex        =   19
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton print 
      Caption         =   "&Print"
      Height          =   375
      Left            =   7200
      TabIndex        =   18
      Top             =   1800
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Tagform.frx":003F
      Left            =   10560
      List            =   "Tagform.frx":0052
      TabIndex        =   17
      Top             =   1200
      Width           =   615
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Tagform.frx":0067
      Left            =   9120
      List            =   "Tagform.frx":007A
      TabIndex        =   13
      Top             =   1200
      Width           =   615
   End
   Begin VB.ComboBox Category 
      Height          =   315
      ItemData        =   "Tagform.frx":008F
      Left            =   7200
      List            =   "Tagform.frx":00B1
      TabIndex        =   12
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox sdate 
      Height          =   315
      Left            =   1200
      TabIndex        =   7
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox edate 
      Height          =   315
      Left            =   4080
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton close 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton Search 
      Caption         =   "&Submit"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton Clear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox weight1 
      Height          =   315
      Left            =   11280
      TabIndex        =   2
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox srange 
      Height          =   315
      Left            =   9840
      TabIndex        =   1
      Top             =   1200
      Width           =   615
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1575
      Left            =   10320
      TabIndex        =   14
      Top             =   5880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2778
      _Version        =   393216
      Cols            =   9
      AllowUserResizing=   3
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2280
      TabIndex        =   15
      Top             =   1200
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yy"
      Format          =   137297921
      CurrentDate     =   37684
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   5280
      TabIndex        =   16
      Top             =   1200
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yy"
      Format          =   137297921
      CurrentDate     =   37684
   End
   Begin VB.Label Label2 
      Caption         =   "S.Label"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   1800
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   1695
      Left            =   9840
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label date1 
      Caption         =   "S.Date: "
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label date2 
      Caption         =   "E.Date:"
      Height          =   375
      Left            =   2880
      TabIndex        =   10
      Top             =   1200
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   11880
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   11880
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "Category"
      Height          =   375
      Left            =   6000
      TabIndex        =   9
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "Range:"
      Height          =   375
      Left            =   8040
      TabIndex        =   8
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "DAKSH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   8895
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   12000
      Y1              =   1080
      Y2              =   1080
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim res1 As New ADODB.Recordset
Dim str1 As String

Private Sub Command1_Click()

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
End Sub

Private Sub Clear_Click()
clearing
End Sub

Private Sub close_Click()
Unload Me
End Sub



Private Sub DTPicker1_Change()
sdate = DTPicker1
End Sub
Private Sub DTPicker2_Change()
edate = DTPicker2
End Sub

Private Sub Form_Load()
DTPicker1.Format = dtpCustom
DTPicker2.Format = dtpCustom
DTPicker2 = date
DTPicker1 = date
MSFlexGrid1.TextMatrix(0, 0) = "Item Code"
MSFlexGrid1.ColWidth(0) = 850
MSFlexGrid1.TextMatrix(0, 1) = "Dia.Wt."
MSFlexGrid1.ColWidth(1) = 600
MSFlexGrid1.TextMatrix(0, 2) = "Stone1."
MSFlexGrid1.ColWidth(2) = 650
MSFlexGrid1.TextMatrix(0, 3) = "St1 Wt."
MSFlexGrid1.ColWidth(3) = 600
MSFlexGrid1.TextMatrix(0, 4) = "Stone2"
MSFlexGrid1.ColWidth(4) = 650
MSFlexGrid1.TextMatrix(0, 5) = "St2 Wt."
MSFlexGrid1.ColWidth(5) = 600
MSFlexGrid1.TextMatrix(0, 6) = "MetalWt."
MSFlexGrid1.ColWidth(6) = 700
MSFlexGrid1.TextMatrix(0, 7) = "Price No."
MSFlexGrid1.ColWidth(7) = 600
MSFlexGrid1.TextMatrix(0, 8) = "Image"
MSFlexGrid1.ColWidth(8) = 2000
End Sub

Private Sub MSFlexGrid1_SelChange()
If (MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 10) <> "" And MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 10) <> Empty) Then
    Image1.Picture = LoadPicture(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 10))
Else
    Image1.Picture = LoadPicture()
End If
End Sub

Private Sub print_Click()
'If (str1 = Empty) Then
'str1 = "SELECT tblcosting.*, (Tblcosting.minrate1 * tblcosting.nweight1 + Tblcosting.minrate2 * tblcosting.nweight2 + Tblcosting.minrate3 * tblcosting.nweight3+Tblcosting.minrate4 * tblcosting.nweight4)+nmaking1 AS total From tblcosting ORDER BY chcategory,ncode"
'End If
'If DataEnvironment2.rsCommand8.state = adStateOpen Then
'DataEnvironment2.rsCommand8.close
'End If
'DataEnvironment2.Commands(8).CommandText = str1
'MsgBox (str1)

'Printer.CurrentY = Printer.Height / 5
'MsgBox (Printer.CurrentY)
'I = 935
'If (slabel.Text = "") Then
'Tagprinting.TopMargin = 150
'ElseIf (slabel.Text = 2) Then
'Tagprinting.TopMargin = I
'ElseIf (slabel.Text = 3) Then
'Tagprinting.TopMargin = I * 2
'ElseIf (slabel.Text = 4) Then
'Tagprinting.TopMargin = I * 3
'ElseIf (slabel.Text = 5) Then
'Tagprinting.TopMargin = I * 4
'ElseIf (slabel.Text = 6) Then
'Tagprinting.TopMargin = I * 5
'ElseIf (slabel.Text = 7) Then
'Tagprinting.TopMargin = I * 6
'ElseIf (slabel.Text = 8) Then
'Tagprinting.TopMargin = I * 7
'ElseIf (slabel.Text = 9) Then
'Tagprinting.TopMargin = I * 8
'ElseIf (slabel.Text = 10) Then
'Tagprinting.TopMargin = I * 9
'End If

'Tagprinting.Show
'Printer.CurrentY = Printer.Height / 5
'Tagprinting.PrintReport True
End Sub
Private Sub Search_Click()
'clearing
str1 = "select *,((minrate1 * nweight1 + minrate2 * nweight2 + minrate3 * nweight3+minrate4 * nweight4)+nmaking1)* 00.01333333 AS total, IIF(nweight2 > 0, nweight2, '') AS cwt2,IIF(nweight3 > 0, nweight3, '') AS cwt3 from tblcosting where "
Dim str2 As String
 
 'If ((sdate = Empty Or edate = Empty) And srange = Empty) Then
'    MsgBox ("Please Select The Start And End Date Or Any Criteria")
' End Sub
' End If
 
 If (sdate <> Empty) Then
  str2 = " dinvoicedate>= #" & Format(sdate, "dd-mmm-yy") & "# "
 End If
  
  If (edate <> Empty) Then
  If (str2 <> Empty) Then
    str2 = str2 & " And dinvoicedate<=#" & Format(edate, "dd-mmm-yy") & "# "
    Else
    str2 = str2 & " dinvoicedate<=#" & Format(edate, "dd-mmm-yy") & "# "
    End If
 End If
 
  If (Category.Text <> Empty) Then
        If (str2 <> Empty) Then
            str2 = str2 & "  and chcategory =" & "'" & Category.Text & "'"
        Else
            str2 = str2 & "  chcategory =" & "'" & Category.Text & "'"
        End If
   End If
   
   str1 = str1 & str2 & " and chstate='Avi.' order by chcategory,ncode"
  
    res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic

Dim i As Integer

i = 1
If (res1.EOF = True) Then
    MsgBox ("No Record Found")
Else
k = 0
While (res1.EOF = False)
    MSFlexGrid1.TextMatrix(i, 0) = i
    MSFlexGrid1.TextMatrix(i, 1) = res1!chcategory & "-" & res1!ncode
    'Printer.CurrentY = 160
    'Printer.CurrentX = Printer.Width / 45
    'Printer.Print (res1!chcategory & "-" & res1!ncode & "         " & "Dia:" & res1!nweight1)
    MSFlexGrid1.TextMatrix(i, 2) = Format(res1!dinvoicedate, "dd-mmm-yy")
    dweight = res1!nweight1
    MSFlexGrid1.TextMatrix(i, 3) = dweight
    MSFlexGrid1.TextMatrix(i, 4) = res1!chcontent2
    swt1 = res1!nweight2
    MSFlexGrid1.TextMatrix(i, 5) = swt1
    MSFlexGrid1.TextMatrix(i, 6) = res1!chcontent3
    swt2 = res1!nweight3
    MSFlexGrid1.TextMatrix(i, 7) = swt2
    mwt = res1!nweight4
    MSFlexGrid1.TextMatrix(i, 8) = mwt
    
    amt1 = Val(res1!minrate1) * Val(res1!nweight1)
    amt2 = Val(res1!minrate2) * Val(res1!nweight2)
    amt3 = Val(res1!minrate3) * Val(res1!nweight3)
    amt4 = Val(res1!minrate4) * Val(res1!nweight4)
    totamt = Val(amt1) + Val(amt2) + Val(amt3) + Val(amt4) + Val(res1!nmaking1)
    MSFlexGrid1.TextMatrix(i, 9) = Round(totamt * 0.01333333)
    'Printer.CurrentX = Printer.Width / 45
    'Text = "Price No:" & MSFlexGrid1.TextMatrix(I, 9) & "         " & res1!chcontent2 & "," & res1!chcontent3 & " " & Val(res1!nweight2 + res1!nweight3)
    'Printer.Print (Text)
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 10

    
    If (IsNull(res1!opicture)) Then
        MSFlexGrid1.TextMatrix(i, 10) = ""
    ElseIf (res1!opicture <> "" And res1!opicture <> Empty) Then
        MSFlexGrid1.TextMatrix(i, 10) = res1!opicture
       'Set MSFlexGrid1.CellPicture = LoadPicture(res1!opicture)
   End If
    res1.MoveNext
    i = i + 1
    MSFlexGrid1.Rows = i + 1
    Wend
End If
     res1.close
End Sub

Private Sub TEST_Click()
'MsgBox (Printer.Height / 10)
Printer.CurrentY = 160
Printer.CurrentX = Printer.Width / 45
Printer.Print ("First Line start ")
Printer.CurrentX = Printer.Width / 45
Printer.Print ("Second  Line first")
Printer.CurrentX = Printer.Width / 45
Printer.Print ("third LINE ")
Printer.Print ("")
Printer.Print ("")
Printer.CurrentX = Printer.Width / 45
Printer.Print ("Sixth Line")
Printer.CurrentX = Printer.Width / 45
Printer.Print ("Seven Line")
Printer.CurrentX = Printer.Width / 45
Printer.Print ("Eight Line")
Printer.Print ("")
Printer.Print ("")
Printer.CurrentX = Printer.Width / 50
Printer.Print ("Nine Line")
Printer.CurrentX = Printer.Width / 50
Printer.Print ("Ten Line")
Printer.CurrentX = Printer.Width / 50
Printer.Print ("Eleven Line")
Printer.Print ("")
Printer.CurrentX = Printer.Width / 50
Printer.Print ("Tw Line")
Printer.CurrentX = Printer.Width / 50
Printer.Print ("Th Line")
Printer.CurrentX = Printer.Width / 50
Printer.Print ("Fou Line")
Printer.Print ("")
Printer.Print ("")
Printer.CurrentX = Printer.Width / 50
Printer.Print ("Six Line")
Printer.CurrentX = Printer.Width / 50
Printer.Print ("Se Line")
Printer.CurrentX = Printer.Width / 50
Printer.Print ("Ei Line")
Printer.Print ("")
Printer.Print ("")
Printer.CurrentX = Printer.Width / 50
Printer.Print ("New1 Line")
Printer.CurrentX = Printer.Width / 50
Printer.Print ("New2 Line")
Printer.CurrentX = Printer.Width / 50
Printer.Print ("New3 Line")
Printer.Print ("")
Printer.CurrentX = Printer.Width / 50
Printer.Print ("New4 Line")
Printer.CurrentX = Printer.Width / 50
Printer.Print ("New5 Line")
Printer.CurrentX = Printer.Width / 50
Printer.Print ("New6 Line")

Printer.EndDoc
End Sub
