VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "ESTSALES"
   ClientHeight    =   6840
   ClientLeft      =   1020
   ClientTop       =   2055
   ClientWidth     =   10095
   DrawMode        =   16  'Merge Pen
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6840
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox gpw 
      Appearance      =   0  'Flat
      DataField       =   "igw"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   7080
      TabIndex        =   60
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox totamt 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   6360
      TabIndex        =   47
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox makingchrages 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   6360
      TabIndex        =   46
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton close 
      Caption         =   "C&lose"
      Height          =   375
      Left            =   3000
      TabIndex        =   45
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton save 
      Caption         =   "&Save"
      Height          =   375
      Left            =   120
      TabIndex        =   44
      Top             =   7800
      Width           =   1215
   End
   Begin VB.ComboBox code 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "form2.frx":0000
      Left            =   1440
      List            =   "form2.frx":0002
      TabIndex        =   43
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox mrate4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6360
      TabIndex        =   15
      Text            =   "0"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox mrate3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6360
      TabIndex        =   12
      Text            =   "0"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox mrate2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6360
      TabIndex        =   9
      Text            =   "0"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox mrate1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6360
      TabIndex        =   6
      Text            =   "0"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Itemname1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   1680
      Width           =   2775
   End
   Begin VB.TextBox Invoicedate 
      Appearance      =   0  'Flat
      DataField       =   "invoicedate"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   4200
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox igw 
      Appearance      =   0  'Flat
      DataField       =   "igw"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   7080
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox content1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox mcharges 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   4800
      TabIndex        =   17
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox amt 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   4800
      TabIndex        =   19
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox amt4 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   4920
      TabIndex        =   28
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox amt3 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   4920
      TabIndex        =   27
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox amt2 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   4920
      TabIndex        =   26
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox amt1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   4920
      TabIndex        =   25
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox rate4 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   3480
      TabIndex        =   24
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox rate3 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   3480
      TabIndex        =   23
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox rate2 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   3480
      TabIndex        =   22
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox rate1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   3480
      TabIndex        =   21
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox weight4 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      TabIndex        =   14
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox weight3 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      TabIndex        =   11
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox weight2 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      TabIndex        =   8
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox weight1 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox content4 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   13
      Text            =   " "
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox content2 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox content3 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox maker 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   960
      TabIndex        =   16
      Top             =   4440
      Width           =   855
   End
   Begin VB.TextBox issueto 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   960
      TabIndex        =   18
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton print 
      Caption         =   "&Print"
      Height          =   375
      Left            =   1560
      TabIndex        =   20
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton clear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   7800
      Width           =   1215
   End
   Begin VB.Label Label26 
      BackColor       =   &H8000000E&
      Caption         =   "Balance"
      Height          =   375
      Left            =   9960
      TabIndex        =   59
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   12000
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Label Label25 
      BackColor       =   &H8000000E&
      Caption         =   "Date:"
      Height          =   255
      Left            =   120
      TabIndex        =   58
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label24 
      BackColor       =   &H8000000E&
      Caption         =   "Due Date:"
      Height          =   255
      Left            =   3240
      TabIndex        =   57
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label23 
      BackColor       =   &H8000000E&
      Caption         =   "Party Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   56
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label Label22 
      BackColor       =   &H8000000E&
      Caption         =   "Through:"
      Height          =   255
      Left            =   3240
      TabIndex        =   55
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label Label21 
      BackColor       =   &H8000000E&
      Caption         =   "Rec3"
      Height          =   375
      Left            =   8400
      TabIndex        =   54
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label20 
      BackColor       =   &H8000000E&
      Caption         =   "Sales Amount"
      Height          =   375
      Left            =   120
      TabIndex        =   53
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label19 
      BackColor       =   &H8000000E&
      Caption         =   "Comm@"
      Height          =   375
      Left            =   1800
      TabIndex        =   52
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label18 
      BackColor       =   &H8000000E&
      Caption         =   "Amount"
      Height          =   375
      Left            =   3480
      TabIndex        =   51
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label17 
      BackColor       =   &H8000000E&
      Caption         =   "Rec1"
      Height          =   375
      Left            =   5040
      TabIndex        =   50
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label16 
      BackColor       =   &H8000000E&
      Caption         =   "Rec2"
      Height          =   375
      Left            =   6720
      TabIndex        =   49
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   12000
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Lahar Exports"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   48
      Top             =   120
      Width           =   9375
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Min.Rate"
      Height          =   255
      Left            =   6240
      TabIndex        =   42
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Item Code:"
      Height          =   255
      Left            =   120
      TabIndex        =   41
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Item Name"
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Caption         =   "Invoice Date"
      Height          =   255
      Left            =   2760
      TabIndex        =   39
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "IGW"
      Height          =   255
      Left            =   5520
      TabIndex        =   38
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "Gold Wes.@(%)"
      Height          =   255
      Left            =   5400
      TabIndex        =   37
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   12000
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000E&
      Caption         =   "Content"
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Weight"
      Height          =   255
      Left            =   1920
      TabIndex        =   35
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Std.Rate"
      Height          =   255
      Left            =   3360
      TabIndex        =   34
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Amount"
      Height          =   255
      Left            =   4800
      TabIndex        =   33
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   12000
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label Label10 
      BackColor       =   &H8000000E&
      Caption         =   "Making Charges:"
      Height          =   375
      Left            =   3240
      TabIndex        =   32
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label11 
      BackColor       =   &H8000000E&
      Caption         =   "Total Amount:"
      Height          =   255
      Left            =   3240
      TabIndex        =   31
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label Label12 
      BackColor       =   &H8000000E&
      Caption         =   "S.No:"
      Height          =   375
      Left            =   120
      TabIndex        =   30
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label Label13 
      BackColor       =   &H8000000E&
      Caption         =   "Price No.:"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   5040
      Width           =   735
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   12000
      Y1              =   5400
      Y2              =   5400
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim res As New ADODB.Recordset


Private Sub clear_Click()
Form2.code = ""
Form2.amt1 = ""
Form2.amt2 = ""
Form2.amt3 = ""
Form2.amt4 = ""
Form2.content1 = ""
Form2.content2 = ""
Form2.content3 = ""
Form2.content4 = ""
Form2.igw = ""
Form2.gpw = ""
Form2.Invoicedate = ""
Form2.issueto = ""
Form2.Itemname1 = ""
Form2.maker = ""
Form2.makingchrages = ""
Form2.mrate1 = ""
Form2.mrate2 = ""
Form2.mrate3 = ""
Form2.mrate4 = ""



End Sub

Private Sub close_Click()
Unload Me
con.close
End Sub

Private Sub code_LostFocus()
str1 = "select * from tblcosting where ncode=" & Val(code.Text)
res.Open str1, con
loading
res.close
End Sub

Private Sub Form_Load()
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\lahar\vb\db2.mdb;Persist Security Info=False"
con.Open
res.Open "select ncode from tblcosting", con
While (res.EOF() = False)
code.AddItem (res.Fields(0))
res.MoveNext
Wend
res.close

End Sub

Private Sub loading()

Itemname1.Text = res!chItemname
Invoicedate = res!dInvoicedate
content1 = res!chcontent1
content2 = res!chcontent2
strcontent3 = res!chcontent3
If (strcontent3 = Empty) Then
content3.Text = ""
Else
content3.Text = strcontent3
End If

If (res!chcontent4 = Empty) Then
content4.Text = ""
Else
content4.Text = res!chcontent4
End If

igw = res!nigw
makingcharges = res!nmaking
If (res!chmaker = Empty) Then
maker.Text = ""
Else
maker = res!chmaker
End If
Invoicedate = res!drecdate
issueto = res!chissueto
gpw = res!ngpw

End Sub

Private Sub mrate1_Change()
totamt = (Val(mrate1) * Val(weight1)) + (Val(mrate2) * Val(weight2)) + (Val(mrate3) * Val(weight3)) + (Val(mrate4) * Val(weight4)) + Val(makingcharges)
End Sub

Private Sub mrate2_Change()
totamt = (Val(mrate1) * Val(weight1)) + (Val(mrate2) * Val(weight2)) + (Val(mrate3) * Val(weight3)) + (Val(mrate4) * Val(weight4)) + Val(makingcharges)
End Sub

Private Sub mrate3_Change()
totamt = (Val(mrate1) * Val(weight1)) + (Val(mrate2) * Val(weight2)) + (Val(mrate3) * Val(weight3)) + (Val(mrate4) * Val(weight4)) + Val(makingcharges)
End Sub

Private Sub mrate4_Change()
totamt = (Val(mrate1) * Val(weight1)) + (Val(mrate2) * Val(weight2)) + (Val(mrate3) * Val(weight3)) + (Val(mrate4) * Val(weight4)) + Val(makingcharges)
End Sub

Private Sub save_Click()
Dim con1 As New ADODB.Connection
con1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\lahar\vb\db2.mdb;Persist Security Info=False"
con1.Open
res.Open "tblestsalevalue", con1, adOpenDynamic, adLockOptimistic
res.Addnew
'////////////
res!ncode = Val(code.Text)
res!chInvoicedate = CDate(Invoicedate)
res!nweight1 = Val(weight1)
res!nweight2 = Val(weight2)
res!nweight3 = Val(weight3)
res!nweight4 = Val(weight4)
res!nsrate1 = Val(rate1)
res!nsrate2 = Val(rate2)
res!nsrate3 = Val(rate3)
res!nsrate4 = Val(rate4)
res!nminrate1 = Val(mrate1)
res!nminrate2 = Val(mrate2)
res!nminrate3 = Val(mrate3)
res!nminrate4 = Val(mrate4)
res.Update

res.close

End Sub

Private Sub Text2_Change()

End Sub


