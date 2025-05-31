VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form purchase 
   Caption         =   "Raw Purchase/Sales"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   6990
   LinkTopic       =   "Form8"
   MDIChild        =   -1  'True
   ScaleHeight     =   3030
   ScaleWidth      =   6990
   WindowState     =   2  'Maximized
   Begin VB.ComboBox content3 
      Height          =   315
      ItemData        =   "Form8.frx":0000
      Left            =   960
      List            =   "Form8.frx":0007
      TabIndex        =   44
      Text            =   "Gold"
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox smaking 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5760
      TabIndex        =   41
      Text            =   "0"
      Top             =   5280
      Width           =   2175
   End
   Begin VB.TextBox stotamt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5760
      TabIndex        =   40
      Text            =   "0"
      Top             =   5880
      Width           =   2175
   End
   Begin VB.TextBox amt3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Left            =   5760
      TabIndex        =   38
      Text            =   "0"
      Top             =   4320
      Width           =   2175
   End
   Begin VB.TextBox amt2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Left            =   5760
      TabIndex        =   37
      Text            =   "0"
      Top             =   3840
      Width           =   2175
   End
   Begin VB.TextBox amt1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Left            =   5760
      TabIndex        =   36
      Text            =   "0"
      Top             =   3360
      Width           =   2175
   End
   Begin VB.TextBox minrate1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4080
      TabIndex        =   35
      Text            =   "0"
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox minrate2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4080
      TabIndex        =   34
      Text            =   "0"
      Top             =   3840
      Width           =   1335
   End
   Begin VB.ComboBox content2 
      Height          =   315
      ItemData        =   "Form8.frx":0011
      Left            =   960
      List            =   "Form8.frx":0054
      TabIndex        =   31
      Top             =   3840
      Width           =   1215
   End
   Begin VB.ComboBox content1 
      Height          =   315
      ItemData        =   "Form8.frx":00D3
      Left            =   960
      List            =   "Form8.frx":0116
      TabIndex        =   30
      Text            =   "Dia."
      Top             =   3360
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
      Height          =   285
      Left            =   2400
      TabIndex        =   29
      Text            =   "0"
      Top             =   3360
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
      Height          =   285
      Left            =   2400
      TabIndex        =   28
      Text            =   "0"
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox minrate3 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   4080
      TabIndex        =   27
      Text            =   "0"
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox weight3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2400
      TabIndex        =   26
      Text            =   "0"
      Top             =   4320
      Width           =   1215
   End
   Begin VB.ComboBox chtype 
      Height          =   315
      ItemData        =   "Form8.frx":0195
      Left            =   6840
      List            =   "Form8.frx":019F
      TabIndex        =   20
      Text            =   "Purchase"
      Top             =   1320
      Width           =   2535
   End
   Begin VB.ComboBox Code 
      Height          =   315
      Left            =   1800
      TabIndex        =   19
      Text            =   "Combo1"
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox tinno 
      Height          =   285
      Left            =   6840
      TabIndex        =   18
      Top             =   1800
      Width           =   2535
   End
   Begin VB.ComboBox city 
      Height          =   315
      Left            =   6840
      TabIndex        =   16
      Top             =   2280
      Width           =   2535
   End
   Begin VB.ComboBox pname 
      Height          =   315
      ItemData        =   "Form8.frx":01B3
      Left            =   1800
      List            =   "Form8.frx":01BA
      TabIndex        =   15
      Top             =   1800
      Width           =   3855
   End
   Begin VB.TextBox address 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      TabIndex        =   12
      Top             =   2280
      Width           =   3855
   End
   Begin VB.CommandButton clear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton Addnew 
      Caption         =   "AddNew"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton load 
      Caption         =   "L&oad"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "C&lose"
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton update 
      Caption         =   "&Update"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton save 
      Caption         =   "&Save"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   6720
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   1200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yy"
      Format          =   137297921
      CurrentDate     =   37684
   End
   Begin VB.Label Label15 
      Caption         =   "Making Charges:"
      Height          =   375
      Left            =   4080
      TabIndex        =   43
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Total Amount:"
      Height          =   375
      Left            =   4080
      TabIndex        =   42
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "3."
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Caption         =   "1."
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "2."
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Stone"
      Height          =   375
      Left            =   840
      TabIndex        =   25
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Wt.(Ct.)"
      Height          =   375
      Left            =   2400
      TabIndex        =   24
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Amount"
      Height          =   375
      Left            =   5760
      TabIndex        =   23
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "SMin. Rate"
      Height          =   375
      Left            =   4080
      TabIndex        =   22
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label18 
      Caption         =   "S.NO."
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   2880
      Width           =   495
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   11880
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   11880
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Label Label11 
      Caption         =   "Tin No."
      Height          =   255
      Left            =   6000
      TabIndex        =   17
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "City"
      Height          =   255
      Left            =   6000
      TabIndex        =   14
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "Address"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "DAKSH"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   11
      Top             =   0
      Width           =   8895
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   12000
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label13 
      Caption         =   "Party Name"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   12000
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label Label3 
      Caption         =   "Date:"
      Height          =   255
      Left            =   3480
      TabIndex        =   9
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "Voucher . No.:"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Type"
      Height          =   255
      Left            =   5880
      TabIndex        =   7
      Top             =   1320
      Width           =   735
   End
End
Attribute VB_Name = "purchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con5 As New ADODB.Connection
Dim res1 As New ADODB.Recordset
Dim res As New ADODB.Recordset

Private Sub Addnew_Click()
'res1.Open "select max(nvoucherno) from tblissuerecv ", MDIForm1.con1, adOpenDynamic, adLockOptimistic
'If (IsNull(res1.Fields(0))) Then
'Code.Text = 1
'Else
'Code = res1.Fields(0) + 1
'End If
'res1.Close
clearing
load_item1
save.Enabled = True
update.Enabled = False
load.Enabled = False
''print1.Enabled = False
End Sub

Private Sub clearing()
Code = ""
pname.Text = ""
tinno.Text = ""
content1 = ""
content2 = ""
'content4 = ""
city = ""
address = ""
weight1 = 0
weight2 = 0
weight3 = 0
minrate1 = 0
minrate2 = 0
minrate3 = 0
smaking = 0
stotamt = 0
amt1 = 0
amt2 = 0
amt3 = 0

'gpur = "24K"
update.Enabled = False
save.Enabled = False
load.Enabled = True
End Sub
Private Sub Clear_Click()
clearing
End Sub
Private Sub Command5_Click()
Unload Me
End Sub
Private Sub Form_Load()
'con5.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\lahar\vb\db2.mdb;Persist Security Info=False"
'con5.Open
DTPicker1.Format = dtpCustom
DTPicker1.Value = date
load_item
save.Enabled = False
update.Enabled = False
'Option1(1).Value = True
End Sub
Private Sub load_item()
Code.Clear
pname.Clear
res1.Open "select distinct(chvoucherno) from tblissuerecv order by chvoucherno", MDIForm1.con1, adOpenDynamic, adLockOptimistic
While (res1.EOF = False)
Code.AddItem (res1!chvoucherno)
res1.MoveNext
Wend
res1.close

res1.Open "select distinct(chpartyname) from tblissuerecv order by chpartyname", MDIForm1.con1, adOpenDynamic, adLockOptimistic
While (res1.EOF = False)
pname.AddItem (res1!chpartyname)
res1.MoveNext
Wend
res1.close
End Sub
Private Sub load_item1()
Code.Clear
pname.Clear

res1.Open "select distinct(chpartyname) from tblissuerecv order by chpartyname", MDIForm1.con1, adOpenDynamic, adLockOptimistic
While (res1.EOF = False)
pname.AddItem (res1!chpartyname)
res1.MoveNext
Wend
res1.close
End Sub


Private Sub Form_Unload(Cancel As Integer)
'con5.close
End Sub
Private Sub loading()
Code = res!chvoucherno
DTPicker1.Value = res!ddate
If (IsNull(res!chcontent1)) Then
content1 = ""
Else
content1 = res!chcontent1
End If

If (IsNull(res!chcontent2)) Then
content2 = ""
Else
content2 = res!chcontent2
End If

If (IsNull(res!chcontent3)) Then
content3 = ""
Else
content3 = res!chcontent3
End If

weight1 = res!nweight1
weight2 = res!nweight2
weight3 = res!nweight3


minrate1 = res!nrate1
minrate2 = res!nrate2
minrate3 = res!nrate3
smaking = res!nmaking

pname.Text = res!chpartyname

If (IsNull(res!chcity)) Then
city = ""
Else
city = res!chcity
End If

If (IsNull(res!chaddress)) Then
address = ""
Else
address = res!chaddress
End If



If (IsNull(res!chtinno)) Then
tinno = ""
Else
tinno = res!chtinno
End If

If (IsNull(res!chissuetype)) Then
chtype = ""
Else
chtype = res!chissuetype
End If
End Sub





Private Sub load_Click()
pname.Clear

str1 = "select * from tblissuerecv where chvoucherno=" & "'" & Code.Text & "'"
res.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
If (res.EOF = True) Then
 MsgBox ("Record Not Found")
res.close
Else
loading
res.close
save.Enabled = False
update.Enabled = True
End If
End Sub

Private Sub smaking_Change()
    stotamt = Val(amt1) + Val(amt2) + Val(amt3) + Val(smaking)
End Sub

Private Sub minrate1_Change()
 amt1 = Val(minrate1) * Val(weight1)
 stotamt = Val(minrate1) * Val(weight1) + Val(minrate2) * Val(weight2) + Val(minrate3) * Val(weight3) + Val(smaking)
End Sub

Private Sub minrate1_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
        content2.SetFocus
End If
End Sub

Private Sub minrate2_Change()
amt2 = Val(minrate2) * Val(weight2)
 stotamt = Val(minrate1) * Val(weight1) + Val(minrate2) * Val(weight2) + Val(minrate3) * Val(weight3) + Val(smaking)
End Sub


Private Sub minrate3_Change()
 amt3 = Val(minrate3) * Val(weight3)
 stotamt = Val(minrate1) * Val(weight1) + Val(minrate2) * Val(weight2) + Val(minrate3) * Val(weight3) + Val(smaking)
End Sub

Private Sub minrate3_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
        smaking.SetFocus
End If
End Sub
Private Sub weight1_Change()
amt1 = Val(minrate1) * Val(weight1)
stotamt = Val(minrate1) * Val(weight1) + Val(minrate2) * Val(weight2) + Val(minrate3) * Val(weight3) + Val(smaking)
End Sub


Private Sub weight1_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
        minrate1.SetFocus
End If

End Sub

Private Sub weight2_Change()
amt2 = Val(rate2) * Val(weight2)
stotamt = Val(minrate1) * Val(weight1) + Val(minrate2) * Val(weight2) + Val(minrate3) * Val(weight3) + Val(smaking)
End Sub

Private Sub weight2_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
        minrate2.SetFocus
End If

End Sub

Private Sub weight3_Change()
amt3 = Val(rate3) * Val(weight4)
stotamt = Val(minrate1) * Val(weight1) + Val(minrate2) * Val(weight2) + Val(minrate3) * Val(weight3) + Val(smaking)
End Sub

Private Sub weight3_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
        minrate3.SetFocus
End If

End Sub
Private Sub save_Click()
If (pname.Text = "" Or Code.Text = "") Then
        MsgBox ("Enter The Correct Entry For Party Name For Purchase Entry")
        Exit Sub
ElseIf ((content1.Text <> "" And Val(weight1) <= 0) Or (content1.Text = "" And Val(weight1) > 0)) Then
        MsgBox ("Enter the Valid Entry For Stone 1")
        Exit Sub
ElseIf ((content2.Text <> "" And Val(weight2) <= 0) Or (content2.Text = "" And Val(weight2) > 0)) Then
        MsgBox ("Enter the Valid Entry For Stone 2")
        Exit Sub
End If

load.Enabled = True
update.Enabled = True

save.Enabled = False
res.Open "tblissuerecv", MDIForm1.con1, adOpenDynamic, adLockOptimistic
res.Addnew
res!chvoucherno = Code.Text
res!ddate = CDate(DTPicker1)
res!chpartyname = pname.Text
res!chaddress = address.Text
res!chcity = city.Text
res!chtinno = tinno.Text
res!chissuetype = chtype.Text
res!chcontent1 = content1
res!nweight1 = Val(weight1)
res!nrate1 = Val(minrate1)
If (content2.Text <> "") Then
res!chcontent2 = content2
res!nweight2 = Val(weight2)
res!nrate2 = Val(minrate2)
End If
If (content3.Text <> "") Then
res!chcontent3 = content3
res!nweight3 = Val(weight3)
res!nrate3 = Val(minrate3)
End If
res!nmaking = Val(smaking.Text)

res.update
res.close
MsgBox ("Record Saved")
clearing
load_item
save.Enabled = False
End Sub

Private Sub Text1_Change()

End Sub

Private Sub update_Click()
If ((content2.Text <> "" And Val(weight2) <= 0) Or (content2.Text = "" And Val(weight2) > 0)) Then
        MsgBox ("Enter the Valid Entry For Stone 2")
        Exit Sub
End If

str1 = "select * from tblissuerecv where chvoucherno=" & "'" & Code.Text & "'"

res.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
res!chvoucherno = Code.Text
res!ddate = CDate(DTPicker1)
res!chpartyname = pname.Text
res!chaddress = address.Text
res!chcity = city.Text
res!chtinno = tinno.Text
res!chissuetype = chtype.Text

If (content1.Text <> "") Then
res!chcontent1 = content1
res!nweight1 = Val(weight1)
res!nrate1 = Val(minrate1)
End If

If (content2.Text <> "") Then
'res.Addnew
res!chcontent2 = content2
res!nweight2 = Val(weight2)
res!nrate2 = Val(minrate2)
End If

If (content3.Text <> "") Then
res!chcontent3 = content3
res!nweight3 = Val(weight3)
res!nrate3 = Val(minrate3)
End If
res.update
res.close
MsgBox ("Record Updated")
clearing
load_item
update.Enabled = False
save.Enabled = False
End Sub
