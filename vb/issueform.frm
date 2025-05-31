VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form issueform 
   Caption         =   "Form6"
   ClientHeight    =   6480
   ClientLeft      =   945
   ClientTop       =   180
   ClientWidth     =   9690
   LinkTopic       =   "Form6"
   MDIChild        =   -1  'True
   ScaleHeight     =   6480
   ScaleWidth      =   9690
   WindowState     =   2  'Maximized
   Begin VB.TextBox wt1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3960
      TabIndex        =   55
      Text            =   "0"
      Top             =   4800
      Width           =   615
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Recv."
      Height          =   375
      Index           =   1
      Left            =   9960
      TabIndex        =   54
      Top             =   1200
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Issue"
      Height          =   375
      Index           =   0
      Left            =   8760
      TabIndex        =   53
      Top             =   1200
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.TextBox Code 
      Appearance      =   0  'Flat
      DataField       =   "ncode"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1800
      TabIndex        =   51
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton save 
      Caption         =   "&Save"
      Height          =   375
      Left            =   1560
      TabIndex        =   49
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton update 
      Caption         =   "&Update"
      Height          =   375
      Left            =   3840
      TabIndex        =   48
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "C&lose"
      Height          =   375
      Left            =   7200
      TabIndex        =   47
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton load 
      Caption         =   "L&oad"
      Height          =   375
      Left            =   2760
      TabIndex        =   46
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton Addnew 
      Caption         =   "AddNew"
      Height          =   375
      Left            =   480
      TabIndex        =   45
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton clear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   6120
      TabIndex        =   44
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton print1 
      Caption         =   "&Print"
      Height          =   375
      Left            =   4920
      TabIndex        =   43
      Top             =   5640
      Width           =   975
   End
   Begin VB.ComboBox gpur 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "issueform.frx":0000
      Left            =   2160
      List            =   "issueform.frx":0023
      TabIndex        =   42
      Text            =   "24K"
      Top             =   4800
      Width           =   735
   End
   Begin VB.ComboBox content4 
      Height          =   315
      ItemData        =   "issueform.frx":004E
      Left            =   480
      List            =   "issueform.frx":0091
      TabIndex        =   38
      Top             =   3960
      Width           =   735
   End
   Begin VB.TextBox qul3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3000
      TabIndex        =   37
      Top             =   3960
      Width           =   735
   End
   Begin VB.TextBox size3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      TabIndex        =   36
      Text            =   "0"
      Top             =   3960
      Width           =   735
   End
   Begin VB.ComboBox color3 
      Height          =   315
      ItemData        =   "issueform.frx":0110
      Left            =   1320
      List            =   "issueform.frx":011D
      TabIndex        =   35
      Top             =   3960
      Width           =   735
   End
   Begin VB.TextBox pcs4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3960
      TabIndex        =   34
      Text            =   "0"
      Top             =   3960
      Width           =   615
   End
   Begin VB.TextBox weight4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4680
      TabIndex        =   33
      Text            =   "0"
      Top             =   3960
      Width           =   735
   End
   Begin VB.TextBox weight5 
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
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "0"
      Top             =   4800
      Width           =   735
   End
   Begin VB.TextBox weight3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4680
      TabIndex        =   20
      Text            =   "0"
      Top             =   3480
      Width           =   735
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
      Left            =   4680
      TabIndex        =   19
      Text            =   "0"
      Top             =   3000
      Width           =   735
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
      Left            =   4680
      TabIndex        =   18
      Text            =   "0"
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox issueto 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3720
      TabIndex        =   17
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox pcs1 
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
      Left            =   3960
      TabIndex        =   16
      Text            =   "0"
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox pcs2 
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
      Left            =   3960
      TabIndex        =   15
      Text            =   "0"
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox pcs3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3960
      TabIndex        =   14
      Text            =   "0"
      Top             =   3480
      Width           =   615
   End
   Begin VB.ComboBox content5 
      Height          =   315
      ItemData        =   "issueform.frx":012D
      Left            =   480
      List            =   "issueform.frx":0140
      TabIndex        =   13
      Text            =   "Gold"
      Top             =   4800
      Width           =   1575
   End
   Begin VB.ComboBox qul 
      Height          =   315
      ItemData        =   "issueform.frx":016F
      Left            =   3000
      List            =   "issueform.frx":0194
      TabIndex        =   12
      Top             =   2520
      Width           =   855
   End
   Begin VB.ComboBox color 
      Height          =   315
      ItemData        =   "issueform.frx":01C5
      Left            =   1320
      List            =   "issueform.frx":01E1
      TabIndex        =   11
      Top             =   2520
      Width           =   735
   End
   Begin VB.ComboBox color1 
      Height          =   315
      ItemData        =   "issueform.frx":020E
      Left            =   1320
      List            =   "issueform.frx":021B
      TabIndex        =   10
      Top             =   3000
      Width           =   735
   End
   Begin VB.ComboBox color2 
      Height          =   315
      ItemData        =   "issueform.frx":022B
      Left            =   1320
      List            =   "issueform.frx":0238
      TabIndex        =   9
      Top             =   3480
      Width           =   735
   End
   Begin VB.ComboBox size 
      Height          =   315
      ItemData        =   "issueform.frx":0248
      Left            =   2160
      List            =   "issueform.frx":0270
      TabIndex        =   8
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox size2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      TabIndex        =   7
      Text            =   "0"
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox size1 
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
      Left            =   2160
      TabIndex        =   6
      Text            =   "0"
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox qul2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3000
      TabIndex        =   5
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox qul1 
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
      Left            =   3000
      TabIndex        =   4
      Top             =   3000
      Width           =   735
   End
   Begin VB.ComboBox content1 
      Height          =   315
      ItemData        =   "issueform.frx":02B3
      Left            =   480
      List            =   "issueform.frx":02F6
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Dia."
      Top             =   2520
      Width           =   735
   End
   Begin VB.ComboBox content2 
      Height          =   315
      ItemData        =   "issueform.frx":0375
      Left            =   480
      List            =   "issueform.frx":03B8
      TabIndex        =   2
      Top             =   3000
      Width           =   735
   End
   Begin VB.ComboBox content3 
      Height          =   315
      ItemData        =   "issueform.frx":0437
      Left            =   480
      List            =   "issueform.frx":047A
      TabIndex        =   1
      Top             =   3480
      Width           =   735
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   6480
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yy"
      Format          =   137297921
      CurrentDate     =   37684
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Type"
      Height          =   255
      Left            =   7800
      TabIndex        =   52
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Issue/Recv. No.:"
      Height          =   375
      Left            =   120
      TabIndex        =   50
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "4."
      Height          =   255
      Left            =   0
      TabIndex        =   41
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "3."
      Height          =   255
      Left            =   0
      TabIndex        =   40
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "2."
      Height          =   255
      Left            =   0
      TabIndex        =   39
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      Caption         =   "Sz."
      Height          =   375
      Left            =   2280
      TabIndex        =   32
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Date:"
      Height          =   255
      Left            =   5400
      TabIndex        =   31
      Top             =   1200
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   12000
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Stone"
      Height          =   375
      Left            =   600
      TabIndex        =   30
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Wt.(Ct.)"
      Height          =   375
      Left            =   4680
      TabIndex        =   29
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   11880
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label Label13 
      Caption         =   "Issue To:"
      Height          =   255
      Left            =   2760
      TabIndex        =   28
      Top             =   1200
      Width           =   735
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   11880
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   12000
      Y1              =   1080
      Y2              =   1080
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
      TabIndex        =   27
      Top             =   0
      Width           =   8895
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "Pc."
      Height          =   375
      Left            =   3960
      TabIndex        =   26
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Caption         =   "1."
      Height          =   255
      Left            =   0
      TabIndex        =   25
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label18 
      Caption         =   "S.NO."
      Height          =   375
      Left            =   0
      TabIndex        =   24
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Caption         =   "Qt."
      Height          =   375
      Left            =   3120
      TabIndex        =   23
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Caption         =   "Col."
      Height          =   375
      Left            =   1560
      TabIndex        =   22
      Top             =   1920
      Width           =   615
   End
End
Attribute VB_Name = "issueform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con5 As New ADODB.Connection
Dim res1 As New ADODB.Recordset
Dim res As New ADODB.Recordset

Private Sub Addnew_Click()
res1.Open "select max(nissueno) from tblissuerecv ", MDIForm1.con1, adOpenDynamic, adLockOptimistic
If (IsNull(res1.Fields(0))) Then
Code.Text = 1
Else
Code = res1.Fields(0) + 1
End If
res1.close
save.Enabled = True
update.Enabled = False
load.Enabled = False
print1.Enabled = False
End Sub

Private Sub clearing()
Code = ""
issueto = ""
Option1(0).Value = True
content2 = ""
content3 = ""
content4 = ""
color = ""
color1 = ""
color2 = ""
color3 = ""

weight1 = 0
weight2 = 0
weight3 = 0
weight4 = 0
wt1 = 0
weight5 = 0
gpur = "24K"
pcs1 = 0
pcs2 = 0
pcs3 = 0
pcs4 = 0

qul = ""
qul1 = ""
qul2 = ""
qul3 = ""
size = 0
size1 = 0
size2 = 0
size3 = 0
print1.Enabled = False
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
save.Enabled = False
update.Enabled = False
print1.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
'con5.close
End Sub
Private Sub loading()
Code = res!nissueno
DTPicker1.Value = res!ddate
content1 = res!chcontent1
content2 = res!chcontent2
content3 = res!chcontent3
content4 = res!chcontent4

weight1 = res!nweight1
weight2 = res!nweight2
weight3 = res!nweight3
weight4 = res!nweight4
weight5 = res!nweight5

issueto = res!chissueto
pcs1 = res!npcs1
pcs2 = res!npcs2
pcs3 = res!npcs3
pcs4 = res!npcs4

gpur = res!ngpur

qul = res!chquality1
qul1 = res!chquality2
qul2 = res!chquality3
If (IsNull(res!chquality4)) Then
qul3 = ""
Else
qul3 = res!chquality4
End If

color = res!chcolor1
color1 = res!chcolor2
color2 = res!chcolor3
color3 = res!chcolor4

size = res!chsize1
size1 = res!chsize2
size2 = res!chsize3
size3 = res!chsize4

If (res!nissuetype = 1) Then
Option1(0).Value = True
Else
Option1(0).Value = True
End If
End Sub


Private Sub gpur_LostFocus()
weight5 = Val(wt1) * (Val(gpur) / 24)
End Sub

Private Sub load_Click()
str1 = "select * from tblissuerecv where nissueno=" & Val(Code.Text)
res.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
If (res.EOF = True) Then
 MsgBox ("Record Not Found")
res.close
Else
loading
res.close
save.Enabled = False
update.Enabled = True
print1.Enabled = True
End If
End Sub

Private Sub save_Click()
If (issueto.Text = "" Or Code.Text = "") Then
        MsgBox ("Enter The Correct Entry For Issue/Recv No or Issuto Entry")
        Exit Sub
ElseIf ((content2.Text <> "" And Val(weight2) <= 0) Or (content2.Text = "" And Val(weight2) > 0)) Then
        MsgBox ("Enter the Valid Entry For Stone 2")
        Exit Sub
ElseIf ((content3.Text <> "" And Val(weight3) <= 0) Or (content3.Text = "" And Val(weight3) > 0)) Then
        MsgBox ("Enter the Valid Entry For Stone 3")
        Exit Sub
ElseIf ((content4.Text <> "" And Val(weight4) <= 0) Or (content4.Text = "" And Val(weight4) > 0)) Then
        MsgBox ("Enter the Valid Entry For Stone 4")
        Exit Sub
End If

load.Enabled = True
update.Enabled = True
print1.Enabled = True
save.Enabled = False
res.Open "tblissuerecv", MDIForm1.con1, adOpenDynamic, adLockOptimistic
res.Addnew
res!nissueno = Val(Code)
res!ddate = CDate(DTPicker1)
res!chcontent1 = content1
res!chcontent2 = content2
res!chcontent3 = content3
res!chcontent4 = content4

res!nweight1 = Val(weight1)
res!nweight2 = Val(weight2)
res!nweight3 = Val(weight3)
res!nweight4 = Val(weight4)
res!nweight5 = Val(weight5)

res!chissueto = issueto

res!npcs1 = Val(pcs1)
res!npcs2 = Val(pcs2)
res!npcs3 = Val(pcs3)
res!npcs4 = Val(pcs4)

res!ngpur = Val(gpur)

res!chquality1 = qul
res!chquality2 = qul1
res!chquality3 = qul2
res!chquality4 = qul3

res!chcolor1 = color
res!chcolor2 = color1
res!chcolor3 = color2
res!chcolor4 = color3

res!chsize1 = size
res!chsize2 = size1
res!chsize3 = size2
res!chsize4 = size3

If (Option1(0).Value = True) Then
res!nissuetype = 1
Else
res!nissuetype = 0
End If

res.update
res.close
MsgBox ("Record Saved")
clearing
save.Enabled = False
End Sub

Private Sub update_Click()
If (issueto.Text = "" Or Code.Text = "") Then
        MsgBox ("Enter The Correct Entry For Issue/Recv No or Issuto Entry")
        Exit Sub
ElseIf ((content2.Text <> "" And Val(weight2) <= 0) Or (content2.Text = "" And Val(weight2) > 0)) Then
        MsgBox ("Enter the Valid Entry For Stone 2")
        Exit Sub
ElseIf ((content3.Text <> "" And Val(weight3) <= 0) Or (content3.Text = "" And Val(weight3) > 0)) Then
        MsgBox ("Enter the Valid Entry For Stone 3")
        Exit Sub
ElseIf ((content4.Text <> "" And Val(weight4) <= 0) Or (content4.Text = "" And Val(weight4) > 0)) Then
        MsgBox ("Enter the Valid Entry For Stone 4")
        Exit Sub
End If

str1 = "select * from tblissuerecv where nissueno=" & Val(Code.Text)
res.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
res!nissueno = Val(Code)
res!ddate = CDate(DTPicker1)
res!chcontent1 = content1
res!chcontent2 = content2
res!chcontent3 = content3
res!chcontent4 = content4

res!nweight1 = Val(weight1)
res!nweight2 = Val(weight2)
res!nweight3 = Val(weight3)
res!nweight4 = Val(weight4)
res!nweight5 = Val(weight5)

res!chissueto = issueto

res!npcs1 = Val(pcs1)
res!npcs2 = Val(pcs2)
res!npcs3 = Val(pcs3)
res!npcs4 = Val(pcs4)

res!ngpur = Val(gpur)

res!chquality1 = qul
res!chquality2 = qul1
res!chquality3 = qul2
res!chquality4 = qul3

res!chcolor1 = color
res!chcolor2 = color1
res!chcolor3 = color2
res!chcolor4 = color3

res!chsize1 = size
res!chsize2 = size1
res!chsize3 = size2
res!chsize4 = size3

If (Option1(0).Value = True) Then
res!nissuetype = 1
Else
res!nissuetype = 0
End If
res.update
res.close
MsgBox ("Record Updated")
clearing
update.Enabled = False
save.Enabled = False
End Sub

Private Sub wt1_Change()
weight5 = Val(wt1) * (Val(gpur) / 24)
End Sub
