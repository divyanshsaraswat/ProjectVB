VERSION 5.00
Begin VB.Form partydetail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Party Details"
   ClientHeight    =   5970
   ClientLeft      =   2715
   ClientTop       =   1740
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   11340
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1080
      TabIndex        =   28
      Top             =   240
      Width           =   1335
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Broker"
      Height          =   255
      Left            =   10200
      TabIndex        =   27
      Top             =   360
      Width           =   855
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Maker"
      Height          =   255
      Left            =   9000
      TabIndex        =   26
      Top             =   360
      Width           =   855
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Buyer"
      Height          =   255
      Left            =   7920
      TabIndex        =   25
      Top             =   360
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Saler"
      Height          =   255
      Left            =   6960
      TabIndex        =   24
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton clear 
      Caption         =   "C&lear"
      Height          =   375
      Left            =   5160
      TabIndex        =   23
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton update 
      Caption         =   "&Update"
      Height          =   375
      Left            =   3840
      TabIndex        =   22
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton load 
      Caption         =   "&Search"
      Height          =   375
      Left            =   720
      TabIndex        =   21
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox phone2 
      Height          =   375
      Left            =   5520
      TabIndex        =   20
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox phone1 
      Height          =   375
      Left            =   5520
      TabIndex        =   18
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox country 
      Height          =   375
      Left            =   1800
      TabIndex        =   16
      Top             =   3960
      Width           =   2055
   End
   Begin VB.TextBox state 
      Height          =   375
      Left            =   1800
      TabIndex        =   13
      Top             =   3360
      Width           =   2055
   End
   Begin VB.TextBox city 
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox street 
      Height          =   855
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Text            =   "partydetail.frx":0000
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox title 
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox cpname 
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox pname 
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6480
      TabIndex        =   1
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&Add New Contact"
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Label Label11 
      Caption         =   "Party Code:"
      Height          =   375
      Left            =   0
      TabIndex        =   29
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "Phone2"
      Height          =   375
      Left            =   4080
      TabIndex        =   19
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Phone1"
      Height          =   375
      Left            =   4080
      TabIndex        =   17
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Country:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "State:"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "City:"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Street Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Job Title:"
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Company Name:"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   11880
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label2 
      Caption         =   "Type:"
      Height          =   255
      Left            =   6120
      TabIndex        =   4
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Person Name:"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "partydetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim con4 As New ADODB.Connection
Dim res1 As New ADODB.Recordset
Dim res As New ADODB.Recordset
Dim partyname As String
Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub Clear_Click()
clearing
OKButton.Enabled = True
Update.Enabled = False
End Sub

Private Sub Form_Load()
'con4.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\lahar\vb\db2.mdb;Persist Security Info=False"
'con4.Open
End Sub
Private Sub load_Click()
partyname = UCase(pname.Text)
Dim str1 As String
str1 = "select * from tblpartydetails where chname = " & "'" & UCase(pname.Text) & "'"
res.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
If (res.EOF = True) Then
 MsgBox ("Record Not Found")
res.CLOSE
Else
loading
res.CLOSE
OKButton.Enabled = False
Update.Enabled = True
End If
End Sub
Private Sub loading()
pname = res!chname
cpname = res!chcompanyname
If (res!ntype = 1) Then
'Option1(0).Value = True
ElseIf (res!ntype = 2) Then
'Option1(1).Value = True
Else
'Option1(1).Value = True
End If
street = res!chaddress
city = res!chcity
state = res!chstate
country = res!chcountry
title = res!chtitle
phone1 = res!nphone1
phone2 = res!nphone2
End Sub

Private Sub OKButton_Click()
res.Open "tblpartydetails", MDIForm1.con1, adOpenDynamic, adLockOptimistic
res.addnew
If (IsEmpty(pname.Text)) Then
MsgBox ("Name Cannot Be Empty")
res.CLOSE
Exit Sub
Else
res!chname = UCase(pname.Text)
End If
res!chcompanyname = cpname
'If (Option1(0).Value = True) Then
'res!ntype = 1
'ElseIf (Option1(1).Value = True) Then
'res!ntype = 2
'Else
'res!ntype = 3
'End If
res!chaddress = street
res!chcity = city
res!chstate = state
res!chcountry = country
res!chtitle = title
res!nphone1 = Val(phone1)
res!nphone2 = Val(phone2)
res.Update
res.CLOSE
MsgBox ("Record Saved")
clearing
End Sub
Private Sub clearing()
partyname = ""
pname = ""
cpname = ""
'Option1(2).Value = True
street = ""
city = ""
state = ""
country = ""
title = ""
phone1 = ""
phone2 = ""

End Sub
Private Sub update_Click()
Dim str1 As String
str1 = "select * from tblpartydetails where chname = " & "'" & partyname & "'"
res.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
If (IsEmpty(pname.Text)) Then
   MsgBox ("Name Cannot Be Empty")
res.CLOSE
Exit Sub
Else
res!chname = pname.Text
End If

res!chcompanyname = cpname
'If (Option1(0).Value = True) Then
'res!ntype = 1
'ElseIf (Option1(1).Value = True) Then
'res!ntype = 2
'Else
'res!ntype = 3
'End If
res!chaddress = street
res!chcity = city
res!chstate = state
res!chcountry = country
res!chtitle = title
res!nphone1 = Val(phone1)
res!nphone2 = Val(phone2)
res.Update
res.CLOSE
MsgBox ("Record Updated")
clearing
End Sub

