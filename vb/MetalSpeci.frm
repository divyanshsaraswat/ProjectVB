VERSION 5.00
Begin VB.Form metalspeci 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Metal & Dollar Specification."
   ClientHeight    =   2625
   ClientLeft      =   3690
   ClientTop       =   2445
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox drt 
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Text            =   "0"
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox mrt 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Text            =   "0"
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Close"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&Save"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "USD Rate."
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Metal Rate."
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "metalspeci"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim con3 As New ADODB.Connection
Dim res1 As New ADODB.Recordset

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub Form_Load()
'con3.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\lahar\vb\db2.mdb;Persist Security Info=False"
'con3.Open
Dim str1 As String
str1 = " SELECT nmetalrt,ndolarrt From tblmdspeci WHERE anum =(select max(anum) from tblmdspeci )"
res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
If (IsNull(res1.Fields(0)) Or res1.BOF = True) Then
mrt.Text = 0
Else
mrt = res1.Fields(0)
End If
If (IsNull(res1.Fields(1)) Or res1.BOF = True) Then
drt.Text = 0
Else
drt = res1.Fields(1)
End If
res1.Close
End Sub
Private Sub Form_Unload(Cancel As Integer)
'con3.close
End Sub
Private Sub OKButton_Click()
If (Val(mrt) = 0 Or Val(drt) = 0) Then
 MsgBox ("Enter the Valid Value In The Fields")
Else
res1.Open "tblmdspeci", MDIForm1.con1, adOpenDynamic, adLockOptimistic
res1.AddNew
res1!nmetalrt = Val(mrt)
res1!ndolarrt = Val(drt)
res1!ddate = CDate(Date)
res1.Update
res1.Close
MsgBox ("Record Saved")
End If
End Sub
