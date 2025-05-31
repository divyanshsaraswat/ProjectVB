VERSION 5.00
Begin VB.Form import 
   Caption         =   "Import Data"
   ClientHeight    =   2055
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   3675
   LinkTopic       =   "Form9"
   ScaleHeight     =   2055
   ScaleWidth      =   3675
   Begin VB.CommandButton close 
      Caption         =   "&Close"
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton import 
      Caption         =   "&Import Data"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Item No."
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
End
Attribute VB_Name = "import"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con1 As New ADODB.Connection
Dim res1 As New ADODB.Recordset
Dim res As New ADODB.Recordset
Private Sub close_Click()
Unload Me
End Sub
Private Sub import_Click()
'DB4 CONNECTION
res1.Open "tblcosting", MDIForm1.con3, adOpenDynamic, adLockOptimistic

While (res1.EOF = False)
'DB2 CONNECTION
Dim str1 As String
str1 = "select * from tblcosting where ncode=" & res1!ncode
res.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
'res.Open "tblcosting", MDIForm1.con1, adOpenDynamic, adLockOptimistic
If (res.EOF = True) Then
res.AddNew
End If
res!ncode = res1!ncode
Label1.Caption = "Item No. " & res1!ncode

res!chItemname = res1!chItemname
res!chowner = res1!chowner
res!dinvoicedate = res1!dinvoicedate
res!chcontent1 = res1!chcontent1

If (IsNull(res1!chcontent2)) Then
res!chcontent2 = ""
Else
res!chcontent2 = res1!chcontent2
End If

If (IsNull(res1!chcontent3)) Then
res!chcontent3 = ""
Else
res!chcontent3 = res1!chcontent3
End If

If (IsNull(res1!chcontent4)) Then
res!chcontent4 = ""
Else
res!chcontent4 = res1!chcontent4
End If

res!nweight1 = res1!nweight1
res!nweight2 = res1!nweight2
res!nweight3 = res1!nweight3
res!nweight4 = res1!nweight4

res!nrate1 = res1!nrate1
res!nrate2 = res1!nrate2
res!nrate3 = res1!nrate3
res!nrate4 = res1!nrate4
res!nigw = res1!nigw

res!nmaking = res1!nmaking
res!chmaker = res1!chmaker
'res!drecdate = CDate(Invoicedate)

res!drecdate = res1!drecdate
res!chissueto = res1!chissueto

res!ngpw = res1!ngpw
res!pcs1 = res1!pcs1
res!pcs2 = res1!pcs2
res!pcs3 = res1!pcs3
res!gpur = res1!gpur
res!minrate1 = res1!minrate1
res!minrate2 = res1!minrate2
res!minrate3 = res1!minrate3
res!minrate4 = res1!minrate4
res!nmaking1 = res1!nmaking1
res!chquality1 = res1!chquality1
res!chquality2 = res1!chquality2
res!chquality3 = res1!chquality3
res!chcolor1 = res1!chcolor1
res!chcolor2 = res1!chcolor2
res!chcolor3 = res1!chcolor3
res!chsize1 = res1!chsize1
res!chsize2 = res1!chsize2
res!chsize3 = res1!chsize3
res!chstate = res1!chstate

res!chcategory = res1!chcategory

If (IsNull(res1!chpcode)) Then
res!chpcode = ""
Else
res!chpcode = res1!chpcode
End If

res!ndno = res1!ndno

'If (picturepath <> Empty) Then
    res!opicture = res1!opicture
 '   picturepath = ""
'End If

res.Update
res.close
res1.Delete
res1.Update

'res1.close
res1.MoveNext
Wend
res1.close
MsgBox ("Data Imported Successfully")
End Sub
