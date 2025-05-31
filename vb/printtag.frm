VERSION 5.00
Begin VB.Form printtag 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tag Printing"
   ClientHeight    =   4485
   ClientLeft      =   3765
   ClientTop       =   2520
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox slabel 
      Height          =   315
      ItemData        =   "printtag.frx":0000
      Left            =   4800
      List            =   "printtag.frx":001F
      TabIndex        =   12
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton rlist 
      Caption         =   "<<"
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton print 
      Caption         =   "E&xport Print data"
      Height          =   375
      Left            =   4560
      TabIndex        =   9
      Top             =   720
      Width           =   1455
   End
   Begin VB.ListBox List2 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   2595
      ItemData        =   "printtag.frx":003F
      Left            =   2640
      List            =   "printtag.frx":0041
      MultiSelect     =   2  'Extended
      TabIndex        =   8
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton addlist 
      Caption         =   ">>"
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   2160
      Width           =   735
   End
   Begin VB.ListBox List1 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   2595
      ItemData        =   "printtag.frx":0043
      Left            =   360
      List            =   "printtag.frx":0045
      MultiSelect     =   2  'Extended
      TabIndex        =   5
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton remove 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton add 
      Caption         =   "&Add"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox ncode 
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "S.Label"
      Height          =   375
      Left            =   3720
      TabIndex        =   13
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Printing List"
      Height          =   255
      Left            =   2640
      TabIndex        =   11
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Pending List"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Code:"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "printtag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con1 As New ADODB.Connection
Dim res1 As New ADODB.Recordset
Private Sub add_Click()
If (List2.ListCount = 18) Then
MsgBox ("You Can't Add More Then 18 Items In Printing List")
Exit Sub
End If
str1 = "Select ncode,chcategory from tblcosting where ncode=" & ncode.Text
res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
ncode.Text = ""
If (res1.EOF = False) Then
    If (List2.ListCount > 0) Then
        For i = 0 To List2.ListCount - 1
            pos = InStr(1, List2.List(i), "-", vbTextCompare)
            num = Val(Mid(List2.List(i), pos + 1))
            If (Val(res1!ncode) = Val(num)) Then
                MsgBox ("Item Allready Added In The List")
               res1.CLOSE
                Exit Sub
            End If
            List2.AddItem (res1!chcategory & "-" & res1!ncode)
            res1.CLOSE
            Exit Sub
        Next
    Else
        List2.AddItem (res1!chcategory & "-" & res1!ncode)
        res1.CLOSE
        Exit Sub
    End If
   Else
      MsgBox ("Item Not Available")
      res1.CLOSE
End If

End Sub

Private Sub addlist_Click()
Dim cpos As Integer
Dim ritems(18) As Integer
j = 0
For i = 0 To List1.ListCount - 1
'If (I = 10) Then
'MsgBox ("You Can Print Only Ten Tag In One Time")
'Exit Sub
'End If

If (List1.Selected(i) = True) Then
    List2.AddItem (List1.List(i))
    pos = InStr(1, List1.List(i), "-", vbTextCompare)
    num = Val(Mid(List1.List(i), pos + 1))
    ritems(j) = i
    j = j + 1
    'List1.RemoveItem I
    str1 = "Select ncode from tblptag where ncode=" & num
    res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
If (res1.EOF = False) Then
    res1.delete
    res1.Update
    res1.CLOSE
End If
End If
Next
For i = 0 To j - 1
    List1.RemoveItem (Val(ritems(i)) - i)
Next
End Sub
Private Sub Form_Load()
str1 = "Select distinct(ncode),chcategory from tblptag"
res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
While (res1.EOF = False)
List1.AddItem (res1!chcategory & "-" & res1!ncode)
 res1.MoveNext
Wend
res1.CLOSE
End Sub
Private Sub ncode_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
add_Click
End If
End Sub

Private Sub OKButton_Click()
Unload Me
End Sub

Private Sub print_Click()

Dim code As String
On Error Resume Next

For i = 0 To List2.ListCount - 1
pos = InStr(1, List2.List(i), "-", vbTextCompare)
num = Val(Mid(List2.List(i), pos + 1))

If (i = 0) Then
code = "(" & num
ElseIf (i = List2.ListCount - 1) Then
code = code & "," & num & ")"
Else
code = code & "," & num
End If
Next





Dim objExcel As Excel.Application, objBook As Excel.Workbook, objsheet As Excel.Worksheet

Set objExcel = GetObject(, "Excel.Application") 'if excel already open you can use GetObject
If Err.Number Then
Err.clear
End If

Set objExcel = CreateObject("Excel.Application") 'or CreateObject to open new Excel Application
''Set objBook = objExcel.Workbooks.Open("c:\daksh\vb\Book3.xls")

Set objBook = objExcel.Workbooks.add
Set objsheet = objBook.Worksheets(1)
objsheet.Cells(1, 1) = "Item"
objsheet.Cells(1, 2) = "Code"
objsheet.Cells(1, 3) = "D.WT"
objsheet.Cells(1, 4) = "C.NAME"
objsheet.Cells(1, 5) = "C.WT"
objsheet.Cells(1, 6) = "G.WT"
objsheet.Cells(1, 7) = "P.NO."
objsheet.Cells(1, 8) = "Purity"

Dim col
col = ""
Dim cweight
cweight = ""
str2 = "SELECT tblcosting.*,IIF(nweight2 > 0, nweight2, '') AS weight2 , IIF(nweight3 > 0, nweight3, '') AS cwt3, (((Tblcosting.minrate1 * tblcosting.nweight1 + Tblcosting.minrate2 * tblcosting.nweight2 + Tblcosting.minrate3 * tblcosting.nweight3+Tblcosting.minrate4 * tblcosting.nweight4)+nmaking1)*1.333333)/100 AS total From tblcosting where ncode in " & code & " ORDER BY chcategory,ncode"
Debug.Print str2

res1.Open str2, MDIForm1.con1, adOpenDynamic, adLockOptimistic

i = 2


While (res1.EOF = False)
cweight = ""
objsheet.Cells(i, 1) = res1!chcategory
objsheet.Cells(i, 2) = res1!ncode
'Debug.Print (res!tblcosting.ncode)
'MsgBox (res!tblcosting.ncode)
objsheet.Cells(i, 3) = res1!nweight1
objsheet.Cells(i, 4) = res1!chcontent2 & " " & res1!chcontent3
cweight = res1!nweight2 + res1!nweight3
objsheet.Cells(i, 5) = Round(cweight, 2)
objsheet.Cells(i, 6) = Round(res1!nweight4, 2)
objsheet.Cells(i, 7) = Round(res1!total, 0)
objsheet.Cells(i, 8) = Round(res1!chpcode, 0)
i = i + 1
res1.MoveNext
Wend
res1.CLOSE
objBook.SaveAs "Tagexpo.xls"
objExcel.Workbooks.Open "TagExpo.xls"
objExcel.Visible = True
'objBook.PrintPreview
''objBook.CLOSE
' cell(1,1) means cell A1 ;
Set objsheet = Null
Set objBook = Null

End Sub

Private Sub remove_Click()
Dim ritems(50) As Integer
j = 0
For i = 0 To List2.ListCount - 1
If (List2.Selected(i) = True) Then
   ' List2.RemoveItem I
   ritems(j) = i
   j = j + 1
End If
Next
For i = 0 To j - 1
    List2.RemoveItem (Val(ritems(i)) - i)
Next

End Sub

Private Sub rlist_Click()
Dim cpos As Integer
Dim ritems(10) As Integer
cpos = List2.ListCount - 1
Dim j As Integer
For i = 0 To cpos
If (List2.Selected(i) = True) Then
    List1.AddItem (List2.List(i))
    pos = InStr(1, List2.List(i), "-", vbTextCompare)
    num = Val(Mid(List2.List(i), pos + 1))
    Category = Mid(List2.List(i), 1, pos - 1)
    'List2.RemoveItem I
     ritems(j) = i
    j = j + 1
    res1.Open "tblptag", MDIForm1.con1, adOpenDynamic, adLockOptimistic
    res1.addnew
    res1!chcategory = Category
    res1!ncode = num
    res1.Update
    res1.CLOSE
   ' cpos = cpos - 1
  '  I = I - 1
  '  List2.Refresh
End If
Next
For i = 0 To j - 1
    List2.RemoveItem (Val(ritems(i)) - i)
Next
End Sub
