VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form appdetails 
   Caption         =   "Approval Details"
   ClientHeight    =   5715
   ClientLeft      =   2655
   ClientTop       =   1605
   ClientWidth     =   7275
   LinkTopic       =   "Form8"
   MDIChild        =   -1  'True
   ScaleHeight     =   5715
   ScaleWidth      =   7275
   WindowState     =   2  'Maximized
   Begin VB.ComboBox dis 
      Height          =   315
      ItemData        =   "AppDetails.frx":0000
      Left            =   4560
      List            =   "AppDetails.frx":001C
      TabIndex        =   16
      Text            =   "25%"
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton cprint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   6240
      TabIndex        =   15
      Top             =   1200
      Width           =   1335
   End
   Begin VB.ListBox List2 
      Height          =   735
      ItemData        =   "AppDetails.frx":0047
      Left            =   4440
      List            =   "AppDetails.frx":0049
      Style           =   1  'Checkbox
      TabIndex        =   14
      Top             =   120
      Width           =   1575
   End
   Begin VB.ComboBox sowner 
      Height          =   315
      ItemData        =   "AppDetails.frx":004B
      Left            =   1080
      List            =   "AppDetails.frx":005B
      TabIndex        =   12
      Top             =   240
      Width           =   615
   End
   Begin VB.ListBox List1 
      Height          =   960
      ItemData        =   "AppDetails.frx":006F
      Left            =   1680
      List            =   "AppDetails.frx":0071
      Style           =   1  'Checkbox
      TabIndex        =   11
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton change 
      Caption         =   "Change &Status"
      Height          =   375
      Left            =   6240
      TabIndex        =   10
      Top             =   840
      Width           =   1335
   End
   Begin VB.ComboBox status 
      Height          =   315
      ItemData        =   "AppDetails.frx":0073
      Left            =   1680
      List            =   "AppDetails.frx":007D
      TabIndex        =   9
      Top             =   1920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox Category 
      Height          =   315
      ItemData        =   "AppDetails.frx":008E
      Left            =   2880
      List            =   "AppDetails.frx":00B6
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton print 
      Caption         =   "&Report"
      Height          =   375
      Left            =   6240
      TabIndex        =   6
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton Clear 
      Caption         =   "C&lear"
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Search 
      Caption         =   "&Search"
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton close 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6240
      TabIndex        =   0
      Top             =   1920
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4815
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   8493
      _Version        =   393216
      Cols            =   15
      RowHeightMin    =   315
      AllowUserResizing=   1
      MousePointer    =   1
   End
   Begin VB.Label cap 
      Height          =   375
      Left            =   360
      TabIndex        =   17
      Top             =   2280
      Width           =   3255
   End
   Begin VB.Label Label4 
      Caption         =   "Owner"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   240
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   7680
      X2              =   7680
      Y1              =   0
      Y2              =   2760
   End
   Begin VB.Image Image1 
      Height          =   2415
      Left            =   8400
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Category"
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "App. No."
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Issue Pr."
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   12000
      Y1              =   2760
      Y2              =   2760
   End
End
Attribute VB_Name = "appdetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim res1 As New ADODB.Recordset
Dim str1 As String

Private Sub change_Click()
On Error Resume Next
str1 = ""
str2 = ""
updatedcode = ""
Dim ritems(200) As Integer
Dim k As Integer
For j = 0 To MSFlexGrid1.Rows - 1
If (MSFlexGrid1.TextMatrix(j, 14) = "Rec.") Then
    pos = InStr(1, MSFlexGrid1.TextMatrix(j, 1), "-", vbTextCompare)
    num = Val(Mid(MSFlexGrid1.TextMatrix(j, 1), pos + 1))
    ritems(k) = j
    k = k + 1

'If (j <> MSFlexGrid1.Rows - 1) Then
str1 = str1 & MSFlexGrid1.TextMatrix(j, 1) & ","
updatedcode = updatedcode & num & ","
'Else
'str1 = str1 & MSFlexGrid1.TextMatrix(j, 1)
'updatedcode = updatedcode & num
'End If
End If
Next j
'MsgBox (str1)
pos = InStrRev(str1, ",")
str1 = Mid(str1, 1, pos - 1)
If (str1 = "") Then
MsgBox ("Please Change The Status Of Any Item")
Exit Sub
End If

pos = InStrRev(updatedcode, ",")
updatedcode = Mid(updatedcode, 1, pos - 1)

result = MsgBox("Do You Want To Change The Status Of Thease Items " & str1, vbYesNo + vbDefaultButton2, "Confirmation")

    If (Val(result) = 7) Then
    
    Exit Sub
    Else
      str2 = "Select * from tblappdetail where ncode in(" & updatedcode & ")"
      'MsgBox (str2)
    ' Exit Sub
      res1.Open str2, MDIForm1.con1, adOpenStatic, adLockOptimistic
    If (res1.RecordCount > 0) Then
        res1.MoveFirst
    Do While Not res1.EOF
        res1.delete
        res1.MoveNext
    Loop
    End If
    res1.close
    For i = 0 To k - 1
    MSFlexGrid1.RemoveItem (Val(ritems(i)) - i)
    Next
    
    str3 = "select * from tblappmaster where nappno not in(select nappno from tblappdetail)"
    res1.Open str3, MDIForm1.con1, adOpenStatic, adLockOptimistic
     If (res1.RecordCount > 0) Then
        res1.MoveFirst
    
    Do While Not res1.EOF
        res1.delete
        res1.MoveNext
    Loop
    End If
    res1.close
    loaditem
    
   MsgBox ("Record Deleted")
End If
End Sub

Private Sub Clear_Click()
Category = ""
appno = ""
issue = ""
cap.Caption = ""
clearing
End Sub

Private Sub close_Click()
Unload Me
End Sub

Private Sub cprint_Click()
Dim ptr As Printer
result = PrintMSFlexgrid("", vbPRORPortrait, MSFlexGrid1, 1)
'Dialog.Show

'If (str1 = "") Then
'str1 = "SELECT tblappmaster.chissue, tblappdetail.chcategory, tblappdetail.*, [tblappdetail].[nrate1]*[tblappdetail].[nweight1]+[tblappdetail].[nrate2]*[tblappdetail].[nweight2]+[tblappdetail].[nrate3]*[tblappdetail].[nweight3]+[tblappdetail].[nrate4]*[tblappdetail].[nweight4]+[tblappdetail].[nmaking1] AS total, (total * 1.333333)/100 AS priceno,tblappmaster.nappno, Tblcosting.chowner "
'str1 = str1 & " FROM Tblcosting RIGHT JOIN (tblappmaster INNER JOIN tblappdetail ON tblappmaster.nappno = tblappdetail.nappno) ON Tblcosting.ncode = tblappdetail.ncode order by chissue,tblappmaster.nappno"
'End If
'res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
''Printer.CurrentY = 160
''Printer.FontSize = 10
''Printer.PaperSize = vbPRPSEnvDL
'
''I = 0
'While (res1.EOF = False)
''If ((I = 0) Or (I = 6) Or (I = 11) Or (I = 16) Or (I = 21) Or (I = 26) Or (I = 31) Or (I = 36) Or (I = 41) Or (I = 46) Or (I = 51) Or (I = 56) Or (I = 61) Or (I = 66) Or (I = 71)) Then
''Printer.CurrentY = Printer.Height / 45
''End If
''Printer.Print (res1!chcategory & "-" & res1!ncode & "/" & Round(res1!priceno)); Spc(1)
''I = I + 1
'res1.MoveNext
'Wend
''Printer.EndDoc
End Sub
Public Function PrintMSFlexgrid(title As String, orientation As Integer, grdName As MSFlexGrid, line_spaces As Integer)

' Local declares
Dim lRowCount As Long
Dim iColCount As Integer
Dim lRowLoop As Long
Dim iColLoop As Integer
Dim iTabPos As Integer
Dim found_legend As Boolean
Dim print_line As Boolean
Dim top As Integer
Dim outstr As String

Printer.orientation = orientation
' print the title if one has been  entered.

If title <> "" Then
    Printer.FontName = grdName.Font
    Printer.FontName = "Comic Sans MS"
    Printer.FontSize = 14
    Printer.Print title
    Printer.Print ""
    Printer.Print ""
End If
Printer.Print ""
Printer.Print ""
Printer.Print ""

'loop through the cells and print them
iTabPos = 0
' Start function
lRowCount = grdName.Rows - 1
iColCount = grdName.Cols - 1
For lRowLoop = 0 To lRowCount
    grdName.row = lRowLoop
    For iColLoop = 0 To iColCount
        If grdName.ColWidth(iColLoop) <> 0 Then
            grdName.col = iColLoop
            ' Grab the flexgrid font properties.

           ' ptr.FontName = "Comic Sans MS" 'grdName.CellFontName
          '  ptr.FontSize = 8 'grdName.CellFontSize
           ' ptr.FontBold = grdName.CellFontBold
           ' ptr.FontItalic = grdName.CellFontItalic
           ' ptr.FontUnderline = grdName.CellFontUnderline
           ' ptr.ForeColor = grdName.CellForeColor
                
            Printer.Print Tab(iTabPos + 1); grdName.Text;
            iTabPos = iTabPos + (grdName.CellWidth / 60)

'           Printer.NewPage
        End If
    Next iColLoop
    Printer.Print ""
    For i = 1 To line_spaces
        Printer.Print ""
    Next i
    iTabPos = 0
Next lRowLoop
Printer.EndDoc
End Function

Private Sub Form_Load()
loaditem
'Combo1.AddItem List2
MSFlexGrid1.TextMatrix(0, 0) = "S.No."
MSFlexGrid1.ColWidth(0) = 350
MSFlexGrid1.TextMatrix(0, 1) = "Item Code"
MSFlexGrid1.ColWidth(1) = 950
MSFlexGrid1.TextMatrix(0, 2) = "App.No."
MSFlexGrid1.ColWidth(2) = 700
MSFlexGrid1.TextMatrix(0, 3) = "Issue"
MSFlexGrid1.ColWidth(3) = 1250
MSFlexGrid1.TextMatrix(0, 4) = "Dia.Wt."
MSFlexGrid1.ColWidth(4) = 750
MSFlexGrid1.TextMatrix(0, 5) = "Dia.Rt."
MSFlexGrid1.ColWidth(5) = 950
MSFlexGrid1.TextMatrix(0, 6) = "Stone1."
MSFlexGrid1.ColWidth(6) = 850
MSFlexGrid1.TextMatrix(0, 7) = "St1 Wt."
MSFlexGrid1.ColWidth(7) = 750
MSFlexGrid1.TextMatrix(0, 8) = "St1 Rt."
MSFlexGrid1.ColWidth(8) = 750
MSFlexGrid1.TextMatrix(0, 9) = "MetalWt."
MSFlexGrid1.ColWidth(9) = 600
MSFlexGrid1.TextMatrix(0, 10) = "MetatRt."
MSFlexGrid1.ColWidth(10) = 600
MSFlexGrid1.TextMatrix(0, 11) = "Making"
MSFlexGrid1.ColWidth(11) = 850
MSFlexGrid1.TextMatrix(0, 12) = "Total Amt."
MSFlexGrid1.ColWidth(12) = 1000
MSFlexGrid1.TextMatrix(0, 13) = ""
MSFlexGrid1.ColWidth(13) = 0
MSFlexGrid1.TextMatrix(0, 14) = "Status"
MSFlexGrid1.ColWidth(14) = 1000
End Sub
Private Sub loaditem()
str1 = "Select distinct(chissue) from tblappmaster order by chissue"
res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
If (List1.ListCount > 0) Then
  List1.Clear
End If
While (res1.EOF = False)
List1.AddItem (res1!chissue)
 res1.MoveNext
Wend
res1.close
str1 = "Select nappno from tblappmaster order by nappno"
res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
If (List2.ListCount > 0) Then
List2.Clear
End If
While (res1.EOF = False)
List2.AddItem (res1!nappno)
'appno.AddItem (res1!nappno)
res1.MoveNext
Wend
res1.close
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
    End Sub
Private Sub MSFlexGrid1_Click()
On Error Resume Next
    If (MSFlexGrid1.col = 14 And MSFlexGrid1.row <> MSFlexGrid1.Rows - 1) Then
       status.Visible = True
        status.Width = MSFlexGrid1.CellWidth
        ' status.Height = MSFlexGrid1.CellHeight
    status.top = MSFlexGrid1.CellTop + MSFlexGrid1.top
    status.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
    status.Text = MSFlexGrid1.Text
    
    'Text1.Text = MSFlexGrid1.Text
    'Text1.SelStart = 0
    'Text1.SelLength = Len(Text1.Text)
   ' status.ZOrder
    status.SetFocus
    End If
End Sub



Private Sub MSFlexGrid1_LeaveCell()
If (MSFlexGrid1.col = 14 And MSFlexGrid1.row <> MSFlexGrid1.Rows - 1) Then
MSFlexGrid1.Text = status.Text
status.Visible = False
End If
End Sub

Private Sub MSFlexGrid1_LostFocus()
If (MSFlexGrid1.col = 14 And MSFlexGrid1.row <> MSFlexGrid1.Rows - 1) Then
MSFlexGrid1.Text = status.Text
End If
End Sub



Private Sub MSFlexGrid1_Scroll()
status.Visible = False
End Sub

Private Sub MSFlexGrid1_SelChange()
On Error Resume Next
If (MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 13) <> "" And MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 13) <> Empty) Then
     Dim picturename As String
     pos = InStrRev(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 13), "\")
     If (pos <> 0) Then
     picturename = Mid(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 13), pos + 1)
     Else
     picturename = MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 13)
     End If
     Image1.Visible = True
     Image1.Picture = LoadPicture(MDIForm1.picturepath & picturename)
     'Image1.Picture = LoadPicture(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 17))
Else
    Image1.Picture = LoadPicture()
End If

        pos = InStr(1, MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1), "-", vbTextCompare)
        num = Val(Mid(MSFlexGrid1.TextMatrix(MSFlexGrid1.row, 1), pos + 1))
        str1 = "select chowner,chpcode,ndno,dinvoicedate from tblcosting where ncode=" & num
        'MsgBox (str1)
        res1.Open str1, MDIForm1.con1, adOpenStatic, adLockOptimistic
        If (res1.EOF = False) Then
              cap.Caption = "Owner:" & res1!chowner & " And Code " & res1!chpcode & " And P.No." & res1!ndno & " Invoice Date:" & res1!dinvoicedate
        End If
        res1.close
End Sub
Private Sub print_Click()
str1 = " SELECT tblappdetail.chcategory, tblappdetail.*, Tblcosting.chowner, Tblcosting.chpcode "
str1 = str1 & " FROM Tblcosting RIGHT JOIN (tblappmaster INNER JOIN tblappdetail ON tblappmaster.nappno = tblappdetail.nappno) ON Tblcosting.ncode = tblappdetail.ncode "


listdt = ""
listapp = ""
    For Id = 0 To List1.ListCount - 1

     If (List1.Selected(Id) <> False) Then
     listdt = listdt & "'" & List1.List(Id) & "',"
     End If
    Next
    'MsgBox (listdt)
    
    If (listdt <> "") Then
     pos = InStrRev(listdt, ",")
     listdt = Mid(listdt, 1, pos - 1)
     str2 = str2 & " chissue in(" & listdt & ")"
    End If

For Id = 0 To List2.ListCount - 1
If (List2.Selected(Id) <> False) Then
     listapp = listapp & List2.List(Id) & ","
     End If
    Next

If (listapp <> "") Then
     pos = InStrRev(listapp, ",")
     listapp = Mid(listapp, 1, pos - 1)
    If (str2 <> Empty) Then
     str2 = str2 & " And tblappmaster.nappno in(" & listapp & ")"
     Else
     str2 = str2 & " tblappmaster.nappno in(" & listapp & ")"
    End If
End If

'If (appno.Text <> Empty) Then
   ' If (str2 <> Empty) Then
  '  str2 = str2 & " And tblappmaster.nappno=" & Val(appno.Text)
  '  Else
  '  str2 = str2 & " tblappmaster.nappno=" & Val(appno.Text)
  '  End If
'End If

If (Category.Text <> Empty) Then
    If (str2 <> Empty) Then
    str2 = str2 & " And tblcosting.chcategory=" & "'" & Category.Text & "'"
    Else
    str2 = str2 & " tblcosting.chcategory=" & "'" & Category.Text & "'"
    End If
End If

If (sowner.Text <> Empty) Then
    If (str2 <> Empty) Then
    str2 = str2 & " And tblcosting.chowner=" & "'" & sowner.Text & "'"
    Else
    str2 = str2 & " tblcosting.chowner=" & "'" & sowner.Text & "'"
    End If
End If

If (str2 <> Empty) Then
str1 = str1 & " where " & str2 & " order by tblappdetail.chcategory,tblappdetail.ncode"
Else
str1 = str1 & " order by tblappdetail.CHCATeGoRY,tblappdetail.NCODE"
End If
'MsgBox (str1)
Debug.Print (str1)
If DataEnvironment1.rsCommand1.state = adStateOpen Then
DataEnvironment1.rsCommand1.close
End If
DataEnvironment1.Commands(1).CommandText = str1
DataReport9.Show
End Sub
Private Sub Search_Click()
clearing
str1 = " SELECT tblappmaster.chissue, tblappdetail.chcategory, tblappdetail.*, [Tblcosting].[minrate1]*[Tblcosting].[nweight1] AS amt1, [Tblcosting].[minrate2]*[Tblcosting].[nweight2] AS amt2, [Tblcosting].[minrate3]*[Tblcosting].[nweight3] AS amt3,"
str1 = str1 & "[Tblcosting].[minrate4]*[Tblcosting].[nweight4] AS amt4, amt1+amt2+amt3+amt4+[Tblcosting].[nmaking1] AS total, (total * 1.333333)/100 AS priceno, total * 1.333333-(((total * 1.333333) * " & Val(dis.Text) & ")/100) AS ftotamt, tblappmaster.nappno, Tblcosting.chowner,(ftotamt-(amt2+amt3+amt4+Tblcosting.nmaking1))/Tblcosting.nweight1 AS dprice, dprice*Tblcosting.nweight1 AS damt"
str1 = str1 & " FROM Tblcosting RIGHT JOIN (tblappmaster INNER JOIN tblappdetail ON tblappmaster.nappno = tblappdetail.nappno) ON Tblcosting.ncode = tblappdetail.ncode "
'str1 = "SELECT tblappmaster.chissue, tblappdetail.chcategory, tblappdetail.*, [tblappdetail].[nrate1]*[tblappdetail].[nweight1]+[tblappdetail].[nrate2]*[tblappdetail].[nweight2]+[tblappdetail].[nrate3]*[tblappdetail].[nweight3]+[tblappdetail].[nrate4]*[tblappdetail].[nweight4]+[tblappdetail].[nmaking1] AS total, (total * 1.333333)/100 AS priceno,priceno-((priceno)/100 * " & Val(dis.Text) & " AS ftotamt,tblappmaster.nappno, Tblcosting.chowner "
'str1 = str1 & ", (ftotamt-(amt2+amt3+amt4+nmaking1))/nweight1 as dprice,dprice*Tblcosting.nweight1 as damt FROM Tblcosting RIGHT JOIN (tblappmaster INNER JOIN tblappdetail ON tblappmaster.nappno = tblappdetail.nappno) ON Tblcosting.ncode = tblappdetail.ncode"
'MsgBox (str1)
listdt = ""
listapp = ""
Dim issuein As String
    For Id = 0 To List1.ListCount - 1

     If (List1.Selected(Id) <> False) Then
     listdt = listdt & "'" & List1.List(Id) & "',"
     End If
    Next
    'MsgBox (listdt)
    
    If (listdt <> "") Then
     pos = InStrRev(listdt, ",")
     issuein = Mid(listdt, 1, pos - 1)
     'listdt = MDIForm1.parserText(listdt)
     'issuein = MDIForm1.parserText(issuein)
     str2 = str2 & " chissue in(" & issuein & ")"
    End If

For Id = 0 To List2.ListCount - 1
If (List2.Selected(Id) <> False) Then
     listapp = listapp & List2.List(Id) & ","
     End If
    Next

If (listapp <> "") Then
     pos = InStrRev(listapp, ",")
     listapp = Mid(listapp, 1, pos - 1)
    If (str2 <> Empty) Then
     str2 = str2 & " And tblappmaster.nappno in(" & listapp & ")"
     Else
     str2 = str2 & " tblappmaster.nappno in(" & listapp & ")"
    End If
End If

'If (appno.Text <> Empty) Then
   ' If (str2 <> Empty) Then
  '  str2 = str2 & " And tblappmaster.nappno=" & Val(appno.Text)
  '  Else
  '  str2 = str2 & " tblappmaster.nappno=" & Val(appno.Text)
  '  End If
'End If

If (Category.Text <> Empty) Then
    If (str2 <> Empty) Then
    str2 = str2 & " And tblcosting.chcategory=" & "'" & Category.Text & "'"
    Else
    str2 = str2 & " tblcosting.chcategory=" & "'" & Category.Text & "'"
    End If
End If

If (sowner.Text <> Empty) Then
    If (str2 <> Empty) Then
    str2 = str2 & " And tblcosting.chowner=" & "'" & MDIForm1.parserText(sowner.Text) & "'"
    Else
    str2 = str2 & " tblcosting.chowner=" & "'" & MDIForm1.parserText(sowner.Text) & "'"
    End If
End If

If (str2 <> Empty) Then
str1 = str1 & " where " & str2 & " order by tblappdetail.chcategory,tblappdetail.ncode"
Else
str1 = str1 & " order by tblappdetail.chcategory,tblappdetail.ncode"
End If


'MsgBox (str1)
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

While (res1.EOF = False)
    MSFlexGrid1.TextMatrix(i, 0) = i
    MSFlexGrid1.TextMatrix(i, 1) = res1.Fields(1) & "-" & res1!ncode
    MSFlexGrid1.TextMatrix(i, 2) = res1.Fields(2)
    MSFlexGrid1.TextMatrix(i, 3) = res1!chissue
    dweight = res1!nweight1
    MSFlexGrid1.TextMatrix(i, 4) = dweight
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 4
    MSFlexGrid1.CellBackColor = RGB(231, 228, 218)
    'MSFlexGrid1.TextMatrix(I, 5) = Round(res1!nrate1)
    MSFlexGrid1.TextMatrix(i, 5) = Round(res1!dprice)
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 5
    MSFlexGrid1.CellBackColor = RGB(231, 228, 218)
    
    MSFlexGrid1.TextMatrix(i, 6) = res1!chcontent2
    swt1 = res1!nweight2
    MSFlexGrid1.TextMatrix(i, 7) = swt1
    
    MSFlexGrid1.TextMatrix(i, 8) = res1!nrate2
    
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 7
    MSFlexGrid1.CellBackColor = RGB(231, 228, 218)
    
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 8
    MSFlexGrid1.CellBackColor = RGB(231, 228, 218)
    
    
    mwt = res1!nweight4
    
    MSFlexGrid1.TextMatrix(i, 9) = Round(mwt, 2)
    MSFlexGrid1.TextMatrix(i, 10) = res1!nrate4
    
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 9
    MSFlexGrid1.CellBackColor = RGB(234, 248, 222)
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 10
    MSFlexGrid1.CellBackColor = RGB(234, 248, 222)
    
    making = res1!nmaking1
    MSFlexGrid1.TextMatrix(i, 11) = making
    
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 11
    MSFlexGrid1.CellBackColor = RGB(255, 255, 236)

'    MSFlexGrid1.TextMatrix(I, 12) = Round(Val(res1!priceno))
    MSFlexGrid1.TextMatrix(i, 12) = Round(res1!ftotamt)
    
    metalamt = metalamt + Val(MSFlexGrid1.TextMatrix(i, 9)) * Val(MSFlexGrid1.TextMatrix(i, 10))
    stonetotamt1 = stonetotamt1 + Val(MSFlexGrid1.TextMatrix(i, 7)) * Val(MSFlexGrid1.TextMatrix(i, 8))
    diatotrt = diatotrt + Val(MSFlexGrid1.TextMatrix(i, 4)) * Val(MSFlexGrid1.TextMatrix(i, 5))
    
    
    If (IsNull(res1!opicture)) Then
        MSFlexGrid1.TextMatrix(i, 13) = ""
    ElseIf (res1!opicture <> "" And res1!opicture <> Empty) Then
        MSFlexGrid1.TextMatrix(i, 13) = res1!opicture
    End If
    MSFlexGrid1.TextMatrix(i, 14) = "Issue"
    
        
    gtotal = gtotal + res1!ftotamt
    dwt = dwt + dweight
    metal = metal + mwt
    stonewt1 = stonewt1 + swt1
        
    makingtot = makingtot + making
     i = i + 1
    MSFlexGrid1.Rows = i + 1
    res1.MoveNext
   
    Wend

    MSFlexGrid1.TextMatrix(i, 1) = "Grand Total:"
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 1
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
       
    MSFlexGrid1.TextMatrix(i, 4) = Round(dwt, 2)
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 4
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
        
    MSFlexGrid1.TextMatrix(i, 5) = Round(diatotrt)
    
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 5
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
        
    MSFlexGrid1.TextMatrix(i, 7) = Round(stonewt1, 2)
    
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 7
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
    
    MSFlexGrid1.TextMatrix(i, 8) = Round(stonetotamt1)
    
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 8
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
    
    
    
    MSFlexGrid1.TextMatrix(i, 9) = Round(metal, 2)
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 9
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
    
    MSFlexGrid1.TextMatrix(i, 10) = Round(metalamt)
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 10
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
    
    MSFlexGrid1.TextMatrix(i, 11) = Round(makingtot)
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 11
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
    
    MSFlexGrid1.TextMatrix(i, 12) = Round(gtotal)
    MSFlexGrid1.row = i
    MSFlexGrid1.col = 12
    MSFlexGrid1.CellForeColor = RGB(255, 51, 51)
    End If
    res1.close
End Sub


Private Sub status_Change()
MSFlexGrid1.Text = status.Text
status.Visible = False
MSFlexGrid1.SetFocus
End Sub

Private Sub status_Click()
MSFlexGrid1.Text = status.Text
status.Visible = False
MSFlexGrid1.SetFocus
End Sub
