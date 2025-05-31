VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form printimg 
   Caption         =   "Print Image"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8955
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7425
   ScaleWidth      =   8955
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   2520
      TabIndex        =   20
      Top             =   360
      Width           =   2055
   End
   Begin VB.DirListBox Dir1 
      Height          =   765
      Left            =   2520
      TabIndex        =   19
      Top             =   720
      Width           =   2055
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid list2 
      Height          =   2295
      Left            =   5640
      TabIndex        =   17
      Top             =   3960
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   4048
      _Version        =   393216
      Rows            =   1
      FixedRows       =   0
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid list1 
      Height          =   2055
      Left            =   360
      TabIndex        =   16
      Top             =   4080
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   3625
      _Version        =   393216
      Rows            =   1
      FixedRows       =   0
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame1 
      Caption         =   "Number Of Images  Per Page"
      Height          =   855
      Left            =   240
      TabIndex        =   11
      Top             =   6480
      Width           =   3135
      Begin VB.OptionButton opt1 
         Caption         =   "1"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton opt2 
         Caption         =   "2"
         Height          =   255
         Left            =   960
         TabIndex        =   14
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton opt4 
         Caption         =   "4"
         Height          =   255
         Left            =   1800
         TabIndex        =   13
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton opt6 
         Caption         =   "6"
         Height          =   255
         Left            =   2400
         TabIndex        =   12
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   5400
      TabIndex        =   8
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "Print"
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   6720
      Width           =   1215
   End
   Begin VB.ListBox scategory 
      Height          =   735
      Left            =   360
      Style           =   1  'Checkbox
      TabIndex        =   6
      Top             =   2160
      Width           =   3255
   End
   Begin VB.CommandButton cmdremall 
      Caption         =   "R&emove All"
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdrem 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdone 
      Caption         =   "&Add"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "A&dd All"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdsearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   2175
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label4 
      Caption         =   "Select Image Path"
      Height          =   255
      Left            =   480
      TabIndex        =   18
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Selected Image"
      Height          =   375
      Left            =   6360
      TabIndex        =   10
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Available Image"
      Height          =   375
      Left            =   600
      TabIndex        =   9
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      Height          =   3135
      Left            =   240
      Top             =   3240
      Width           =   8535
   End
   Begin VB.Label Label1 
      Caption         =   "Select Category"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   1680
      Width           =   1455
   End
End
Attribute VB_Name = "printimg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim res As ADODB.Recordset
Dim con As Connection
Public npict As Integer
Public imgpath As String

Private Sub cmdadd_Click()
On Error Resume Next
Dim i As Integer
res.MoveFirst
List2.Clear
While Not res.EOF
    If Not IsNull(res.Fields(2)) And res.Fields(2) <> "" Then
        With List2
         .TextMatrix(List2.Rows - 1, 0) = res.Fields(0) & "-" & str(res.Fields(1))
           
          i = InStrRev(res.Fields(2), "\")
          If i > 0 Then
            .TextMatrix(List2.Rows - 1, 1) = Mid(res.Fields(2), i + 1, Len(res.Fields(2)))
          Else
            .TextMatrix(List2.Rows - 1, 1) = res.Fields(2)
          End If
          i = 0
         List2.Rows = List2.Rows + 1
         End With
    End If
    res.MoveNext
Wend
List1.Clear
List1.Rows = 1
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdone_Click()
On Error Resume Next
If List1.row >= 0 Then
    If List2.TextMatrix(List2.Rows - 1, 0) <> "" Then
        List2.Rows = List2.Rows + 1
     End If
    List2.TextMatrix(List2.Rows - 1, 0) = List1.TextMatrix(List1.row, 0)
    List2.TextMatrix(List2.Rows - 1, 1) = List1.TextMatrix(List1.row, 1)
    List1.RemoveItem List1.row
End If
End Sub

Private Sub cmdprint_Click()
'On Error GoTo errload
Dim com As Command
Dim i As Integer
Set com = New Command
If opt1.Value = True Then
    npict = 1
ElseIf opt2.Value = True Then
    npict = 2
ElseIf opt4.Value = True Then
    npict = 4
ElseIf opt6.Value = True Then
    npict = 6
End If
imgpath = Dir1.path & "\"
MDIForm1.con1.BeginTrans
'com.CommandText = "create table temp (cnode char(50),img char(50))"
'com.ActiveConnection = MDIForm1.con1
'com.Execute
'For i = 0 To list2.Rows - 1
 '   com.CommandText = "insert into temp values('" & list2.TextMatrix(i, 0) & "','" & list2.TextMatrix(i, 1) & "')"
    'com.Execute
'Next
'MDIForm1.con1.CommitTrans
printform1.Show
Exit Sub
errload:
'MDIForm1.con1.BeginTrans
'com.CommandText = "drop table temp"
'com.Execute
'com.CommandText = "create table temp (cnode char(50),img char(50))"
'com.Execute
'For i = 0 To list2.Rows - 1
 '  com.CommandText = "insert into temp values('" & list2.TextMatrix(i, 0) & "','" & list2.TextMatrix(i, 1) & "')"
 '  com.Execute
'Next
'MDIForm1.con1.CommitTrans
End Sub

Private Sub cmdrem_Click()
On Error Resume Next
If List2.row >= 0 Then
    If List1.TextMatrix(List1.Rows - 1, 0) <> "" Then
        List1.Rows = List1.Rows + 1
     End If
    List1.TextMatrix(List1.Rows - 1, 0) = List2.TextMatrix(List2.row, 0)
    List1.TextMatrix(List1.Rows - 1, 1) = List2.TextMatrix(List2.row, 1)
    If List2.row = List2.Rows - 1 Then
        List2.Clear
    End If
    List2.RemoveItem List2.row
End If

End Sub

Private Sub cmdremall_Click()
On Error Resume Next
Dim i As Integer
'list1.Clear
res.MoveFirst
List1.Clear
While Not res.EOF
    If Not IsNull(res.Fields(2)) And res.Fields(2) <> "" Then
        With List1
         .TextMatrix(List1.Rows - 1, 0) = res.Fields(0) & "-" & str(res.Fields(1))
           
          i = InStrRev(res.Fields(2), "\")
          If i > 0 Then
            .TextMatrix(List1.Rows - 1, 1) = Mid(res.Fields(2), i + 1, Len(res.Fields(2)))
          Else
            .TextMatrix(List1.Rows - 1, 1) = res.Fields(2)
          End If
          i = 0
         List1.Rows = List1.Rows + 1
         End With
    End If
    res.MoveNext
Wend
List2.Clear
List2.Rows = 1
List1.Rows = List1.Rows - 1
End Sub
Private Sub cmdsearch_Click()
Dim st As String
Dim i As Integer, j As Integer
Set res = New ADODB.Recordset
st = "select chcategory,ncode,opicture from tblcosting"
j = 0
List1.Clear
List2.Clear
List1.Rows = 1
List2.Rows = 1
For i = 0 To Me.scategory.ListCount - 1
    If Me.scategory.Selected(i) = True Then
        If j = 0 Then
           st = st + " where chcategory ='" & Me.scategory.List(i) & "'"
           j = 1
        Else
            st = st + " or chcategory ='" & Me.scategory.List(i) & "'"
            End If
     End If
Next
If j = 0 Then
    st = st + " where chstate='Avi.' order by chcategory,ncode"
Else
    st = st + " and chstate='Avi.' order by chcategory,ncode"
End If
res.Open st, MDIForm1.con1, adOpenDynamic, adLockOptimistic
List1.Clear
While Not res.EOF
    If Not IsNull(res.Fields(2)) And res.Fields(2) <> "" Then
        With List1
         .TextMatrix(List1.Rows - 1, 0) = res.Fields(0) & "-" & str(res.Fields(1))
           
          i = InStrRev(res.Fields(2), "\")
          If i > 0 Then
            .TextMatrix(List1.Rows - 1, 1) = Mid(res.Fields(2), i + 1, Len(res.Fields(2)))
          Else
            .TextMatrix(List1.Rows - 1, 1) = res.Fields(2)
          End If
          i = 0
         List1.Rows = List1.Rows + 1
         End With
    End If
    res.MoveNext
Wend
List1.Rows = List1.Rows - 1
End Sub

Private Sub Drive1_Change()
On Error GoTo errload
Dir1.path = Drive1.List(Drive1.ListIndex)
Exit Sub
errload:
Dir1.path = ""
End Sub

Private Sub Form_Load()
On Error GoTo errload
Set res = New ADODB.Recordset
'Set con = New ADODB.Connection
'With con
  '  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "db2.mdb;Persist Security Info=False"
  '  .Open
'End With
res.Open "Select distinct chcategory from tblcosting ", MDIForm1.con1, adOpenDynamic, adLockOptimistic
If res.EOF <> True And res.BOF <> True Then
    While Not res.EOF
        If res.Fields(0) <> "" Then
        scategory.AddItem res.Fields(0)
        End If
        res.MoveNext
        
    Wend
End If
Exit Sub
errload:
Me.Image1.Picture = LoadPicture("")
MsgBox Err.Description
End Sub

Private Sub list1_Click()
On Error GoTo imgload
Me.Image1.Picture = LoadPicture(Dir1.path & "\" & List1.TextMatrix(List1.row, 1))
Exit Sub
imgload:
MsgBox "Image Location Not Valid"
End Sub

