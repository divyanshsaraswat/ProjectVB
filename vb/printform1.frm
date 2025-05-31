VERSION 5.00
Begin VB.Form printform1 
   Caption         =   "Printform"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7995
   MDIChild        =   -1  'True
   ScaleHeight     =   6210
   ScaleWidth      =   7995
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      Begin VB.CommandButton cmdnext 
         Caption         =   ">>"
         Height          =   375
         Left            =   2880
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdpre 
         Caption         =   "<<"
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Print"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label lab 
      Caption         =   "Label1"
      Height          =   615
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   3960
      Width           =   3615
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2895
      Index           =   0
      Left            =   240
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   3615
   End
End
Attribute VB_Name = "printform1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As Connection
Dim res As ADODB.Recordset

Private Sub cmdnext_Click()
show_next
End Sub

Private Sub cmdpre_Click()
show_pre
End Sub

Private Sub Command1_Click()
On Error GoTo errprint
Me.Frame1.Visible = False
Me.printform
Me.Frame1.Visible = True
Exit Sub
errprint:
MsgBox "Printer Error"
End Sub

Private Sub Form_Activate()
On Error Resume Next
Set con = New Connection
Set res = New ADODB.Recordset
'With con
 '   .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db2.mdb;Persist Security Info=False"
  '  .Open
'End With
res.Open "select * from temp ", MDIForm1.con1, adOpenDynamic, adLockOptimistic
If printimg.npict = 1 Then
Me.Image1(0).Height = Me.Height - 1000
Me.Image1(0).Width = Me.Width - 900
Me.lab(0).top = Me.Image1(0).Height
Me.lab(0).Left = Me.Image1(0).Left
ElseIf printimg.npict = 2 Then
Me.Image1(0).Height = (Me.Height - 1000) / 2
Me.Image1(0).Width = Me.Width - 900
Me.lab(0).top = Me.Image1(0).Height
Me.lab(0).Left = Me.Image1(0).Left
Load Image1(1)
Load lab(1)
Image1(1).top = Me.Image1(0).Height
Image1(1).Left = Me.Image1(0).Left + 100
Me.Image1(1).Height = (Me.Height - 1000) / 2
Me.Image1(1).Width = Me.Width - 900
Me.Image1(1).Visible = True
Me.Image1(1).BorderStyle = Me.Image1(0).BorderStyle
Me.lab(1).top = Image1(1).Height
Me.lab(1).Left = Image1(1).Left
Me.lab(1).Height = Me.lab(0).Height
Me.lab(1).Width = Me.lab(0).Width
Me.lab(1).Visible = True
ElseIf printimg.npict = 4 Then
Me.Image1(0).Height = (Me.Height - 1000) / 2
Me.Image1(0).Width = (Me.Width - 900) / 2
Me.lab(0).top = Me.Image1(0).Height
Me.lab(0).Left = Me.Image1(0).Left
Load Image1(1)
Load lab(1)
Me.Image1(1).top = Me.Image1(0).top
Me.Image1(1).Left = Me.Image1(0).Width
Me.Image1(1).Visible = True
Me.Image1(1).BorderStyle = Me.Image1(0).BorderStyle
Me.Image1(1).Height = Me.Image1(0).Height
Me.Image1(1).Width = Me.Image1(0).Width
Me.lab(1).top = Image1(1).Height
Me.lab(1).Left = Image1(1).Left
Me.lab(1).Height = Me.lab(0).Height
Me.lab(1).Width = Me.lab(0).Width
Me.lab(1).Visible = True
Load Image1(2)
Load lab(2)
Me.Image1(2).top = Me.Image1(0).Height
Me.Image1(2).Left = Me.Image1(0).Left
Me.Image1(2).Visible = True
Me.Image1(2).BorderStyle = Me.Image1(0).BorderStyle
Me.Image1(2).Height = Me.Image1(0).Height
Me.Image1(2).Width = Me.Image1(0).Width
Me.lab(2).top = Image1(2).Height
Me.lab(2).Left = Image1(2).Left
Me.lab(2).Height = Me.lab(0).Height
Me.lab(2).Width = Me.lab(0).Width
Me.lab(2).Visible = True
Load Image1(3)
Load lab(3)
Me.Image1(3).top = Me.Image1(0).Height
Me.Image1(3).Left = Me.Image1(2).Width
Me.Image1(3).Visible = True
Me.Image1(3).BorderStyle = Me.Image1(0).BorderStyle
Me.Image1(3).Height = Me.Image1(0).Height
Me.Image1(3).Width = Me.Image1(0).Width
Me.lab(3).top = Image1(3).Height
Me.lab(3).Left = Image1(3).Left
Me.lab(3).Height = Me.lab(0).Height
Me.lab(3).Width = Me.lab(0).Width
Me.lab(3).Visible = True
ElseIf printimg.npict = 6 Then

Me.Image1(0).Height = (Me.Height - 1000) / 3
Me.Image1(0).Width = (Me.Width - 900) / 2
Me.lab(0).top = Me.Image1(0).Height
Me.lab(0).Left = Me.Image1(0).Left
Load Image1(1)
Load lab(1)
Me.Image1(1).top = Me.Image1(0).top
Me.Image1(1).Left = Me.Image1(0).Width
Me.Image1(1).Visible = True
Me.Image1(1).BorderStyle = Me.Image1(0).BorderStyle
Me.Image1(1).Height = Me.Image1(0).Height
Me.Image1(1).Width = Me.Image1(0).Width
Me.lab(1).top = Image1(1).Height
Me.lab(1).Left = Image1(1).Left
Me.lab(1).Height = Me.lab(0).Height
Me.lab(1).Width = Me.lab(0).Width
Me.lab(1).Visible = True
Load Image1(2)
Load lab(2)
Me.Image1(2).top = Me.Image1(0).Height
Me.Image1(2).Left = Me.Image1(0).Left
Me.Image1(2).Visible = True
Me.Image1(2).BorderStyle = Me.Image1(0).BorderStyle
Me.Image1(2).Height = Me.Image1(0).Height
Me.Image1(2).Width = Me.Image1(0).Width
Me.lab(2).top = Image1(2).Height
Me.lab(2).Left = Image1(2).Left
Me.lab(2).Height = Me.lab(0).Height
Me.lab(2).Width = Me.lab(0).Width
Me.lab(2).Visible = True
Load Image1(3)
Load lab(3)
Me.Image1(3).top = Me.Image1(0).Height
Me.Image1(3).Left = Me.Image1(2).Width
Me.Image1(3).Visible = True
Me.Image1(3).BorderStyle = Me.Image1(0).BorderStyle
Me.Image1(3).Height = Me.Image1(0).Height
Me.Image1(3).Width = Me.Image1(0).Width
Me.lab(3).top = Image1(3).Height
Me.lab(3).Left = Image1(3).Left
Me.lab(3).Height = Me.lab(0).Height
Me.lab(3).Width = Me.lab(0).Width
Me.lab(3).Visible = True
Load Image1(4)
Load lab(4)
Me.Image1(4).top = (Me.Image1(2).Height) * 2
Me.Image1(4).Left = Me.Image1(2).Left
Me.Image1(4).Visible = True
Me.Image1(4).BorderStyle = Me.Image1(0).BorderStyle
Me.Image1(4).Height = Me.Image1(0).Height
Me.Image1(4).Width = Me.Image1(0).Width
Me.lab(4).top = Image1(4).Height
Me.lab(4).Left = Image1(4).Left
Me.lab(4).Height = Me.lab(0).Height
Me.lab(4).Width = Me.lab(0).Width
Me.lab(4).Visible = True
Load Image1(5)
Load lab(5)
Me.Image1(5).top = (Me.Image1(2).Height) * 2
Me.Image1(5).Left = Me.Image1(4).Width
Me.Image1(5).Visible = True
Me.Image1(5).BorderStyle = Me.Image1(0).BorderStyle
Me.Image1(5).Height = Me.Image1(0).Height
Me.Image1(5).Width = Me.Image1(0).Width
Me.lab(5).top = Image1(5).Height
Me.lab(5).Left = Image1(5).Left
Me.lab(5).Height = Me.lab(0).Height
Me.lab(5).Width = Me.lab(0).Width
Me.lab(5).Visible = True
End If
show_next
End Sub
Public Sub show_next()
    On Error Resume Next
    Dim i As Integer
    If res.EOF = True Then
        res.MovePrevious
    End If
    If res.EOF <> True And res.BOF <> True Then
        If printimg.npict = 1 Then
            Image1(0).Picture = LoadPicture(printimg.imgpath & res.Fields(1))
            lab(0).Caption = res.Fields(0)
        ElseIf printimg.npict = 2 Then
            Image1(0).Picture = LoadPicture(printimg.imgpath & res.Fields(1))
            lab(0).Caption = res.Fields(0)
            res.MoveNext
            If res.EOF <> True And res.BOF <> True Then
                 Image1(1).Picture = LoadPicture(printimg.imgpath & res.Fields(1))
                 lab(1).Caption = res.Fields(0)
            End If
        ElseIf printimg.npict = 4 Then
            For i = 0 To 3
            If res.EOF <> True And res.BOF <> True Then
                Image1(i).Picture = LoadPicture(printimg.imgpath & res.Fields(1))
                lab(i).Caption = res.Fields(0)
                res.MoveNext
            End If
            Next
        ElseIf printimg.npict = 6 Then
            For i = 0 To 5
            If res.EOF <> True And res.BOF <> True Then
                Image1(i).Picture = LoadPicture(printimg.imgpath & res.Fields(1))
                lab(i).Caption = res.Fields(0)
                res.MoveNext
            End If
            Next
       End If
   End If
          
End Sub
Public Sub show_pre()
On Error Resume Next
    Dim i As Integer
    If res.BOF = True Then
        res.MoveNext
    End If
    If res.EOF <> True And res.BOF <> True Then
        If printimg.npict = 1 Then
            Image1(0).Picture = LoadPicture(printimg.imgpath & res.Fields(1))
            lab(0).Caption = res.Fields(0)
        ElseIf printimg.npict = 2 Then
            Image1(0).Picture = LoadPicture(printimg.imgpath & res.Fields(1))
            lab(0).Caption = res.Fields(0)
            res.MovePrevious
            If res.EOF <> True And res.BOF <> True Then
                Image1(1).Picture = LoadPicture(printimg.imgpath & res.Fields(1))
                lab(1).Caption = res.Fields(0)
            End If
        ElseIf printimg.npict = 4 Then
            For i = 0 To 3
            If res.EOF <> True And res.BOF <> True Then
                Image1(i).Picture = LoadPicture(printimg.imgpath & res.Fields(1))
                lab(i).Caption = res.Fields(0)
                res.MovePrevious
            End If
            Next
        ElseIf printimg.npict = 6 Then
            For i = 0 To 5
            If res.EOF <> True And res.BOF <> True Then
                Image1(i).Picture = LoadPicture(printimg.imgpath & res.Fields(1))
                lab(i).Caption = res.Fields(0)
                res.MovePrevious
            End If
            Next
       End If
    End If
End Sub
