VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "VBALGRID6.OCX"
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   3195
   ClientLeft      =   2595
   ClientTop       =   2835
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin vbAcceleratorGrid6.vbalGrid vbalGrid1 
      Height          =   1815
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   3201
      GridLines       =   -1  'True
      GridLinesBox    =   -1  'True
      CheckBoxes      =   -1  'True
      MultipleCheck   =   0   'False
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Editable        =   -1  'True
      DisableIcons    =   -1  'True
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub Form_Load()
vbalGrid1.Header = True
vbalGrid1.Rows = 10


'If (str1 = "") Then
'str1 = "SELECT tblappmaster.chissue, tblappdetail.chcategory, tblappdetail.*, [tblappdetail].[nrate1]*[tblappdetail].[nweight1]+[tblappdetail].[nrate2]*[tblappdetail].[nweight2]+[tblappdetail].[nrate3]*[tblappdetail].[nweight3]+[tblappdetail].[nrate4]*[tblappdetail].[nweight4]+[tblappdetail].[nmaking1] AS total, (total * 1.25)/100 AS priceno,tblappmaster.nappno, Tblcosting.chowner "
'str1 = str1 & " FROM Tblcosting RIGHT JOIN (tblappmaster INNER JOIN tblappdetail ON tblappmaster.nappno = tblappdetail.nappno) ON Tblcosting.ncode = tblappdetail.ncode order by chissue,tblappmaster.nappno"
'End If
'res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
'Printer.CurrentY = 160
'Printer.FontSize = 10
''Printer.PaperSize = vbPRPSEnvDL
'
'I = 0
'While (res1.EOF = False)
''If ((I = 0) Or (I = 6) Or (I = 11) Or (I = 16) Or (I = 21) Or (I = 26) Or (I = 31) Or (I = 36) Or (I = 41) Or (I = 46) Or (I = 51) Or (I = 56) Or (I = 61) Or (I = 66) Or (I = 71)) Then
''Printer.CurrentY = Printer.Height / 45
'
''End If
''Printer.Print (res1!chcategory & "-" & res1!ncode & "/" & Round(res1!priceno)); Spc(1)
''I = I + 1
'res1.MoveNext
'Wend
End Sub

