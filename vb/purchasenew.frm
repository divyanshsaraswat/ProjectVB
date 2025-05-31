VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   4680
   LinkTopic       =   "Form8"
   ScaleHeight     =   8130
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsavenew 
      BackColor       =   &H00BAEFFA&
      Caption         =   "Save && New"
      Height          =   315
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton cmdclose 
      BackColor       =   &H00BAEFFA&
      Caption         =   "Close"
      Height          =   315
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6000
      Width           =   1215
   End
   Begin VB.ComboBox cmbThrough 
      Height          =   315
      Left            =   8400
      TabIndex        =   0
      Text            =   "Combo2"
      Top             =   1440
      Width           =   2775
   End
   Begin VB.CommandButton cmdexitsave 
      BackColor       =   &H00BAEFFA&
      Caption         =   "Save && Exit"
      Height          =   315
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00BAEFFA&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   6615
      Left            =   0
      TabIndex        =   4
      Top             =   840
      Width           =   11775
      Begin VB.ComboBox through 
         Height          =   315
         Left            =   8400
         TabIndex        =   46
         Text            =   "Combo1"
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtinvnno 
         Height          =   285
         Left            =   2280
         TabIndex        =   38
         Text            =   "Text1"
         Top             =   240
         Width           =   3615
      End
      Begin VB.ComboBox cmbpaymode 
         Height          =   315
         ItemData        =   "purchasenew.frx":0000
         Left            =   8400
         List            =   "purchasenew.frx":000D
         TabIndex        =   36
         Text            =   "Combo1"
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txtfbfob 
         Height          =   285
         Left            =   2280
         TabIndex        =   35
         Text            =   "Text5"
         Top             =   960
         Width           =   3615
      End
      Begin VB.Frame framegrid 
         BackColor       =   &H00BAEFFA&
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   120
         TabIndex        =   6
         Top             =   2280
         Width           =   11055
         Begin VB.CommandButton cmdremove 
            BackColor       =   &H00CEF8FF&
            Caption         =   "Remove"
            Height          =   255
            Left            =   10080
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   120
            Width           =   975
         End
         Begin VB.CommandButton cmdmodify 
            BackColor       =   &H00CEF8FF&
            Caption         =   "Modify"
            Height          =   255
            Left            =   9120
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton cmdadd 
            BackColor       =   &H00CEF8FF&
            Caption         =   "Add"
            Height          =   255
            Left            =   8160
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   120
            Width           =   855
         End
         Begin VB.TextBox txtOther 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6120
            TabIndex        =   10
            Text            =   "Text10"
            Top             =   2040
            Width           =   855
         End
         Begin VB.TextBox txtTotalAmount 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8400
            TabIndex        =   9
            Text            =   "Text14"
            Top             =   2040
            Width           =   975
         End
         Begin VB.TextBox txtTotalQty 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3960
            TabIndex        =   8
            Text            =   "Text13"
            Top             =   2040
            Width           =   1455
         End
         Begin VB.TextBox txtTotalItem 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            TabIndex        =   7
            Text            =   "Text12"
            Top             =   2040
            Width           =   1455
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid mgrid 
            Height          =   1215
            Left            =   240
            TabIndex        =   11
            Top             =   480
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   2143
            _Version        =   393216
            Cols            =   13
            BackColorFixed  =   12251130
            BackColorBkg    =   13564159
            _NumberOfBands  =   1
            _Band(0).Cols   =   13
         End
         Begin VB.Frame framelot 
            Appearance      =   0  'Flat
            BackColor       =   &H00BAEFFA&
            ForeColor       =   &H80000008&
            Height          =   1815
            Left            =   120
            TabIndex        =   15
            Top             =   0
            Visible         =   0   'False
            Width           =   10935
            Begin VB.ComboBox cmbcolor 
               Height          =   315
               Left            =   3240
               TabIndex        =   24
               Text            =   "Combo2"
               Top             =   480
               Width           =   1215
            End
            Begin VB.ComboBox cmbqgrade 
               Height          =   315
               Left            =   1800
               TabIndex        =   23
               Text            =   "Combo1"
               Top             =   480
               Width           =   1335
            End
            Begin VB.ComboBox cmbstonea 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   360
               TabIndex        =   22
               Text            =   "Combo15"
               Top             =   480
               Width           =   1335
            End
            Begin VB.TextBox txtamt 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   7800
               TabIndex        =   21
               Text            =   "Text20"
               Top             =   480
               Width           =   1575
            End
            Begin VB.TextBox txtrate 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   6120
               TabIndex        =   20
               Text            =   "Text21"
               Top             =   480
               Width           =   1095
            End
            Begin VB.CommandButton cmdsaveaddnew 
               BackColor       =   &H00CEF8FF&
               Caption         =   "Save && Add New"
               Height          =   255
               Left            =   2520
               Style           =   1  'Graphical
               TabIndex        =   19
               Top             =   1320
               Width           =   1695
            End
            Begin VB.CommandButton cmdsaveexit 
               BackColor       =   &H00CEF8FF&
               Caption         =   "Save && Exit"
               Height          =   255
               Left            =   4320
               Style           =   1  'Graphical
               TabIndex        =   18
               Top             =   1320
               Width           =   1575
            End
            Begin VB.CommandButton cmdcancel 
               BackColor       =   &H00CEF8FF&
               Caption         =   "Cancel"
               Height          =   255
               Left            =   6120
               Style           =   1  'Graphical
               TabIndex        =   17
               Top             =   1320
               Width           =   1215
            End
            Begin VB.TextBox txtmodwt 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   4560
               TabIndex        =   16
               Text            =   "Text16"
               Top             =   480
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label Label24 
               BackStyle       =   0  'Transparent
               Caption         =   "Color"
               Height          =   255
               Left            =   3240
               TabIndex        =   30
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label23 
               BackStyle       =   0  'Transparent
               Caption         =   "Quality/Grade"
               Height          =   255
               Index           =   1
               Left            =   1800
               TabIndex        =   29
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label14 
               BackStyle       =   0  'Transparent
               Caption         =   "Stone "
               Height          =   255
               Left            =   360
               TabIndex        =   28
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label19 
               BackStyle       =   0  'Transparent
               Caption         =   "Qty/Weight(CT)"
               Height          =   255
               Left            =   4560
               TabIndex        =   27
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label Label20 
               BackStyle       =   0  'Transparent
               Caption         =   "Rate"
               Height          =   255
               Left            =   6120
               TabIndex        =   26
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label21 
               BackStyle       =   0  'Transparent
               Caption         =   "Amount"
               Height          =   255
               Left            =   7800
               TabIndex        =   25
               Top             =   840
               Width           =   1095
            End
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Other"
            Height          =   255
            Index           =   0
            Left            =   5520
            TabIndex        =   34
            Top             =   2040
            Width           =   615
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Amount"
            Height          =   255
            Index           =   1
            Left            =   7200
            TabIndex        =   33
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Quantity"
            Height          =   255
            Index           =   1
            Left            =   2760
            TabIndex        =   32
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Item"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   31
            Top             =   2040
            Width           =   975
         End
      End
      Begin VB.ComboBox cmbbuyer 
         Height          =   315
         Left            =   2280
         TabIndex        =   5
         Text            =   "Combo4"
         Top             =   600
         Width           =   3615
      End
      Begin MSComCtl2.DTPicker dtpdate 
         Height          =   255
         Left            =   8400
         TabIndex        =   37
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   450
         _Version        =   393216
         Format          =   66453505
         CurrentDate     =   38715
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "'Voucher  No."
         Height          =   255
         Left            =   480
         TabIndex        =   44
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   255
         Left            =   6960
         TabIndex        =   43
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Party Name"
         Height          =   255
         Left            =   480
         TabIndex        =   42
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Through"
         Height          =   255
         Left            =   6960
         TabIndex        =   41
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Mode"
         Height          =   255
         Left            =   6960
         TabIndex        =   40
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "FB/FOB No."
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   39
         Top             =   960
         Width           =   1095
      End
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00BAEFFA&
      Caption         =   "DAKSH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   45
      Top             =   0
      Width           =   11775
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdadd_Click()
On Error Resume Next
    Me.framelot.Visible = True

End Sub

Private Sub cmdmodify_Click()
On Error Resume Next

       ''  cmbstonea = mgrid.TextMatrix(mgrid.row, 3)
       ''  cmbqgrade = mgrid.TextMatrix(mgrid.row, 4)
       ''  cmbcolor = mgrid.TextMatrix(mgrid.row, 5)
       ''  cmbqty = mgrid.TextMatrix(mgrid.row, 12)
       ''  txtmodwt = mgrid.TextMatrix(mgrid.row, 12)
       ''  txtrate = mgrid.TextMatrix(mgrid.row, 13)
       ''  txtamt = mgrid.TextMatrix(mgrid.row, 14)
       ''   cmdsaveaddnew.Caption = "Update && New"
      ''   cmdsaveexit.Caption = "Update && Exit"
      ''   framelot.Visible = True

End Sub
