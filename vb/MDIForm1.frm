VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "DAKSH"
   ClientHeight    =   6675
   ClientLeft      =   2385
   ClientTop       =   1635
   ClientWidth     =   5550
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Menu Master 
      Caption         =   "Master"
   End
   Begin VB.Menu Forms 
      Caption         =   "Stock"
      Begin VB.Menu main 
         Caption         =   "StockSheet"
      End
      Begin VB.Menu find 
         Caption         =   "Find Details With Sales"
      End
      Begin VB.Menu sm 
         Caption         =   "Jewellery Stock Summary"
      End
      Begin VB.Menu rsm 
         Caption         =   "Raw Stock Summry"
      End
      Begin VB.Menu Stde 
         Caption         =   "Stock Tally Report"
      End
   End
   Begin VB.Menu sales 
      Caption         =   "Sale"
      Begin VB.Menu sinvoice 
         Caption         =   "Sales Invoice"
      End
      Begin VB.Menu essummry 
         Caption         =   "Invoice Details"
      End
      Begin VB.Menu estsales 
         Caption         =   "Sale Sheet"
      End
      Begin VB.Menu esales 
         Caption         =   "Sale Sheet Details"
      End
      Begin VB.Menu sqt 
         Caption         =   "Sale Qutation."
      End
      Begin VB.Menu qutdetail 
         Caption         =   "Sales Quotation Details"
      End
   End
   Begin VB.Menu approval 
      Caption         =   "Approval"
      Index           =   2
      Begin VB.Menu appsh 
         Caption         =   "Approval Sheet"
         Index           =   0
      End
      Begin VB.Menu appdet 
         Caption         =   "App. Details"
         Index           =   1
      End
      Begin VB.Menu expinv 
         Caption         =   "Export Invoice"
      End
   End
   Begin VB.Menu Purcase 
      Caption         =   "Purchase/Sales"
      Begin VB.Menu rawpur 
         Caption         =   "Raw Purchase/Sales"
      End
      Begin VB.Menu pursale 
         Caption         =   "Purchase/Sales Details"
      End
   End
   Begin VB.Menu utl 
      Caption         =   "Utilities"
      Index           =   3
      Begin VB.Menu impdata 
         Caption         =   "Import"
      End
      Begin VB.Menu cat 
         Caption         =   "Catlog Printing"
      End
      Begin VB.Menu tag 
         Caption         =   "Tag Printing"
      End
      Begin VB.Menu addcontact 
         Caption         =   "Add New Contact"
      End
   End
   Begin VB.Menu exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public curcon As Class1
Public con1 As Connection
Public con2 As Connection
Public con3 As Connection
Public picturepath As String
Dim res1 As New ADODB.Recordset
Private Sub addcontact_Click()
load partydetail
partydetail.Show
End Sub
Public Function parserText(msg As String) As String
    Dim msgLength As Long
    Dim start1 As Integer
    Dim j As Integer
    Dim pos As Integer
    Dim msg1 As String
    
             start1 = 2
             msgLength = Len(msg)
             For j = 0 To msgLength - 3
                 pos = InStr(start1, msg, "'", vbTextCompare)
                 If (pos <> 0) Then
                  start1 = pos + 1
                  msg1 = Mid(msg, 1, pos - 1) + "\'"
                  msg = msg1 + Mid(msg, start1, msgLength - 3)
                  j = start1 + 1
                  start1 = start1 + 1
                 End If
             Next j
      parserText = msg
End Function
'Private Sub app_Click()
'load Approvel
'Approvel.Show
'End Sub
Private Sub enq_Click()
load Form6
Form6.Show
End Sub

Private Sub appdet_Click(Index As Integer)
load appdetails
appdetails.Show
End Sub
Private Sub appsh_Click(Index As Integer)
load Approvel
Approvel.Show
End Sub

Private Sub cat_Click()
load printimg
printimg.Show
End Sub

Private Sub esales_Click()
load dsails
dsails.Show
End Sub

Private Sub essummry_Click()
load saledetails
saledetails.Show
End Sub
Private Sub estsales_Click()
load Form3
Form3.Show
End Sub
Private Sub exit_Click()
Unload Me
End Sub
Private Sub fcost_Click()
load Form5
Form5.Show
End Sub

Private Sub expinv_Click()
load expinvoice
expinvoice.Show
End Sub

Private Sub find_Click()
load Form2
Form2.Show
End Sub

Private Sub findd_Click()
load ViewDetails
ViewDetails.Show
End Sub


Private Sub irs_Click(Index As Integer)
'Private Sub irs_Click()
load issueform
issueform.Show
'load issrec
'issrec.Show
End Sub

Private Sub impdata_Click()
load import
import.Show
End Sub

Private Sub main_Click()
load Form1
Form1.Show
End Sub

Private Sub MDIForm_Load()
Set curcon = New Class1
Set con1 = curcon.con1
Set con2 = curcon.con2
Set con3 = curcon.con3
'Set con3 = curcon.con3
Dim str1
str1 = "select chpicturepath from tblpicinfo"

res1.Open str1, con1, adOpenDynamic, adLockOptimistic
picturepath = res1.Fields(0)
res1.close
End Sub

Private Sub mdspeci_Click()
load metalspeci
metalspeci.Show
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
On Error Resume Next
curcon.con1.close
curcon.con2.close
'MsgBox ("Enter in the unload")
End Sub

Private Sub qutdetail_Click()
load qutdetails
qutdetails.Show
End Sub

Private Sub rawpur_Click()
load purchase
purchase.Show
End Sub

Private Sub rsm_Click()
Dim objExcel As Excel.Application, objBook As Excel.Workbook, objsheet As Excel.Worksheet
Dim str1
On Error Resume Next
Set objExcel = GetObject(, "Excel.Application") 'if excel already open you can use GetObject
If Err.Number Then
Err.Clear
End If

Set objExcel = CreateObject("Excel.Application") 'or CreateObject to open new Excel Application
Set objBook = objExcel.Workbooks.Open("c:\\daksh\vb\stockreport.xls")
Set objsheet = objBook.Worksheets(1)

Dim diawt As Double
Dim diaamt As Double

str1 = "select sum(ncode),sum(nweight1) as diawt,sum(nweight1*minrate1)as diaamt ,sum(nweight2+nweight3) as colwt,sum(nweight2*minrate2+nweight3*minrate3) as colamt "
str1 = str1 & " ,sum(nweight4) as metal, sum(nweight4*minrate4) as metalamt from tblcosting where chstate='Avi.' and chcontent2<>'Dia.'"
res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic

While (res1.EOF = False)
    diawt = res1!diawt
    diaamt = res1!diaamt
    ''objsheet.Cells(9, 2) = res1!diawt
    ''objsheet.Cells(9, 4) = res1!diaamt
    objsheet.Cells(9, 5) = res1!colwt
    objsheet.Cells(9, 7) = Round(res1!colamt, 0)
    objsheet.Cells(9, 8) = res1!metal
    objsheet.Cells(9, 10) = Round(res1!metalamt, 0)
    
res1.MoveNext
Wend
res1.close

str1 = "select sum(ncode),sum(nweight2) as diawt2,sum(nweight2*minrate2) as diaamt2 "
str1 = str1 & " from tblcosting where chstate='Avi.' and chcontent2='Dia.'"
res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic

While (res1.EOF = False)
    diawt = diawt + res1!diawt2
    diaamt = diaamt + res1!diaamt2
   res1.MoveNext
Wend
res1.close
 objsheet.Cells(9, 2) = diawt
 objsheet.Cells(9, 4) = Round(diaamt, 0)
diawt = 0
diaamt = 0

str1 = "select sum(ncode),sum(nweight1) as diawt,sum(nweight1*minrate1) as diaamt ,sum(nweight2+nweight3) as colwt,sum(nweight2*minrate2+nweight3*minrate3) as colamt "
str1 = str1 & " ,sum(nweight4) as metal, sum(nweight4*minrate4) as metalamt from tblsale where chcontent2<>'Dia.' "
res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic

While (res1.EOF = False)
    diawt = res1!diawt
    diaamt = res1!diaamt
    ''objsheet.Cells(10, 2) = res1!diawt
    ''objsheet.Cells(10, 4) = Round(res1!diaamt, 0)
    objsheet.Cells(10, 5) = res1!colwt
    objsheet.Cells(10, 7) = Round(res1!colamt, 0)
    objsheet.Cells(10, 8) = res1!metal
    objsheet.Cells(10, 10) = Round(res1!metalamt, 0)
res1.MoveNext
Wend
res1.close
str1 = "select sum(ncode),sum(nweight1) as diawt,sum(nweight1*minrate1) as diaamt ,sum(nweight2+nweight3) as colwt,sum(nweight2*minrate2+nweight3*minrate3) as colamt "
str1 = str1 & " ,sum(nweight4) as metal, sum(nweight4*minrate4) as metalamt from tblsale where chcontent2='Dia.' "
res1.Open str1, MDIForm1.con1, adOpenDynamic, adLockOptimistic
While (res1.EOF = False)
    diawt = diawt + res1!diawt
    diaamt = diaamt + res1!diaamt
res1.MoveNext
Wend
res1.close

objsheet.Cells(10, 2) = diawt
objsheet.Cells(10, 4) = Round(diaamt, 0)
    
diawt = 0
diaamt = 0

objBook.save
objExcel.Visible = True
objBook.PrintPreview
''objBook.CLOSE
' cell(1,1) means cell A1 ;
Set objsheet = Null
Set objBook = Null

End Sub

Private Sub sheet_Click()

End Sub

Private Sub sinvoice_Click()
load Invoice
Invoice.Show
End Sub

Private Sub sm_Click()
'load stocksmy
'stocksmy.Show
load stsmy
stsmy.Show
End Sub

Private Sub ssr_Click()
'i = 0
'With DataReport5.Sections("Section1")
 'i = i + 1
 '.Controls("Label10").Caption = "" & i
'End With
DataReport5.Show
'DataReport5.PrintReport True
End Sub

Private Sub sqt_Click()
load estinv
estinv.Show
End Sub

Private Sub Stde_Click()
DataReport3.Show
'DataReport3.PrintReport True
End Sub

Private Sub stv_Click()
load Form4
Form4.Show
End Sub

Private Sub tag_Click()
'load Form7
'Form7.Show
load printtag
printtag.Show
End Sub
