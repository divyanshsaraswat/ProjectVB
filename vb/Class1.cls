VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents con1 As Connection
Attribute con1.VB_VarHelpID = -1
Public WithEvents con2 As Connection
Attribute con2.VB_VarHelpID = -1
Public WithEvents con3 As Connection
Attribute con3.VB_VarHelpID = -1
Private Sub Class_Initialize()
On Error Resume Next
Set con1 = New Connection
Set con2 = New Connection
Set con3 = New Connection

With con1
.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db2.mdb;Persist Security Info=False"

'.ConnectionString = "DSN=db2;;;Persist Security Info=False"
.Open
End With

With con2
.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db3.mdb;Persist Security Info=False"
'.ConnectionString = "DSN=db3;;;Persist Security Info=False"
.Open
End With

With con3
.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db4.mdb;Persist Security Info=False"
.Open
End With

End Sub
