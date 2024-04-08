VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSlide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public strName As String
Public strId As String
Private strStatus As String

Public dicTests As Dictionary


Private Sub Class_Initialize()
'    Set dicTests = New Dictionary
End Sub


Public Sub Initialize(slideName As String, slideId As String)
6950      Set Me.dicTests = New Dictionary
          
6960      strName = slideName
6970      strId = slideId
          
6980  If DEFINE_DEBUG Then MsgBox "Slide: " & strName & " " & strId
          
6990      Call FillTestList
End Sub

Public Sub SetStatus(strStatus_ As String)
7000      strStatus = strStatus_
End Sub

Public Function GetStatus() As String
7010      GetStatus = strStatus
End Function

Public Sub UpdateStatus()
7020  On Error GoTo ERR_UpdateStatus

          Dim sql As String
          Dim i As Integer
          
7030      If strStatus = "X" Then
          
7040          sql = "  update lims_sys.aliquot  "
7050          sql = sql & " set status = '" & strStatus & "'  "
7060          sql = sql & " where aliquot_id = '" & strId & "'  "
              
7070          Call con.Execute(sql)
          
7080      End If
          
7090      For i = 0 To dicTests.Count - 1
7100          Call dicTests.Items(i).UpdateStatus
7110      Next i
          
7120      Exit Sub
ERR_UpdateStatus:
7130  MsgBox "ERR_UpdateStatus" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub


Private Sub FillTestList()
7140  On Error GoTo ERR_FillTestsList

          Dim test As CTest
          Dim rs As Recordset
          
7150      Set rs = con.Execute( _
                   " select name, test_id " & _
                   " from lims_sys.test " & _
                   " where aliquot_id = " & strId & _
                   " order by test_id " _
                   )
          
7160      While Not rs.EOF
7170          Set test = New CTest
7180          Call test.Initialize(nte(rs("name")), nte(rs("test_id")))
              
7190          Call Me.dicTests.Add(dicTests.Count, test)
7200          rs.MoveNext
7210      Wend
          
7220      Exit Sub
ERR_FillTestsList:
7230  MsgBox "ERR_FillTestsList" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description

End Sub




