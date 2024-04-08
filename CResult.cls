VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public strName As String
Public strId As String
Public strWorkflowName As String
Private strStatus As String

Public Sub Initialize(resultName As String, resultId As String)
5910      strName = resultName
5920      strId = resultId
          
5930  If DEFINE_DEBUG Then MsgBox "Result: " & strName & " " & strId
          
5940      Call InitWorkflowName
End Sub


Public Sub SetStatus(strStatus_ As String)
5950      strStatus = strStatus_
End Sub

Public Function GetStatus() As String
5960      GetStatus = strStatus
End Function

Public Sub UpdateStatus()
5970  On Error GoTo ERR_UpdateStatus

          Dim sql As String
          
5980      If strStatus = "X" Then
          
5990          sql = "  update lims_sys.result  "
6000          sql = sql & " set status = '" & strStatus & "'  "
6010          sql = sql & " where result_id = '" & strId & "'  "
              
6020          Call con.Execute(sql)
          
6030      End If
          
6040      Exit Sub
ERR_UpdateStatus:
6050  MsgBox "ERR_UpdateStatus" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub


Private Sub InitWorkflowName()
6060      On Error GoTo ERR_InitWorkflowName
              
          Dim sql As String
          Dim rs As Recordset
              
          'get the name of the event that created this result:
6070      sql = " select name from lims_sys.workflow_node"
6080      sql = sql & " where workflow_node_id = "
6090      sql = sql & " ("
6100      sql = sql & "    select wn.PARENT_ID"
6110      sql = sql & "    from  lims_sys.workflow_node wn, lims_sys.result r"
6120      sql = sql & "    where r.WORKFLOW_NODE_ID = wn.WORKFLOW_NODE_ID"
6130      sql = sql & "    and r.result_id = " & strId
6140      sql = sql & " )"
          
6150      Set rs = con.Execute(sql)
          
6160      If Not rs.EOF Then
6170          strWorkflowName = nte(rs("NAME"))
6180      End If
              
6190      Exit Sub
ERR_InitWorkflowName:
6200  MsgBox "ERR_InitWorkflowName" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub


