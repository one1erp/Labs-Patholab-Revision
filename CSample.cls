VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public strName As String
Public strId As String
Public strWorkflowName As String
Private strStatus As String

Public dicBlocks As Dictionary


Private Sub Class_Initialize()
'    Set dicBlocks = New Dictionary
End Sub

Public Sub Initialize(sampleName As String, sampleId As String)
6210      strName = sampleName
6220      strId = sampleId
          
6230      Set dicBlocks = New Dictionary
          
6240  If DEFINE_DEBUG Then MsgBox "Sample: " & strName & " " & strId

6250      Call InitWorkflowName
6260      Call FillBlockList
End Sub

Public Sub SetStatus(strStatus_ As String)
6270      strStatus = strStatus_
End Sub

Public Function GetStatus() As String
6280      GetStatus = strStatus
End Function

Public Sub UpdateStatus()
6290  On Error GoTo ERR_UpdateStatus

          Dim sql As String
          Dim i As Integer
          
6300      If strStatus = "X" Then
          
6310          sql = "  update lims_sys.sample  "
6320          sql = sql & " set status = '" & strStatus & "'  "
6330          sql = sql & " where sample_id = '" & strId & "'  "
              
6340          Call con.Execute(sql)
          
6350      End If
          
6360      For i = 0 To dicBlocks.Count - 1
6370          Call dicBlocks.Items(i).UpdateStatus
6380      Next i
          
6390      Exit Sub
ERR_UpdateStatus:
6400  MsgBox "ERR_UpdateStatus" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub


Private Sub InitWorkflowName()
6410      On Error GoTo ERR_InitWorkflowName
              
          Dim sql As String
          Dim rs As Recordset
              
              
6420      sql = "   select w.NAME  from"
6430      sql = sql & "  lims_sys.workflow_node wn, lims_sys.workflow w , lims_sys.sample s"
6440      sql = sql & "  where s.WORKFLOW_NODE_ID = wn.WORKFLOW_NODE_ID"
6450      sql = sql & "  and wn.WORKFLOW_ID=w.WORKFLOW_ID"
6460      sql = sql & "  and s.sample_id = " & strId
          
6470      Set rs = con.Execute(sql)
          
6480      If Not rs.EOF Then
6490          strWorkflowName = nte(rs("NAME"))
6500      End If
              
6510      Exit Sub
ERR_InitWorkflowName:
6520  MsgBox "ERR_InitWorkflowName" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub

Private Sub FillBlockList()
6530  On Error GoTo ERR_FillBlockList

          Dim block As CBlock
          Dim rs As Recordset
          
6540      Set rs = con.Execute _
                (" select a.name, a.aliquot_id " & _
                 " from lims_sys.aliquot a " & _
                 " where sample_id = " & strId & _
                 " and not exists " & _
                 " (select aliquot_id from lims_sys.aliquot_formulation " & _
                 " where child_aliquot_id = a.aliquot_id )" & _
                " order by TO_NUMBER( substr(name,INSTR( name,'.',1,2)+1)) " _
                 )
                 
6550      While Not rs.EOF
6560          Set block = New CBlock
6570          Call block.Initialize(nte(rs("name")), nte(rs("aliquot_id")))
              
6580          Call dicBlocks.Add(dicBlocks.Count, block)
6590          rs.MoveNext
6600      Wend
          
6610      Exit Sub
ERR_FillBlockList:
6620  MsgBox "FillBlockList" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub




