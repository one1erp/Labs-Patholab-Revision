VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public strName As String
Public strId As String
Public strWorkflowName As String
Private strStatus As String

Public dicResults As Dictionary


Private Sub Class_Initialize()
'    Set dicResults = New Dictionary
End Sub


Public Sub Initialize(testName As String, testId As String)
7240      Set dicResults = New Dictionary
          
7250      strName = testName
7260      strId = testId
          
7270  If DEFINE_DEBUG Then MsgBox "Test: " & strName & " " & strId
          
7280      Call InitWorkflowName
7290      Call FillResultList
End Sub


Public Sub SetStatus(strStatus_ As String)
7300      strStatus = strStatus_
End Sub

Public Function GetStatus() As String
7310      GetStatus = strStatus
End Function

Public Sub UpdateStatus()
7320  On Error GoTo ERR_UpdateStatus

          Dim sql As String
          Dim i As Integer
          
7330      If strStatus = "X" Then
          
7340          sql = "  update lims_sys.test  "
7350          sql = sql & " set status = '" & strStatus & "'  "
7360          sql = sql & " where test_id = '" & strId & "'  "
              
7370          Call con.Execute(sql)
          
7380      End If
          
7390      For i = 0 To dicResults.Count - 1
7400          Call dicResults.Items(i).UpdateStatus
7410      Next i
          
7420      Exit Sub
ERR_UpdateStatus:
7430  MsgBox "ERR_UpdateStatus" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub


Private Sub InitWorkflowName()
7440      On Error GoTo ERR_InitWorkflowName
              
          Dim sql As String
          Dim rs As Recordset
              
7450      sql = "   select w.NAME  from"
7460      sql = sql & "  lims_sys.workflow_node wn, lims_sys.workflow w , lims_sys.test t"
7470      sql = sql & "  where t.WORKFLOW_NODE_ID = wn.WORKFLOW_NODE_ID"
7480      sql = sql & "  and wn.WORKFLOW_ID=w.WORKFLOW_ID"
7490      sql = sql & "  and t.test_id = " & strId
7500      sql = sql & "  and w.version_status = 'A'"
          
7510      Set rs = con.Execute(sql)
          
7520      If Not rs.EOF Then
7530          strWorkflowName = nte(rs("NAME"))
7540      End If
              
7550      Exit Sub
ERR_InitWorkflowName:
7560  MsgBox "ERR_InitWorkflowName" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub



Private Sub FillResultList()
7570  On Error GoTo ERR_FillResultList
          
          Dim result As CResult
          Dim rs As Recordset
          
7580      Set rs = con.Execute( _
                    " select name, result_id " & _
                    " from lims_sys.result " & _
                    " where test_id = " & strId & _
                    " order by result_id " _
                    )
                    
7590      While Not rs.EOF
7600          Set result = New CResult
7610          Call result.Initialize(nte(rs("name")), nte(rs("result_id")))
              'save the dicResults in that way
              'dicResults(key=name of the result,itme=the result itsalf)
7620              Call dicResults.Add(nte(rs("name")), result)
                  
7630          rs.MoveNext
7640      Wend
                     
7650      Exit Sub
ERR_FillResultList:
7660  MsgBox "ERR_FillResultList" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub






