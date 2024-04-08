VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSdg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public strName As String
Public strId As String
Public strExternalReference As String
Public strWorkflowName As String

Public dicSamples As Dictionary


Private Sub Class_Initialize()
'    strId = CStr(sdg_id)
'    strName = sdg_name
    
'    Set dicSamples = New Dictionary
    
'If DEFINE_DEBUG Then MsgBox "sdg: " & strId & " " & strName
End Sub

Public Sub Initialize(sdgName As String, sdgId As String)
6630      strId = sdgId
6640      strName = sdgName
          
6650      Set dicSamples = New Dictionary
          
6660  If DEFINE_DEBUG Then MsgBox "sdg: " & strId & " " & strName

6670      Call GetExternalReferenceAndWorkflowName
6680      Call FillSampleList
End Sub


Public Sub UpdateStatus()
6690  On Error GoTo ERR_UpdateStatus

          Dim i As Integer
          
6700      For i = 0 To dicSamples.Count - 1
6710          Call dicSamples.Items(i).UpdateStatus
6720      Next i
          
6730      Exit Sub
ERR_UpdateStatus:
6740  MsgBox "ERR_UpdateStatus" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub


Private Sub GetExternalReferenceAndWorkflowName()
6750      On Error GoTo ERR_GetExternalReferenceAndWorkflowName
              
          Dim sql As String
          Dim rs As Recordset
              
6760      sql = " select w.NAME, d.EXTERNAL_REFERENCE from"
6770      sql = sql & " lims_sys.workflow_node wn, lims_sys.workflow w , lims_sys.sdg d"
6780      sql = sql & " where d.WORKFLOW_NODE_ID = wn.WORKFLOW_NODE_ID"
6790      sql = sql & " and wn.WORKFLOW_ID=w.WORKFLOW_ID"
6800      sql = sql & " and d.SDG_ID =  " & strId
          
6810      Set rs = con.Execute(sql)
          
6820      If Not rs.EOF Then
6830          strExternalReference = nte(rs("EXTERNAL_REFERENCE"))
6840          strWorkflowName = nte(rs("NAME"))
6850      End If
              
6860      Exit Sub
ERR_GetExternalReferenceAndWorkflowName:
6870  MsgBox "ERR_GetExternalReferenceAndWorkflowName" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub


Private Sub FillSampleList()
          Dim rs As Recordset
          Dim sample As CSample
          
6880      Set rs = con.Execute _
                       (" select name, sample_id " & _
                        " from lims_sys.sample " & _
                        " where sdg_id = " & strId & _
                        " order by " & _
                        "   TO_NUMBER(substr(name,INSTR( name,'.')+1))" _
)
                        
6890      While Not rs.EOF
6900          Set sample = New CSample
6910          Call sample.Initialize(nte(rs("name")), nte(rs("sample_id")))
              
6920          Call dicSamples.Add(dicSamples.Count, sample)
              
6930          rs.MoveNext
6940      Wend
End Sub




