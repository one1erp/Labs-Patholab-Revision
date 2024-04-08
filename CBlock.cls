VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public strName As String
Public strId As String
Public strWorkflowName As String
Private strStatus As String

Public dicSlides As Dictionary
Public dicTests As Dictionary

Private Sub Class_Initialize()
'    Set dicSlides = New Dictionary
'    Set dicTests = New Dictionary
End Sub


Public Sub Initialize(blockName As String, blockId As String)
5340      strName = blockName
5350      strId = blockId
          
5360      Set dicSlides = New Dictionary
5370      Set Me.dicTests = New Dictionary

5380  If DEFINE_DEBUG Then MsgBox "Block: " & strName & " " & strId

5390      Call InitWorkflowName
5400      Call FillSlidesList
5410      Call FillTestsList
End Sub

Public Sub SetStatus(strStatus_ As String)
5420      strStatus = strStatus_
End Sub

Public Function GetStatus() As String
5430      GetStatus = strStatus
End Function

Public Sub UpdateStatus()
5440  On Error GoTo ERR_UpdateStatus

          Dim sql As String
          Dim i As Integer
          
5450      If strStatus = "X" Then
          
5460          sql = "  update lims_sys.aliquot  "
5470          sql = sql & " set status = '" & strStatus & "'  "
5480          sql = sql & " where aliquot_id = '" & strId & "'  "
              
5490          Call con.Execute(sql)
          
5500      End If
          
5510      For i = 0 To dicSlides.Count - 1
5520          Call dicSlides.Items(i).UpdateStatus
5530      Next i
          
5540      For i = 0 To dicTests.Count - 1
5550          Call dicTests.Items(i).UpdateStatus
5560      Next i
          
5570      Exit Sub
ERR_UpdateStatus:
5580  MsgBox "ERR_UpdateStatus" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub

Private Sub InitWorkflowName()
5590      On Error GoTo ERR_InitWorkflowName
              
          Dim sql As String
          Dim rs As Recordset
              
5600      sql = "   select w.NAME  from"
5610      sql = sql & "  lims_sys.workflow_node wn, lims_sys.workflow w , lims_sys.aliquot a"
5620      sql = sql & "  where a.WORKFLOW_NODE_ID = wn.WORKFLOW_NODE_ID"
5630      sql = sql & "  and wn.WORKFLOW_ID=w.WORKFLOW_ID"
5640      sql = sql & "  and a.aliquot_id = " & strId
          
5650      Set rs = con.Execute(sql)
          
5660      If Not rs.EOF Then
5670          strWorkflowName = nte(rs("NAME"))
5680      End If
              
5690      Exit Sub
ERR_InitWorkflowName:
5700  MsgBox "ERR_InitWorkflowName" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub


Private Sub FillSlidesList()
5710  On Error GoTo ERR_FillSlidesList

          Dim slide As CSlide
          Dim rs As Recordset
          
5720      Set rs = con.Execute( _
                   " select a.name, a.aliquot_id " & _
                   " from lims_sys.aliquot a " & _
                   " where a.aliquot_id in " & _
                   " ( " & _
                   " select child_aliquot_id from lims_sys.aliquot_formulation " & _
                   " where parent_aliquot_id = " & strId & _
                   " ) " & _
                   "  order by TO_NUMBER( substr(name,INSTR( name,'.',1,3)+1)) " _
                   )
          
5730      While Not rs.EOF
5740          Set slide = New CSlide
5750          Call slide.Initialize(nte(rs("name")), nte(rs("aliquot_id")))
              
5760          Call dicSlides.Add(dicSlides.Count, slide)
              
5770          rs.MoveNext
5780      Wend
                   
5790      Exit Sub
ERR_FillSlidesList:
5800  MsgBox "ERR_FillSlidesList" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub


Private Sub FillTestsList()
5810  On Error GoTo ERR_FillTestsList

          Dim test As CTest
          Dim rs As Recordset
          
5820      Set rs = con.Execute( _
                   " select name, test_id " & _
                   " from lims_sys.test " & _
                   " where aliquot_id = " & strId & _
                   " order by test_id " _
                   )
          
5830      While Not rs.EOF
5840          Set test = New CTest
5850          Call test.Initialize(nte(rs("name")), nte(rs("test_id")))
              
5860          Call Me.dicTests.Add(dicTests.Count, test)
5870          rs.MoveNext
5880      Wend
          
5890      Exit Sub
ERR_FillTestsList:
5900  MsgBox "ERR_FillTestsList" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub













