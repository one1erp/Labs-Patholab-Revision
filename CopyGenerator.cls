VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CopyGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



Private ProcessXML As LSSERVICEPROVIDERLib.NautilusProcessXML
Private NtlsCon As LSSERVICEPROVIDERLib.NautilusDBConnection
Private NtlsUser As LSSERVICEPROVIDERLib.NautilusUser

Private rsSDG As Recordset

Private dicStopListFields As Dictionary

Private iRevisionNum As Integer
Private sdg_log As New SdgLog.CreateLog
Private sdg_log_desc As String
Private WorkFolder As String
Private strRevisionCause As String
Private isRevForLab As Boolean



Private dicCancelEntities As New Dictionary
Private RRC As RequestRemark.RequestRemarkCtrl




'get input data from the client application:
Public Sub Initialize(sp As LSSERVICEPROVIDERLib.NautilusServiceProvider, _
                      strSdgId As String, _
                      strRevisionCause_ As String)
                            
10    On Error GoTo ERR_Initialize
          
          Dim strCon As String
          Dim rstisRevForLab As Recordset
          Dim sql As String
20        isRevForLab = False
          
30        Set NtlsCon = sp.QueryServiceProvider("DBConnection")
40        Set NtlsUser = sp.QueryServiceProvider("User")
50        Set con = New ADODB.Connection
          
60        strCon = "Provider=OraOLEDB.Oracle" & _
                   ";Data Source=" & NtlsCon.GetServerDetails & _
                   ";User ID=" & NtlsCon.GetUsername & _
                   ";Password=" & NtlsCon.GetPassword
                   
        
          
          If NtlsCon.GetServerIsProxy Then
            strCon = "Provider=OraOLEDB.Oracle;Data Source=" & _
            NtlsCon.GetServerDetails & ";User id=/;Persist Security Info=True;"
          End If
          
                   
70        Call con.Open(strCon)
80        con.CursorLocation = adUseServer
90        Call con.Execute("SET ROLE LIMS_USER")
100       Call ConnectSameSession(CDbl(NtlsCon.GetSessionId))
          
          
          
110       Set ProcessXML = sp.QueryServiceProvider("ProcessXML")
              
120       Set rsSDG = con.Execute(" select *                              " & _
                                  " from lims_sys.sdg d,                  " & _
                                  "      lims_sys.sdg_user du             " & _
                                  " where   du.sdg_id = d.sdg_id          " & _
                                  " and     d.sdg_id = '" & strSdgId & "' ")
                                  
130       strRevisionCause = strRevisionCause_
          
      '    sql = " select 1 from lims_sys.phrase_entry "
      '    sql = sql & " where phrase_id = (select phrase_id from lims_sys.phrase_header where "
      '    sql = sql & " name = 'Revision for lab only') "
      '   ' sql = sql & " and replace(upper(phrase_description),' ','')=replace(upper('" & strRevisionCause & "'),' ','')"
      '    sql = sql & " and phrase_name='" & Trim(strRevisionCause) & "'"
      '    Set rstisRevForLab = con.Execute(sql)
      '
      '
      '    If Not rstisRevForLab.EOF Then isRevForLab = True

140       If Trim(strRevisionCause) = "40" Then isRevForLab = True

           
150       Call InitializeStopListFields
           
160       Set rsSDG = con.Execute(" select *                              " & _
                                  " from lims_sys.sdg d,                  " & _
                                  "      lims_sys.sdg_user du             " & _
                                  " where   du.sdg_id = d.sdg_id          " & _
                                  " and     d.sdg_id = '" & strSdgId & "' ")
          'Call CreateSdgCopy
          
          
170       Exit Sub
ERR_Initialize:
180   MsgBox "ERR_Initialize" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub


'must be called after Initialize was called
Public Sub Execute()
190       Call CreateSdgCopy
End Sub



'performs the actual copy.
'will replace the original function: runSDGCopy
Private Function CreateSdgCopy()
200   On Error GoTo ERR_CreateSdgCopy

          Dim sdgOriginal As New CSdg
          Dim sdgCopy As New CSdg
          Dim s As String

210       If Not isSdgValidForRevision(rsSDG("sdg_id")) Then
220           MsgBox " SDG status must be Authorized or Rejected "
230           Exit Function
240       End If

250       Call sdgOriginal.Initialize(rsSDG("name"), rsSDG("sdg_id"))
260       s = LoginNewSdg("Login SDG", sdgOriginal.strExternalReference & "xx", sdgOriginal.strWorkflowName)

270       Call sdgCopy.Initialize(sdgOriginal.strExternalReference, s)

280       Call UpdateSdgLevel(sdgOriginal, sdgCopy)

290       Call UpdateSampleLevel(sdgOriginal, sdgCopy)
          
          'after the tree is built, cancel the
          'entities that are canceled in the original tree:
300       Call sdgCopy.UpdateStatus
          
310       If isRevForLab Then
              ' make a fake letter that won't be sent to DR
320          ' Call CreateLetterControl(sdgOriginal, sdgCopy)
              
330           Load frmRemark
              ' Make a Normal remark
340           Set RRC = frmRemark.RequestRemarkCtrl
350           Call RRC.InitializeConnection(con)
360           Call RRC.GetOperatorId(NtlsUser.GetOperatorId)
370           Call RRC.GetsdgName(sdgCopy.strExternalReference)
380           Call RRC.AddNormalRemark("נוצר כרוויזיה לצורך המעבדה בלבד.")
390           Unload frmRemark
400           con.Execute ("update lims_sys.sdg set inspection_plan_id= null where name= '" & sdgCopy.strName & "'")
410
                  
            
420       End If
 
440    con.Execute ("update lims_sys.sdg set  completed_by=null, completed_on=null where name= '" & sdgCopy.strName & "'")
                              'roy
450       Exit Function
ERR_CreateSdgCopy:
460   MsgBox "ERR_CreateSdgCopy" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Function
Private Sub CreateLetterControl(sdgOriginal As CSdg, sdgCopy As CSdg)
470   On Error GoTo ERR_CreateLetterControl
          Dim sql As String
          Dim rst As Recordset
          Dim LetterID As String
          
480       sql = "select lims.sq_U_LETTER_control.nextval from sys.dual"
       
490       Set rst = con.Execute(sql)
500       LetterID = rst(0)
          
510       sql = "  INSERT INTO lims_sys.U_LETTER_control"
520       sql = sql & "              (NAME,U_LETTER_control_id, VERSION, version_status"
530       sql = sql & "              )"
540       sql = sql & "       VALUES ('" & sdgCopy.strName & " - " & Now & "' ," & LetterID & ", 1, 'A'"
550       sql = sql & "              )"
        
560       Set rst = con.Execute(sql)
           
570       sql = "  INSERT INTO lims_sys.U_LETTER_control_user"
580       sql = sql & "              (U_LETTER_control_id,U_GRP_CODE,u_SDG_ID,u_reported"
590       sql = sql & "              )"
600       sql = sql & "       VALUES ( " & LetterID & ",11," & sdgCopy.strId & ",'T'"
610       sql = sql & "              )"
          
620       Set rst = con.Execute(sql)
           
630       sql = " select * from lims_sys.U_LETTER_control_user " & _
                                       " where u_sdg_id = " & sdgOriginal.strId
640       Set rst = con.Execute(sql)
650       If Not rst.EOF Then
660           sql = "UpdateRecordById"
670          Call UpdateRecordById("U_LETTER_control_user", "U_LETTER_control_id", _
             LetterID, nte(rst("U_LETTER_control_id")), rst)
680       End If
690       Exit Sub
ERR_CreateLetterControl:
700   MsgBox "ERR_CreateLetterControl" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
710   MsgBox sql
End Sub

Private Sub UpdateSdgLevel(sdgOriginal As CSdg, sdgCopy As CSdg)
720   On Error GoTo ERR_UpdateSdgLevel

          Dim rs As Recordset
          Dim sql As String
          Dim U_PATHOLAB_NUMBER As String
            
          'decide on the right name for the old sdg
          '(there might have already been some revisions)
730       Set rs = con.Execute(" select count(*) from lims_sys.sdg " & _
              " where substr(name,1,10)  = '" & sdgOriginal.strName & "' " & _
              " and instr(name, 'V', 1) > 1 ")
              
740       iRevisionNum = ntz(rs(0))
750       iRevisionNum = iRevisionNum + 1
760       Set rs = con.Execute(" select U_PATHOLAB_NUMBER from lims_sys.sdg_USER " & _
              " where sdg_id  = " & sdgOriginal.strId)
770       U_PATHOLAB_NUMBER = nte(rs(0))
              
          'change the names in memory:
780       sdgCopy.strName = sdgOriginal.strName
790       sdgOriginal.strName = sdgOriginal.strName & "V" & CStr(iRevisionNum)
          
          
          'change the names and status in DB:
800       con.Execute (" update lims_sys.sdg set name = '" & sdgOriginal.strName & "' " & _
                               " where sdg_id = " & sdgOriginal.strId)
810       con.Execute (" update lims_sys.sdg_USER set  U_PATHOLAB_NUMBER = U_PATHOLAB_NUMBER ||" & "'V" & CStr(iRevisionNum) & "' " & _
                               " where sdg_id = " & sdgOriginal.strId)
                               
820       con.Execute (" update lims_sys.sdg set name = '" & sdgCopy.strName & "', " & _
                               " status = 'V' " & _
                               " where sdg_id = " & sdgCopy.strId)
                               
830     con.Execute (" update lims_sys.sdg_USER set  U_PATHOLAB_NUMBER ='" & U_PATHOLAB_NUMBER & "' " & _
                               " where sdg_id = " & sdgCopy.strId)
          
          
          'update all the NEW SDG fields in the DB according to sdgOriginal
840       Set rs = con.Execute(" select * from lims_sys.sdg " & _
                                       " where sdg_id = " & sdgOriginal.strId)
                                       
850       Call UpdateRecordById("sdg", "sdg_id", sdgCopy.strId, sdgOriginal.strId, rs)
                                       
                                       
860       Set rs = con.Execute(" select * from lims_sys.sdg_user " & _
                                       " where sdg_id = " & sdgOriginal.strId)
                                       
870       Call UpdateRecordById("sdg_user", "sdg_id", sdgCopy.strId, sdgOriginal.strId, rs)
          
          
          'update revision cause:
880       sql = "update lims_sys.sdg_user set u_revision_cause = '" & strRevisionCause & "' " _
              & "where sdg_id = '" & sdgOriginal.strId & "'"
890       con.Execute (sql)
          
          
          'update who is the last SDG update and who isn't:
900       con.Execute (" update lims_sys.sdg_user " & _
                               " set u_is_last_update = 'F' " & _
                               " where sdg_id = " & sdgOriginal.strId)
          
910       con.Execute (" update lims_sys.sdg_user " & _
                               " set u_is_last_update = 'T' " & _
                               " where sdg_id = " & sdgCopy.strId)
          
          
          'enter a record to the sdg log table:
920       Set sdg_log.con = con
930       sdg_log.Session = CDbl(NtlsCon.GetSessionId)
940       sdg_log_desc = ""
950       Call sdg_log.InsertLog(rsSDG("sdg_id"), "REV.UPD", sdg_log_desc)
          
960   Exit Sub
ERR_UpdateSdgLevel:
970   MsgBox "ERR_UpdateSdgLevel" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub


Private Sub UpdateSampleLevel(sdgOriginal As CSdg, sdgCopy As CSdg)
980   On Error GoTo ERR_UpdateSampleLevel

          Dim i As Integer
          Dim sampleCopy As CSample
          Dim sampleOriginal As CSample
          Dim rs As Recordset
          Dim originalName As String
           

990       For i = 0 To sdgOriginal.dicSamples.Count - 1
1000            Set sampleOriginal = sdgOriginal.dicSamples(i)

1010            originalName = sampleOriginal.strName
                
                ' rename the old sample before creating a new one, since the name will be the same.
1020            sampleOriginal.strName = sampleOriginal.strName & "V" & CStr(iRevisionNum)
1030            con.Execute (" update lims_sys.sample set name = '" & sampleOriginal.strName & "' " & _
                               " where sample_id = " & sampleOriginal.strId)
                               
1040            con.Execute (" update lims_sys.aliquot set name = name || '" & "V" & CStr(iRevisionNum) & "' " & _
                               " where sample_id = " & sampleOriginal.strId)
1050
1060          If Not sdgCopy.dicSamples.Exists(i) Then
                  'create the missing sample for the new sdg:
1070              Set sampleCopy = New CSample
                
1080             sampleCopy.strId = LoginNewSample("Login Sample", "sample_" & CStr(i), _
                                     sampleOriginal.strWorkflowName, sdgCopy.strId)

1090              con.Execute (" update lims_sys.sample " & _
                                       " set sdg_id = " & sdgCopy.strId & _
                                       " where sample_id = " & sampleCopy.strId)
                  
1100              Call sampleCopy.Initialize("sample_" & CStr(i), sampleCopy.strId)
                  
                  'add the sample to memory in location i using CSample.initialize():
1110              Set sdgCopy.dicSamples(i) = sampleCopy
                  
1120          Else
                 'change the names in memory:
1130

1140              Set sampleCopy = sdgCopy.dicSamples(i)
1150          End If

             
1160           sampleCopy.strName = originalName

              'change the names and status in DB:
1170
              
1180          con.Execute (" update lims_sys.sample set name = '" & sampleCopy.strName & "' " & _
                                   " where sample_id = " & sampleCopy.strId)
              
1190          con.Execute (" update lims_sys.sample set status = 'V' " & _
                                   " where sample_id = '" & sampleCopy.strId & "'" & _
                                   " and   status <> 'X' ")
                                   
      '        con.Execute (" update lims_sys.sample set name = '" & sampleCopy.strName & "', " & _
                                   " status = 'V' " & _
                                   " where sample_id = " & sampleCopy.strId)
                  

              
              'update all relevant fields in DB:
1200          Set rs = con.Execute(" select * from lims_sys.sample " & _
                                           " where sample_id = " & sampleOriginal.strId)
                                           
1210          Call sampleCopy.SetStatus(rs("status"))
                                           
1220          Call UpdateRecordById("sample", "sample_id", _
                                     sampleCopy.strId, sampleOriginal.strId, rs)
                                           

1230          Set rs = con.Execute(" select * from lims_sys.sample_user " & _
                                           " where sample_id = " & sampleOriginal.strId)
                                           
1240          Call UpdateRecordById("sample_user", "sample_id", _
                                     sampleCopy.strId, sampleOriginal.strId, rs)
              
              'update all blocks for this sample:
1250          Call UpdateBlocksForSample(sampleOriginal, sampleCopy)

1260      Next i
          


1270  Exit Sub
ERR_UpdateSampleLevel:
1280  MsgBox "ERR_UpdateSampleLevel" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub


Private Sub UpdateBlocksForSample(sampleOriginal As CSample, sampleCopy As CSample)
1290  On Error GoTo ERR_UpdateBlocksForSample

        Dim i As Integer
        Dim blockCopy As CBlock
        Dim blockOriginal As CBlock
        Dim rs As Recordset
        Dim originalName As String
        
      
          
1300      For i = 0 To sampleOriginal.dicBlocks.Count - 1

1310        Set blockOriginal = sampleOriginal.dicBlocks(i)
           
1320        originalName = blockOriginal.strName
             'change the name in memory:
1330         blockOriginal.strName = blockOriginal.strName & "V" & CStr(iRevisionNum)
             'change the name in DB:
            'change original block's name before logging in a new block,
            'for its name will be the same (Unique Constraint):
1340        con.Execute (" update lims_sys.aliquot " & _
                        "  set name = '" & blockOriginal.strName & "' " & _
                        "  where aliquot_id = " & blockOriginal.strId)
                        
1350          If Not sampleCopy.dicBlocks.Exists(i) Then
1360              Set blockCopy = New CBlock
                  
                  
1370              blockCopy.strId = _
                     LoginNewBlock("Login Block", "block__" & CStr(i), _
                                    blockOriginal.strWorkflowName, sampleCopy.strId)
                  
                  
1380              con.Execute (" update lims_sys.aliquot " & _
                                       " set sample_id = " & sampleCopy.strId & _
                                       " where aliquot_id = " & blockCopy.strId)
                  
1390              Call blockCopy.Initialize("block_" & CStr(i), blockCopy.strId)
                  
1400              Set sampleCopy.dicBlocks(i) = blockCopy
1410          Else
1420              Set blockCopy = sampleCopy.dicBlocks(i)
1430          End If
              
              'change the names in memory:
1440          blockCopy.strName = originalName
1450
              
              'change the names and status in DB:
                       
1460          con.Execute (" update lims_sys.aliquot set name = '" & blockCopy.strName & "' " & _
                                   " where aliquot_id = " & blockCopy.strId)
                                   
1470          con.Execute (" update lims_sys.aliquot set status = 'V' " & _
                                   " where aliquot_id = " & blockCopy.strId & _
                                   " and status <> 'X' ")
                                   
      '        con.Execute (" update lims_sys.aliquot set name = '" & blockCopy.strName & "', " & _
                                   " status = 'V' " & _
                                   " where aliquot_id = " & blockCopy.strId)
                  
              
              'update all relevant fields in DB:
1480          Set rs = con.Execute(" select * from lims_sys.aliquot " & _
                                           " where aliquot_id = " & blockOriginal.strId)
                                           
                                           
1490          Call blockCopy.SetStatus(nte(rs("status")))
                                           
1500          Call UpdateRecordById("aliquot", "aliquot_id", _
                                     blockCopy.strId, blockOriginal.strId, rs)
                                           
                                           
1510          Set rs = con.Execute(" select * from lims_sys.aliquot_user " & _
                                           " where aliquot_id = " & blockOriginal.strId)
                                           
1520          Call UpdateRecordById("aliquot_user", "aliquot_id", _
                                     blockCopy.strId, blockOriginal.strId, rs)
                   
                   
              'update all slides & tests for this block:
1530          Call UpdateTestsForBlock(blockOriginal, blockCopy)
1540          Call UpdateSlidesForBlock(blockOriginal, blockCopy)
1550      Next i
              
1560      Exit Sub
ERR_UpdateBlocksForSample:
1570  MsgBox "ERR_UpdateBlocksForSample" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub


Private Sub UpdateTestsForBlock(blockOriginal As CBlock, blockCopy As CBlock)
1580  On Error GoTo ERR_UpdateTestsForBlock
          
          Dim i As Integer
          Dim testOriginal As CTest
          Dim testCopy As CTest
          Dim rs As Recordset
          
1590  If DEFINE_DEBUG Then MsgBox "update tests for block" & vbCrLf & _
                                  "block original: " & blockOriginal.dicTests.Count & vbCrLf & _
                                  "block copy: " & blockCopy.dicTests.Count
          
1600      For i = 0 To blockOriginal.dicTests.Count - 1
1610          Set testOriginal = blockOriginal.dicTests(i)
              
1620          If Not blockCopy.dicTests.Exists(i) And Trim(testOriginal.strWorkflowName) <> "" Then
1630              Set testCopy = New CTest
                  
1640              Call LoginNewTest("Login Test", "test__" & CStr(i), _
                                    testOriginal.strWorkflowName, blockCopy.strId)

1650              testCopy.strId = GetMaxTest(blockCopy.strId)

1660              con.Execute (" update lims_sys.test " & _
                                       " set aliquot_id = " & blockCopy.strId & _
                                       " where test_id = " & testCopy.strId)

1670              Call testCopy.Initialize("test_" & CStr(i), testCopy.strId)

1680              Set blockCopy.dicTests(i) = testCopy
1690          ElseIf blockCopy.dicTests.Exists(i) Then
1700              Set testCopy = blockCopy.dicTests(i)
1710          Else
1720              GoTo NEXT_IN_LOOP
1730          End If
          
          'never login more tests than those created by creating the block:
      '    For i = 0 To blockCopy.dicTests.Count - 1
      '        Set testOriginal = blockOriginal.dicTests(i)
      '        Set testCopy = blockCopy.dicTests(i)
          
              'change the names in memory:
1740          testCopy.strName = testOriginal.strName
              
              'change the names and status in DB:

1750          con.Execute (" update lims_sys.test set name = '" & testCopy.strName & "' " & _
                                   " where test_id = " & testCopy.strId)
                                   
1760          con.Execute (" update lims_sys.test set status = 'V' " & _
                                   " where test_id = " & testCopy.strId & _
                                   " and status <> 'X' ")
                                   
      '        con.Execute (" update lims_sys.test set name = '" & testCopy.strName & "', " & _
                                   " status = 'V' " & _
                                   " where test_id = " & testCopy.strId)
              
              'update all relevant fields in DB:
1770          Set rs = con.Execute(" select * from lims_sys.test " & _
                                           " where test_id = " & testOriginal.strId)
                                           
1780          Call testCopy.SetStatus(nte(rs("status")))
                                           
1790          Call UpdateRecordById("test", "test_id", _
                                     testCopy.strId, testOriginal.strId, rs)
                                           
                                           
1800          Set rs = con.Execute(" select * from lims_sys.test_user " & _
                                           " where test_id = " & testOriginal.strId)
                                           
1810          Call UpdateRecordById("test_user", "test_id", _
                                     testCopy.strId, testOriginal.strId, rs)
          
              'update the results of each test:
1820          Call UpdateResultsForTest(testOriginal, testCopy)
NEXT_IN_LOOP:
1830      Next i
          
1840      Exit Sub
ERR_UpdateTestsForBlock:
1850  MsgBox "ERR_UpdateTestsForBlock" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub


Private Sub UpdateSlidesForBlock(blockOriginal As CBlock, blockCopy As CBlock)
1860  On Error GoTo ERR_UpdateSlidesForBlock
          
          Dim i As Integer
          Dim slideOriginal As CSlide
          Dim slideCopy As CSlide
          Dim rs As Recordset
          
1870      For i = 0 To blockOriginal.dicSlides.Count - 1
1880          Set slideOriginal = blockOriginal.dicSlides(i)
          
1890          If Not blockCopy.dicSlides.Exists(i) Then
1900              Set slideCopy = New CSlide
                  
1910              Call LoginNewSlide("Add Slide", blockCopy.strId)
1920              slideCopy.strId = GetMaxSlide(blockCopy.strId)
                  
      '            con.Execute (" update lims_sys.aliquot " & _
                                       " set aliquot_id = " & blockCopy.strId & _
                                       " where aliquot_id = " & slideCopy.strId)
                  
1930              Call slideCopy.Initialize("slide_" & CStr(i), slideCopy.strId)
                  
1940              Set blockCopy.dicSlides(i) = slideCopy
1950          Else
1960              Set slideCopy = blockCopy.dicSlides(i)
1970          End If
          
              'change the names in memory:
1980          slideCopy.strName = slideOriginal.strName
1990          slideOriginal.strName = slideOriginal.strName & "V" & CStr(iRevisionNum)
              
              'change the names and status in DB:
2000          con.Execute (" update lims_sys.aliquot set name = '" & slideOriginal.strName & "' " & _
                                   " where aliquot_id = " & slideOriginal.strId)
              
2010          con.Execute (" update lims_sys.aliquot set name = '" & slideCopy.strName & "' " & _
                                   " where aliquot_id = " & slideCopy.strId)
              
2020          con.Execute (" update lims_sys.aliquot set status = 'V' " & _
                                   " where aliquot_id = " & slideCopy.strId & _
                                   " and status <> 'X' ")
              
      '        con.Execute (" update lims_sys.aliquot set name = '" & slideCopy.strName & "', " & _
                                   " status = 'V' " & _
                                   " where aliquot_id = " & slideCopy.strId)
              
              'update all relevant fields in DB:
2030          Set rs = con.Execute(" select * from lims_sys.aliquot " & _
                                           " where aliquot_id = " & slideOriginal.strId)
                                           
2040          Call slideCopy.SetStatus(nte(rs("status")))
                                           
2050          Call UpdateRecordById("aliquot", "aliquot_id", _
                                     slideCopy.strId, slideOriginal.strId, rs)
                                           
                                           
2060          Set rs = con.Execute(" select * from lims_sys.aliquot_user " & _
                                           " where aliquot_id = " & slideOriginal.strId)
                                           
2070          Call UpdateRecordById("aliquot_user", "aliquot_id", _
                                     slideCopy.strId, slideOriginal.strId, rs)
          
              'update the results of each test:
2080          Call UpdateTestsForSlide(slideOriginal, slideCopy)
2090      Next i
          
2100      Exit Sub
ERR_UpdateSlidesForBlock:
2110  MsgBox "ERR_UpdateSlidesForBlock" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub

Private Sub UpdateTestsForSlide(slideOriginal As CSlide, slideCopy As CSlide)
2120  On Error GoTo ERR_UpdateTestsForSlide
          
          Dim i As Integer
          Dim testOriginal As CTest
          Dim testCopy As CTest
          Dim rs As Recordset
          
2130  If DEFINE_DEBUG Then MsgBox "update tests for slide" & vbCrLf & _
                                  "slide original: " & slideOriginal.dicTests.Count & vbCrLf & _
                                  "slide copy: " & slideCopy.dicTests.Count
                                  
2140      For i = 0 To slideOriginal.dicTests.Count - 1
2150          Set testOriginal = slideOriginal.dicTests(i)

2160          If Not slideCopy.dicTests.Exists(i) Then
2170              Set testCopy = New CTest
                  
2180              Call LoginNewTest("Login Test", "test__" & CStr(i), _
                                    testOriginal.strWorkflowName, slideCopy.strId)

2190              testCopy.strId = GetMaxTest(slideCopy.strId)
                  
2200              con.Execute (" update lims_sys.test " & _
                                       " set aliquot_id = " & slideCopy.strId & _
                                       " where test_id = " & testCopy.strId)

2210              Call testCopy.Initialize("test_" & CStr(i), testCopy.strId)

2220              Set slideCopy.dicTests(i) = testCopy
2230          Else
2240              Set testCopy = slideCopy.dicTests(i)
2250          End If
          
          'never login more tests than those created by creating the slide:
      '    For i = 0 To slideCopy.dicTests.Count - 1
      '        Set testOriginal = slideOriginal.dicTests(i)
      '        Set testCopy = slideCopy.dicTests(i)
              
              'change the names in memory:
2260          testCopy.strName = testOriginal.strName
              
              'change the names and status in DB:

2270          con.Execute (" update lims_sys.test set name = '" & testCopy.strName & "' " & _
                                   " where test_id = " & testCopy.strId)

2280          con.Execute (" update lims_sys.test set status = 'V' " & _
                                   " where test_id = " & testCopy.strId & _
                                   " and status <> 'X' ")

      '        con.Execute (" update lims_sys.test set name = '" & testCopy.strName & "', " & _
                                   " status = 'V' " & _
                                   " where test_id = " & testCopy.strId)
              
              'update all relevant fields in DB:
2290          Set rs = con.Execute(" select * from lims_sys.test " & _
                                           " where test_id = " & testOriginal.strId)
                                           
2300          Call testCopy.SetStatus(nte(rs("status")))
                                           
2310          Call UpdateRecordById("test", "test_id", _
                                     testCopy.strId, testOriginal.strId, rs)
                                           
                                           
2320          Set rs = con.Execute(" select * from lims_sys.test_user " & _
                                           " where test_id = " & testOriginal.strId)
                                           
2330          Call UpdateRecordById("test_user", "test_id", _
                                     testCopy.strId, testOriginal.strId, rs)
          
              'update the results of each test:
2340          Call UpdateResultsForTest(testOriginal, testCopy)
2350      Next i
          
2360      Exit Sub
ERR_UpdateTestsForSlide:
2370  MsgBox "ERR_UpdateTestsForSlide" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub


Private Sub UpdateResultsForTest(testOriginal As CTest, testCopy As CTest)
2380  On Error GoTo ERR_UpdateResultsForTest
          
          Dim i As Integer
          Dim resultOriginal As CResult
          Dim resultcopy As CResult
          Dim rs As Recordset
          Dim resultOriginalName As String
          
2390      For i = 0 To testOriginal.dicResults.Count - 1
2400          Set resultOriginal = testOriginal.dicResults.Items(i)
2410           resultOriginalName = resultOriginal.strName

              'If a result is not exsist then creat one
2420          If (i > testCopy.dicResults.Count - 1) Then
2430              Set resultcopy = New CResult
                  
2440              resultcopy.strId = LoginNewResult(resultOriginal.strWorkflowName, _
                                                     testCopy.strId)

                     'LoginNewResult("Login Result", "result__" & CStr(i), _
                                    resultOriginal.strWorkflowName, testCopy.strId)
                      
2450              Set rs = con.Execute( _
                            " select name" & _
                            " from lims_sys.result " & _
                            " where result_id  = " & resultcopy.strId)
                                
2460              resultOriginalName = rs("name")
                  
2470              con.Execute (" update lims_sys.result " & _
                                       " set test_id = " & testCopy.strId & _
                                       " where result_id = " & resultcopy.strId)
                  
2480              Call resultcopy.Initialize(resultOriginalName, resultcopy.strId)
2490              resultcopy.strName = resultOriginalName
2500              Set testCopy.dicResults(i) = resultcopy
2510          Else
              'by the name of the original result get the copy result for Update
2520              Set resultcopy = testCopy.dicResults(resultOriginalName)
2530          End If
          
              'change the names in memory:
2540          resultcopy.strName = resultOriginal.strName
              
              'change the names and status in DB:

2550          con.Execute (" update lims_sys.result set name = '" & resultcopy.strName & "' " & _
                                   " where result_id = " & resultcopy.strId)

2560          con.Execute (" update lims_sys.result set status = 'V' " & _
                                   " where result_id = " & resultcopy.strId & _
                                   " and status <> 'X' ")
                                   

      '        con.Execute (" update lims_sys.result set name = '" & resultcopy.strName & "', " & _
                                   " status = 'V' " & _
                                   " where result_id = " & resultcopy.strId)
              
              'update all relevant fields in DB:
2570          Set rs = con.Execute(" select * from lims_sys.result " & _
                                           " where result_id = " & resultOriginal.strId)
                                           
2580          Call resultcopy.SetStatus(nte(rs("status")))
                                           
2590          Call UpdateRecordById("result", "result_id", _
                                     resultcopy.strId, resultOriginal.strId, rs)
                                           
                                           
2600          Set rs = con.Execute(" select * from lims_sys.result_user " & _
                                           " where result_id = " & resultOriginal.strId)
                                           
2610          Call UpdateRecordById("result_user", "result_id", _
                                     resultcopy.strId, resultOriginal.strId, rs)
          
          
              'change the status in DB:
2620          con.Execute (" update lims_sys.result set " & _
                                   " status = 'C' " & _
                                   " where result_id = " & resultcopy.strId & _
                                   " and status <> 'X' ")
                                   
              
2630          Call UpdateRTFResultForResult(resultOriginal.strId, resultcopy.strId)
2640      Next i
          
2650      Exit Sub
ERR_UpdateResultsForTest:
2660  MsgBox "ERR_UpdateResultsForTest" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub

'update the record for the results that contain FREE TEXT
'and have a record in the rtf_result table
Private Sub UpdateRTFResultForResult(strIdOriginal As String, strIdCopy As String)
2670  On Error GoTo UpdateRTFResultForResult
          
          Dim sql As String
          Dim rs As Recordset
          
          'check if this result has a record in the rtf_result table:
2680      sql = " select * from lims_sys.rtf_result where rtf_result_id = " & strIdOriginal
2690      Set rs = con.Execute(sql)
2700      If rs.EOF Then
2710          Exit Sub
2720      End If
          
          'insert a new record with the ID of the copied result:
2730      sql = " insert into lims_sys.rtf_result"
2740      sql = sql & " (rtf_result_id)"
2750      sql = sql & " values (" & strIdCopy & ")"
2760      con.Execute (sql)
          
          'update the result value to be the same as the original result:
2770      sql = " update lims_sys.rtf_result set rtf_text = "
2780      sql = sql & " ("
2790      sql = sql & "    select rtf_text from lims_sys.rtf_result "
2800      sql = sql & "    where rtf_result_id = " & strIdOriginal
2810      sql = sql & " )"
2820      sql = sql & " where  rtf_result_id = " & strIdCopy
2830      con.Execute (sql)
          
2840      Exit Sub
UpdateRTFResultForResult:
2850  MsgBox "UpdateRTFResultForResult" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub


'updates the relevant table by the given id
'to the values held by the recordset
'strIdFieldName - the id field name for this table
'strIdTarget    - the id field value for the record to be modified
'strIdSource    - the id field value for the record to read from
Private Sub UpdateRecordById(strTableName As String, strIdFieldName As String, _
                             strIdTarget As String, strIdSource As String, _
                             rs As Recordset)
                                  
2860  On Error GoTo ERR_UpdateRecordById
          
          Dim i As Integer
          Dim strFieldName As String
          Dim varFieldValue As Variant
          Dim sql As String
          
         ' sql = ""
         ' sql = sql & " update lims_sys." & strTableName & " set "
          
2870      For i = 0 To rs.Fields.Count - 1
2880          strFieldName = rs.Fields(i).Name
2890          varFieldValue = rs.Fields(i).Value
              
2900          If shouldCopyField(strFieldName, strTableName, varFieldValue) = True Then  'Or _
      '           (UCase(strFieldName) = UCase("STATUS") And nte(varFieldValue) = "X") Then
                 
                 
                  'If GetFieldType(strFieldName, strTableName) <> "D" Then
        '          If shouldCopyDirectly(strFieldName, strTableName, varFieldValue) = False Then

        '              sql = " update lims_sys." & strTableName
        '              sql = sql & " set " & strFieldName & " = '" & rs.Fields(i) & "'"
        '              sql = sql & " where " & strIdFieldName & " = " & strIdTarget
        '          Else
                      'copy the data
                      'directly from the data base:
                      
2910                  sql = " update lims_sys." & strTableName
2920                  sql = sql & " set " & strFieldName & " = "
2930                  sql = sql & " ( "
2940                  sql = sql & "    select " & strFieldName
2950                  sql = sql & "    from lims_sys." & strTableName
2960                  sql = sql & "    where " & strIdFieldName & " = " & strIdSource
2970                  sql = sql & " ) "
2980                  sql = sql & " where " & strIdFieldName & " = " & strIdTarget
        '          End If
                  
2990              con.Execute (sql)
         '         sql = sql & " " & strFieldName & " = '" & varFieldValue & "',"
3000          End If
              
3010      Next i
          

          
         ' sql = Left(sql, Len(sql) - 1)
         ' sql = sql & " where " & strIdFieldName & " = " & strId

      'MsgBox sql

         ' con.Execute (sql)

3020  Exit Sub
ERR_UpdateRecordById:
3030  MsgBox "ERR_UpdateTableById" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
3040  MsgBox sql
End Sub


'decides weather this field should be copied or not:
'1. is its value NULL?
'2. is it a STOP-LIST field?
Private Function shouldCopyField(strFieldName As String, strTableName As String, _
                                 varFieldValue As Variant) As Boolean
          
3050  On Error GoTo ERR_shouldCopyField

          Dim rs As Recordset
          
3060      shouldCopyField = False
          
3070      If IsNull(varFieldValue) Then Exit Function
          
3080      If dicStopListFields.Exists(LCase(strFieldName)) Then Exit Function
          
3090      shouldCopyField = True
                 
3100      Exit Function
ERR_shouldCopyField:
3110  MsgBox "ERR_shouldCopyField" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Function


Private Function LoginNewSdg(EventName As String, _
                             strExternalReference As String, _
                             strWorkflowName As String) As String
                                   
3120  On Error GoTo ERR_LoginNewSdg

          Dim doc As New DOMDocument
          Dim res As New DOMDocument
          Dim xmlLogin As IXMLDOMElement
          Dim xmlSdg As IXMLDOMElement
          Dim xmlEL As IXMLDOMElement
          Dim e As IXMLDOMElement
          Dim element As IXMLDOMElement
          Dim XmlECreateByWorkflow As IXMLDOMElement
          Dim XmlEWFName As IXMLDOMElement
          Dim XmlExternal As IXMLDOMElement
          Dim XmlExternal2 As IXMLDOMElement
          Dim FileName As String
          Dim RetError As String


3130      Set xmlEL = doc.createElement("lims-request")
3140      Call doc.appendChild(xmlEL)

3150      Set e = doc.createElement("login-request")
3160      Call xmlEL.appendChild(e)
          
          'Set xmlLogin = doc.createElement("login-request")
          'Call e.appendChild(xmlLogin)
3170      Set xmlSdg = doc.createElement("SDG")
3180      Call e.appendChild(xmlSdg)
          
3190      Set XmlECreateByWorkflow = doc.createElement("create-by-workflow")
3200      Call xmlSdg.appendChild(XmlECreateByWorkflow)
          
3210      Set XmlEWFName = doc.createElement("workflow-name")
3220      XmlEWFName.Text = strWorkflowName
3230      Call XmlECreateByWorkflow.appendChild(XmlEWFName)
          
3240      Set XmlExternal = doc.createElement("EXTERNAL_REFERENCE")
3250      XmlExternal.Text = strExternalReference
          
3260      Call xmlSdg.appendChild(XmlExternal)


3270      Set XmlExternal2 = doc.createElement("U_PATHOLAB_NUMBER")
3280      XmlExternal2.Text = strExternalReference
          
3290      Call xmlSdg.appendChild(XmlExternal2)
          'XmlEWF.Text = OriginSDG.Fields(i)
          
          
          
          'Set element = doc.createElement("find-by-id")
          'element.Text = AliquotID
      'Call xmlSdg.appendChild(element)
          'Set element = doc.createElement("login-request")
          'Set element = doc.createElement("fire-event")
          'element.Text = "Login SDG"
          'Call xmlSdg.appendChild(xmlLogin)

       '   If Trim(WorkFolder) <> "" Then
             ' FileName = "C:\SdgCOPY_DOC1"
             ' Call xmlManager.SaveXmlFile(doc, FileName)
       '   End If

          'doc.Save ("c:\login new sdg.xml")

3300      RetError = ProcessXML.ProcessXMLWithResponse(doc, res)
3310      If Trim(RetError) <> "" Then
3320          MsgBox "Error occurred while trying process xml file. " & vbCrLf & _
                     "Event Name: " & EventName & vbCrLf & _
                     "Error: " & RetError, vbCritical, "Nautilus - Sdg Copy"
3330      End If
          
          'res.Save ("c:\login new sdg res.xml")

       '   If Trim(WorkFolder) <> "" Then
             ' FileName = "C:\SdgCopy_RES1"
             ' Call xmlManager.SaveXmlFile(res, FileName)
       '   End If

3340       LoginNewSdg = res.selectSingleNode("//SDG_ID").Text

3350      Exit Function
ERR_LoginNewSdg:
3360  MsgBox "ERR_LoginNewSdg" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Function


Private Function LoginNewSample(EventName As String, _
                                strExternalReference As String, _
                                strWorkflowName As String, _
                                sdgId As String) As String
                                   
3370  On Error GoTo ERR_LoginNewSample

          Dim doc As New DOMDocument
          Dim res As New DOMDocument
          Dim xmlSdg As IXMLDOMElement
          Dim xmlFind As IXMLDOMElement
          Dim xmlLogin As IXMLDOMElement
          Dim xmlSample As IXMLDOMElement
          Dim xmlEL As IXMLDOMElement
          Dim e As IXMLDOMElement
          Dim element As IXMLDOMElement
          Dim XmlECreateByWorkflow As IXMLDOMElement
          Dim XmlEWFName As IXMLDOMElement
          Dim XmlExternal As IXMLDOMElement
          Dim FileName As String
          Dim RetError As String


3380      Set xmlEL = doc.createElement("lims-request")
3390      Call doc.appendChild(xmlEL)

3400      Set e = doc.createElement("login-request")
3410      Call xmlEL.appendChild(e)
          
          'Set xmlLogin = doc.createElement("login-request")
          'Call e.appendChild(xmlLogin)
          
          
          'SDG:
3420      Set xmlSdg = doc.createElement("SDG")
3430      Call e.appendChild(xmlSdg)
          
          'find by id:
3440      Set xmlFind = doc.createElement("find-by-id")
3450      xmlFind.Text = sdgId
3460      Call xmlSdg.appendChild(xmlFind)
          
3470      Set xmlSample = doc.createElement("SAMPLE")
3480      Call xmlSdg.appendChild(xmlSample)
          
3490      Set XmlECreateByWorkflow = doc.createElement("create-by-workflow")
3500      Call xmlSample.appendChild(XmlECreateByWorkflow)
          
3510      Set XmlEWFName = doc.createElement("workflow-name")
3520      XmlEWFName.Text = strWorkflowName
3530      Call XmlECreateByWorkflow.appendChild(XmlEWFName)
'
'3430      Set XmlExternal = doc.createElement("EXTERNAL_REFERENCE")
'3440      XmlExternal.Text = strExternalReference
'
'3450      Call xmlSample.appendChild(XmlExternal)
          'XmlEWF.Text = OriginSDG.Fields(i)
          
          
          
          'Set element = doc.createElement("find-by-id")
          'element.Text = AliquotID
      'Call xmlSdg.appendChild(element)
          'Set element = doc.createElement("login-request")
          'Set element = doc.createElement("fire-event")
          'element.Text = "Login SDG"
          'Call xmlSdg.appendChild(xmlLogin)

       '   If Trim(WorkFolder) <> "" Then
             ' FileName = "C:\SdgCOPY_createSample_DOC1"
             ' Call xmlManager.SaveXmlFile(doc, FileName)
       '   End If

'3540      doc.Save ("c:\login new sample.xml")

3550      RetError = ProcessXML.ProcessXMLWithResponse(doc, res)
3560      If Trim(RetError) <> "" Then
3570          MsgBox "Error occurred while trying process xml file. " & vbCrLf & _
                     "Event Name: " & EventName & vbCrLf & _
                     "Error: " & RetError, vbCritical, "Nautilus - Sdg Copy"
3580      End If
          
'3590      res.Save ("c:\login new sample res.xml")

       '   If Trim(WorkFolder) <> "" Then
             ' FileName = "C:\SdgCopy_createSample_RES1"
             ' Call xmlManager.SaveXmlFile(res, FileName)
       '   End If

3600       LoginNewSample = res.selectSingleNode("//SAMPLE_ID").Text

3610      Exit Function
ERR_LoginNewSample:
3620  MsgBox "ERR_LoginNewSample" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Function


Private Function LoginNewBlock(EventName As String, _
                               strExternalReference As String, _
                               strWorkflowName As String, _
                               strSampleId As String) As String
                                   
3630  On Error GoTo ERR_LoginNewBlock

          Dim doc As New DOMDocument
          Dim res As New DOMDocument
          Dim xmlLogin As IXMLDOMElement
          Dim xmlSample As IXMLDOMElement
          Dim xmlFind As IXMLDOMElement
          Dim xmlBlock As IXMLDOMElement
          Dim xmlEL As IXMLDOMElement
          Dim e As IXMLDOMElement
          Dim element As IXMLDOMElement
          Dim XmlECreateByWorkflow As IXMLDOMElement
          Dim XmlEWFName As IXMLDOMElement
          Dim XmlExternal As IXMLDOMElement
          Dim FileName As String
          Dim RetError As String

3640      Set xmlEL = doc.createElement("lims-request")
3650      Call doc.appendChild(xmlEL)

3660      Set e = doc.createElement("login-request")
3670      Call xmlEL.appendChild(e)
          
          'Set xmlLogin = doc.createElement("login-request")
          'Call e.appendChild(xmlLogin)
          
          
          'sample:
3680      Set xmlSample = doc.createElement("SAMPLE")
3690      Call e.appendChild(xmlSample)
          
          'find by id:
3700      Set xmlFind = doc.createElement("find-by-id")
3710      xmlFind.Text = strSampleId
3720      Call xmlSample.appendChild(xmlFind)
          
3730      Set xmlBlock = doc.createElement("ALIQUOT")
3740      Call xmlSample.appendChild(xmlBlock)
          
3750      Set XmlECreateByWorkflow = doc.createElement("create-by-workflow")
3760      Call xmlBlock.appendChild(XmlECreateByWorkflow)
          
3770      Set XmlEWFName = doc.createElement("workflow-name")
3780      XmlEWFName.Text = strWorkflowName
3790      Call XmlECreateByWorkflow.appendChild(XmlEWFName)
          
3800      Set XmlExternal = doc.createElement("EXTERNAL_REFERENCE")
3810      XmlExternal.Text = strExternalReference
          
3820      Call xmlBlock.appendChild(XmlExternal)
          'XmlEWF.Text = OriginSDG.Fields(i)

       '   If Trim(WorkFolder) <> "" Then
             ' FileName = "C:\SdgCOPY_createBlock_DOC1"
             ' Call xmlManager.SaveXmlFile(doc, FileName)
       '   End If

          'doc.Save ("c:\login new block.xml")

3830      RetError = ProcessXML.ProcessXMLWithResponse(doc, res)
3840      If Trim(RetError) <> "" Then
3850          MsgBox "Error occurred while trying process xml file. " & vbCrLf & _
                     "Event Name: " & EventName & vbCrLf & _
                     "Error: " & RetError, vbCritical, "Nautilus - Sdg Copy"
3860      End If
          
          'res.Save ("c:\login new block res.xml")

       '   If Trim(WorkFolder) <> "" Then
             ' FileName = "C:\SdgCopy_createBlock_RES1"
             ' Call xmlManager.SaveXmlFile(res, FileName)
       '   End If

3870       LoginNewBlock = res.selectSingleNode("//ALIQUOT_ID").Text

3880      Exit Function
ERR_LoginNewBlock:
3890  MsgBox "ERR_LoginNewBlock" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Function


Private Function LoginNewTest(EventName As String, _
                              strExternalReference As String, _
                              strWorkflowName As String, _
                              strAliquotId As String)
                                   
3900  On Error GoTo ERR_LoginNewTest

          Dim doc As New DOMDocument
          Dim res As New DOMDocument
          Dim xmlLogin As IXMLDOMElement
          Dim xmlAliquot As IXMLDOMElement
          Dim xmlFind As IXMLDOMElement
          Dim xmlTest As IXMLDOMElement
          Dim xmlEL As IXMLDOMElement
          Dim e As IXMLDOMElement
          Dim element As IXMLDOMElement
          Dim XmlECreateByWorkflow As IXMLDOMElement
          Dim XmlEWFName As IXMLDOMElement
          Dim XmlExternal As IXMLDOMElement
          Dim FileName As String
          Dim RetError As String

3910      Set xmlEL = doc.createElement("lims-request")
3920      Call doc.appendChild(xmlEL)

3930      Set e = doc.createElement("login-request")
3940      Call xmlEL.appendChild(e)
          
          'Set xmlLogin = doc.createElement("login-request")
          'Call e.appendChild(xmlLogin)
          
          
          'sample:
3950      Set xmlAliquot = doc.createElement("ALIQUOT")
3960      Call e.appendChild(xmlAliquot)
          
          'find by id:
3970      Set xmlFind = doc.createElement("find-by-id")
3980      xmlFind.Text = strAliquotId
3990      Call xmlAliquot.appendChild(xmlFind)
          
4000      Set xmlTest = doc.createElement("TEST")
4010      Call xmlAliquot.appendChild(xmlTest)
          
4020      Set XmlECreateByWorkflow = doc.createElement("create-by-workflow")
4030      Call xmlTest.appendChild(XmlECreateByWorkflow)
          
4040      Set XmlEWFName = doc.createElement("workflow-name")
4050      XmlEWFName.Text = strWorkflowName
4060      Call XmlECreateByWorkflow.appendChild(XmlEWFName)
          
      '    Set XmlExternal = doc.createElement("EXTERNAL_REFERENCE")
      '    XmlExternal.Text = strExternalReference
      '    Call xmlTest.appendChild(XmlExternal)
          
          'XmlEWF.Text = OriginSDG.Fields(i)

       '   If Trim(WorkFolder) <> "" Then
            '  FileName = "C:\SdgCOPY_createTest_DOC1"
            '  Call xmlManager.SaveXmlFile(doc, FileName)
       '   End If

          'doc.Save ("c:\login new test.xml")

4070      RetError = ProcessXML.ProcessXMLWithResponse(doc, res)
4080      If Trim(RetError) <> "" Then
4090          MsgBox "Error occurred while trying process xml file. " & vbCrLf & _
                     "Event Name: " & EventName & vbCrLf & _
                     "Error: " & RetError, vbCritical, "Nautilus - Sdg Copy"
4100      End If
          
          'res.Save ("c:\login new test res.xml")

       '   If Trim(WorkFolder) <> "" Then
            '  FileName = "C:\SdgCopy_createTest_RES1"
            '  Call xmlManager.SaveXmlFile(res, FileName)
       '   End If

      '     LoginNewTest = res.selectSingleNode("//TEST_ID").Text

4110      Exit Function
ERR_LoginNewTest:
4120  MsgBox "ERR_LoginNewTest" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Function


Private Function LoginNewResult(EventName As String, _
                                strTestId As String) As String
                                   
4130  On Error GoTo ERR_LoginNewResult

          Dim doc As New DOMDocument
          Dim res As New DOMDocument
          Dim xmlLogin As IXMLDOMElement
          Dim xmlTest As IXMLDOMElement
          Dim xmlFind As IXMLDOMElement
      '    Dim xmlResult As IXMLDOMElement
          Dim xmlEL As IXMLDOMElement
          Dim e As IXMLDOMElement
          Dim element As IXMLDOMElement
      '    Dim XmlECreateByWorkflow As IXMLDOMElement
      '    Dim XmlEWFName As IXMLDOMElement
      '    Dim XmlExternal As IXMLDOMElement
          Dim xmlFireEvent As IXMLDOMElement
          
          Dim FileName As String
          Dim RetError As String

4140      Set xmlEL = doc.createElement("lims-request")
4150      Call doc.appendChild(xmlEL)

4160      Set e = doc.createElement("login-request")
4170      Call xmlEL.appendChild(e)
          
4180      Set xmlTest = doc.createElement("TEST")
4190      Call e.appendChild(xmlTest)
          
          'find by id:
4200      Set xmlFind = doc.createElement("find-by-id")
4210      xmlFind.Text = strTestId
4220      Call xmlTest.appendChild(xmlFind)
          
4230      Set xmlFireEvent = doc.createElement("fire-event")
4240      xmlFireEvent.Text = EventName
4250      Call xmlTest.appendChild(xmlFireEvent)


       '   If Trim(WorkFolder) <> "" Then
       '       FileName = "C:\SdgCOPY_createResult_DOC1"
       '       Call xmlManager.SaveXmlFile(doc, FileName)
       '   End If

          'doc.Save ("c:\login new result.xml")

4260      RetError = ProcessXML.ProcessXMLWithResponse(doc, res)
4270      If Trim(RetError) <> "" Then
4280          MsgBox "Error occurred while trying process xml file. " & vbCrLf & _
                     "Event Name: " & EventName & vbCrLf & _
                     "Error: " & RetError, vbCritical, "Nautilus - Sdg Copy"
4290      End If

          'res.Save ("c:\login new result res.xml")

       '   If Trim(WorkFolder) <> "" Then
       '       FileName = "C:\SdgCopy_createResult_RES1"
       '       Call xmlManager.SaveXmlFile(res, FileName)
       '   End If

4300       LoginNewResult = res.selectSingleNode("//RESULT_ID").Text

4310      Exit Function
ERR_LoginNewResult:
4320  MsgBox "ERR_LoginNewResult" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Function


Private Function LoginNewSlide(strEventName As String, strBlockId As String) 'As String
4330      On Error GoTo Err_LoginNewSlide
          
          Dim doc As New DOMDocument
          Dim res As New DOMDocument
          Dim xmlLogin As IXMLDOMElement
          Dim xmlBlock As IXMLDOMElement
          Dim e As IXMLDOMElement
          Dim element As IXMLDOMElement
          Dim FileName As String
          Dim RetError As String

4340      Set e = doc.createElement("lims-request")
4350      Call doc.appendChild(e)
4360      Set xmlLogin = doc.createElement("login-request")
4370      Call e.appendChild(xmlLogin)
4380      Set xmlBlock = doc.createElement("ALIQUOT")
4390      Call xmlLogin.appendChild(xmlBlock)
4400      Set element = doc.createElement("find-by-id")
4410      element.Text = strBlockId
4420      Call xmlBlock.appendChild(element)
4430      Set element = doc.createElement("fire-event")
4440      element.Text = strEventName
4450      Call xmlBlock.appendChild(element)

4460      If Trim(WorkFolder) <> "" Then
             ' FileName = "C:\SDgCopy_" & strEventName & "_" & strBlockId & "_DOC1"
             ' Call xmlManager.SaveXmlFile(doc, FileName)
4470      End If

          'doc.Save ("c:\login new slide.xml")


4480      RetError = ProcessXML.ProcessXMLWithResponse(doc, res)
4490      If Trim(RetError) <> "" Then
4500          MsgBox "Error occurred while trying process xml file. (TriggerSlideEvent) " & vbCrLf & _
                     "Block ID: " & strBlockId & vbCrLf & _
                     "Event Name: " & strEventName & vbCrLf & _
                     "Error: " & RetError, vbCritical, "Nautilus - Sdg Copy"
4510      End If

          'res.Save ("c:\login new slide res.xml")

4520      If Trim(WorkFolder) <> "" Then
             ' FileName = "SdgCopy_" & strEventName & "_" & strBlockId & "_RES1"
             ' Call xmlManager.SaveXmlFile(res, FileName)
4530      End If

          'LoginNewSlide = res.selectSingleNode("//ALIQUOT_ID").Text
4540      Exit Function

Err_LoginNewSlide:
4550      MsgBox "Err_LoginNewSlide" & vbCrLf & _
                  "Block ID = " & strBlockId & vbCrLf & _
                  "Event Name = " & strEventName & vbCrLf & _
                 "In Line #" & Erl & vbCrLf & Err.Description
End Function

'used for getting the latest slide created for this block
'after the login of a new slide (the login action doesn't return the slide id)
Private Function GetMaxSlide(ParentID As String) As String
4560  On Error GoTo Err_GetMaxSlide
          Dim strSql As String
          Dim SlideRec As ADODB.Recordset

4570      GetMaxSlide = 0
4580      strSql = "select max(a.aliquot_id) " & _
                   "from lims_sys.aliquot a " & _
                   "where a.aliquot_id in " & _
                      "(select child_aliquot_id from lims_sys.aliquot_formulation " & _
                  "where aliquot_formulation.parent_aliquot_id = '" & ParentID & "') " & _
                  "order by a.aliquot_id"
4590      Set SlideRec = con.Execute(strSql)

4600      If Not SlideRec.EOF Then
4610          GetMaxSlide = SlideRec(0)
4620      End If
4630      SlideRec.Close
4640      Exit Function
          
Err_GetMaxSlide:
4650      MsgBox "Err_GetMaxSlide... " & vbCrLf & _
                  "Parent Aliquot ID = " & ParentID & vbCrLf & _
                 "In Line #" & Erl & vbCrLf & Err.Description
End Function


Private Function GetMaxTest(strAliquotId As String) As String
4660  On Error GoTo ERR_GetMaxTest
          
          Dim rs As Recordset
          Dim sql As String
          
4670      sql = "  select max(t.test_id)"
4680      sql = sql & "  from lims_sys.test t"
4690      sql = sql & "  where t.ALIQUOT_ID = " & strAliquotId
          
4700      Set rs = con.Execute(sql)
4710      GetMaxTest = nte(rs(0))

4720      Exit Function
ERR_GetMaxTest:
4730  MsgBox "ERR_GetMaxTest" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Function

Private Function isSdgValidForRevision(strSdgId As String) As Boolean
          Dim rs As Recordset
          Dim sql As String
          
4740      isSdgValidForRevision = False
          
4750      sql = " select status "
4760      sql = sql & " from lims_sys.sdg "
4770      sql = sql & " where sdg_id = " & strSdgId
4780      sql = sql & " and status in ('R','A') "
          
4790      Set rs = con.Execute(sql)
          
4800      If Not rs.EOF Then
4810          isSdgValidForRevision = True
4820      End If
End Function


Private Sub ConnectSameSession(ByVal aSessionID)
4830  On Error GoTo ERR_ConnectSameSession

          Dim aProc As New ADODB.Command
          Dim aSession As New ADODB.Parameter
          
4840      aProc.ActiveConnection = con
4850      aProc.CommandText = "lims.lims_env.connect_same_session"
4860      aProc.CommandType = adCmdStoredProc

4870      aSession.Type = adDouble
4880      aSession.Direction = adParamInput
4890      aSession.Value = aSessionID
4900      aProc.Parameters.Append aSession

4910      aProc.Execute
4920      Set aSession = Nothing
4930      Set aProc = Nothing
          
4940      Exit Sub
ERR_ConnectSameSession:
4950  MsgBox "ERR_ConnectSameSession" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub


Private Sub InitializeStopListFields()
4960  On Error GoTo ERR_InitializeStopListFields
4970      Set dicStopListFields = New Dictionary
          
4980      Call dicStopListFields.Add("name", "name")
4990      Call dicStopListFields.Add("sdg_id", "sdg_id")
      '   Call dicStopListFields.Add("reported", "reported")
5000      Call dicStopListFields.Add("sample_id", "sample_id")
5010      Call dicStopListFields.Add("aliquot_id", "aliquot_id")
5020      Call dicStopListFields.Add("test_id", "test_id")
5030      Call dicStopListFields.Add("result_id", "result_id")
5040      Call dicStopListFields.Add("status", "status")
5050      Call dicStopListFields.Add("authorised_on", "authorised_on")
5060      Call dicStopListFields.Add("authorised_by", "authorised_by")
5070      Call dicStopListFields.Add("workflow_node_id", "workflow_node_id")
5080      Call dicStopListFields.Add("sdg_template_id", "sdg_template_id")
5090      Call dicStopListFields.Add("sample_template_id", "sample_template_id")
5100      Call dicStopListFields.Add("aliquot_template_id", "aliquot_template_id")
5110      Call dicStopListFields.Add("test_template_id", "test_template_id")
5120      Call dicStopListFields.Add("result_template_id", "result_template_id")
5130      Call dicStopListFields.Add("u_pdf_path", "u_pdf_path")
5140      Call dicStopListFields.Add("u_patholab_number", "u_patholab_number")
5150      Call dicStopListFields.Add("u_week_nbr", "u_week_nbr")
5160      Call dicStopListFields.Add("u_fax_email_sent_on", "u_fax_email_sent_on")

      '--------------------

5170      Call dicStopListFields.Add("old_status", "old_status")
5180      Call dicStopListFields.Add("completed_on", "completed_on")
5190      Call dicStopListFields.Add("completed_by", "completed_by")
5200      Call dicStopListFields.Add("previous_sample", "previous_sample")
5210      Call dicStopListFields.Add("plate_aliquot_type", "plate_aliquot_type")
5220      Call dicStopListFields.Add("container_type_id", "container_type_id")
5230      If isRevForLab Then
5240          Call dicStopListFields.Add("inspection_plan_id", "inspection_plan_id")
5250          Call dicStopListFields.Add("u_letter_control_id", "u_letter_control_id")
5260          Call dicStopListFields.Add("u_grp_code", "u_grp_code")
5270          Call dicStopListFields.Add("u_sdg_id", "u_sdg_id")
5280          Call dicStopListFields.Add("u_reported", "u_reported")
5290      End If

          
5300      Exit Sub
ERR_InitializeStopListFields:
5310  MsgBox "ERR_InitializeStopListFields" & vbCrLf & "In Line #" & Erl & vbCrLf & Err.Description
End Sub


