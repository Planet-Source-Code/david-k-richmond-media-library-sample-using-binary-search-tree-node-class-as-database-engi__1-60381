Attribute VB_Name = "MediaLibarySample"
'---------------------------------------------------------------------------------------
' Module    : MediaLibarySample
' DateTime  : 29/04/2005 09:03
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' BUG FIX   :
' Version   :
' Details   :
'
'
' Other changes:
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'
' Purpose:
'        This is a beta media library sample using the BSTN class as it's database engine.
'        Other p-s-c visitors offered their help in testing the beta sample, checking the adds, removes,
'        ordered lists etc. and it's overall general usage.
'        The sample maintains details of media such as CDs, DVDs, Audio Tapes, VHS Videos etc.
'
'---------------------------------------------------------------------------------------

' Project   : Media Library
' DateTime  : 28/04/2005 20:37
' Author    : D K Richmond
' Status    : BETA RELEASE 1.0
' Version   : Visual Basic 6 (SP6)
' User Type : Beginner/Intermediate/Advanced
' Applicable: Database/Data Sorting/Key Retrieval/Binary Trees/Node Traversals
'
' Credits :    Chan Wilson - for all extra beta testing
' ---------
'
' Disclaimer:
' -----------
' This software is provided as-is, no liability accepted in using
' this code or part of this code where information is used and relied upon in any
' other system or systems, projects, modules etc.
' In other words, it is provided for educational value and where experimenting
' in writing applications needing fast ordered key retrieval for testing and
' running experimental non-critical non-business applications.
'
' Copyright:
' -----------
' This was written by myself based on information from various Data Abstraction
' and storage techniques.  Some of the deletion routines are based on information (only information not code)
' from various pages on the internet (non-copyright).
' There are some excellent articles on the internet, try google.com .. searching for "Binary Trees"
'
' ALL OF THIS CODE IN THIS CLASS WAS WRITTEN FROM A BLANK WORKSHEET.  IN OTHER WORDS NONE OF THIS CODE
' WAS STRIPPED, EXTRACTED OR COPIED FROM OTHER PEOPLES WORK.  PLEASE RESPECT THAT AND
' IF YOU USE THIS CODE IN ANY OF YOUR APPLICATIONS/EXAMPLES PLEASE GIVE CREDIT FOR MY WORK.
'
' DAVID K RICHMOND dk.richmond@ntlworld.com
'
' GUIDELINES:
' -----------
'
' OVERVIEW:
' -----------
'
' DEBUGGING:
' ---------
' In BSTN Class Initialise :  gbDebug = False   '  change this to true to see lots of debug message boxes
'
' SURPRISES:
'  I have left some msgbox's and some redundant error trapping in some routines, this will be taken out real soon.
'  The are some extra definitions that I have left for the future.

' FUTURE:
' -------
' There is a lot of scope to improve this code.
'
' Nomenclature (normal):
' -------------
'      i - integer prefix
'      l - long prefix
'      d - double prefix
'      s - string prefix
'      e - enum
'      c - original class prefix
'      cls - instance of a class prefix
'
'      m_i - member integer prefix
'      m_l - member long prefix
'      m_d - member double prefix
'      m_s - member string prefix
'      m_e - member enum
'      m_cls - member instance of a class prefix
'
' Nomenclature (not normal):
' -------------------------
'      ip - parameter integer prefix
'      lp - parameter long prefix
'      dp - parameter double prefix
'      sp - parameter string prefix
'
'      ipo - parameter optional integer prefix
'      lpo - parameter optional long prefix
'      dpo - parameter optional double prefix
'      spo - parameter optional string prefix
'
' ##########################################################################################################

Option Explicit

' ### globals declarations ##
 Global Const sgcLibraryFileName = "MediaLibrary.dat"
 Global sgLibraryFileName As String
 Global sgLibraryFilePathName As String
'
' the record data store
 Global vgaLibraryRecords() As Variant
 Global vgaLibraryRecordsShadow() As Variant
'
 Global lgLibraryRecord As Long
' global check for records added, update etc to prompt save library
 Global bgDirtyFlag As Boolean
' global key field identifier - default is 0
 Global igKeyFieldNumber As Integer
' mulit-dimensional array params
 Global igKeyFields As Integer
' how record fields are divided
 Global Const sgcFieldDelimiter = "|"

' BSTN class storage global
 Global clsgBSTStore As cBSTNStorage

' ### globals declarations ##

'---------------------------------------------------------------------------------------
' Procedure : InitialiseLibrary
' DateTime  : 29/04/2005 09:03
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function InitialiseLibrary(Optional ByVal ipoKeyFieldNumber As Variant)

   On Error GoTo InitialiseLibrary_Error

    Set clsgBSTStore = Nothing
     
    If Not IsMissing(ipoKeyFieldNumber) Then
        igKeyFieldNumber = ipoKeyFieldNumber
    Else
        igKeyFieldNumber = 0
    End If
    
    Erase vgaLibraryRecords()
    lgLibraryRecord = -1
    bgDirtyFlag = False
    
    Set clsgBSTStore = New cBSTNStorage

   On Error GoTo 0
   Exit Function

InitialiseLibrary_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure InitialiseLibrary of Module MediaLibarySample"
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : NewLibrary
' DateTime  : 29/04/2005 09:45
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function NewLibrary() As Boolean

    Set clsgBSTStore = Nothing
    Call InitialiseLibrary
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : LoadLibrary
' DateTime  : 29/04/2005 09:03
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function LoadLibrary(ByRef lblpLoadingStatus As Label) As Boolean
 Dim sLibraryFilePathName As String
 Dim iInChan As Integer
 Dim iKeyFields As Integer
 Dim iKeyField As Integer
' Dim vaLibraryRecord() As Variant
 Dim sLibRecordLine As String
 Dim bOk As Boolean
 Dim sKeyValue As String
 Dim saLibraryRecord() As String
 Dim lNodeIndex As Long
 
    LoadLibrary = False
    
   On Error GoTo LoadLibrary_Error
    
    lgLibraryRecord = 0
    iInChan = FreeFile()
    Open sgLibraryFilePathName For Input As #iInChan
    While Not EOF(iInChan)
        Line Input #iInChan, sLibRecordLine
        saLibraryRecord() = Split(sLibRecordLine, sgcFieldDelimiter)
        iKeyFields = UBound(saLibraryRecord())
         ' update global to track multi-dimensional array largest field
        If igKeyFields < iKeyFields Then
            igKeyFields = iKeyFields
        End If
        lgLibraryRecord = lgLibraryRecord + 1
        lblpLoadingStatus.Caption = "Scanning Media Record .." & lgLibraryRecord
        lblpLoadingStatus.Refresh
    Wend
    Close #iInChan

    ReDim vgaLibraryRecords(lgLibraryRecord, igKeyFields)
    lgLibraryRecord = 0
    iInChan = FreeFile()
    Open sgLibraryFilePathName For Input As #iInChan
    While Not EOF(iInChan)
        Line Input #iInChan, sLibRecordLine
        saLibraryRecord = Split(sLibRecordLine, sgcFieldDelimiter)
        iKeyFields = UBound(saLibraryRecord())
        
        For iKeyField = 0 To iKeyFields
            If iKeyField <= igKeyFields Then
                vgaLibraryRecords(lgLibraryRecord, iKeyField) = RemoveQuotes(saLibraryRecord(iKeyField))
            Else
                vgaLibraryRecords(lgLibraryRecord, iKeyField) = ""
            End If
        Next
        sKeyValue = vgaLibraryRecords(lgLibraryRecord, igKeyFieldNumber)
        If lgLibraryRecord = 0 Then
            bOk = clsgBSTStore.InitRootNode(sKeyValue)
            lNodeIndex = 0
            bOk = True  ' need to sort out InitRootNode return value
        Else
            bOk = clsgBSTStore.InsertBSTN(sKeyValue, lNodeIndex)
        End If
        If bOk = True Then
            ' set the data index value of the node we inserted to the index of the data array to use to get data from key finds
            bOk = clsgBSTStore.SetDataNodeIndex(lNodeIndex, lgLibraryRecord)
        End If
        If bOk = False Then
            MsgBox "Error Load Library failed setting key/index for value: " & sKeyValue, vbCritical
            Exit Function
        End If
        lgLibraryRecord = lgLibraryRecord + 1
        lblpLoadingStatus.Caption = "Loading Media Record .. " & lgLibraryRecord
        lblpLoadingStatus.Refresh
    Wend
    Close #iInChan
    
    ' final status
    lblpLoadingStatus.Caption = "Loaded (" & lgLibraryRecord & ") Records." & " on : " & Format(Now, "ddd-mmm-yyyy @ hh:mm")
    lblpLoadingStatus.Refresh

     LoadLibrary = True
     
   On Error GoTo 0
   Exit Function

LoadLibrary_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure LoadLibrary of Module MediaLibarySample"
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : SaveLibrary
' DateTime  : 29/04/2005 09:18
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function SaveLibrary(ByRef lblpLoadingStatus As Label) As Boolean
 Dim iOutChan As Integer
 Dim lLibraryRecord As Long
 Dim iKeyField As Integer
 Dim sFieldDelimiter As String
 
    SaveLibrary = False
 
   On Error GoTo SaveLibrary_Error

    iOutChan = FreeFile()
    Open sgLibraryFilePathName For Output As #iOutChan
    For lLibraryRecord = 0 To lgLibraryRecord - 1
        If (vgaLibraryRecords(lLibraryRecord, igKeyFieldNumber) <> "") Then
            sFieldDelimiter = sgcFieldDelimiter
            For iKeyField = 0 To igKeyFields
                ' dont keep putting extra comma on records
                If iKeyField = igKeyFields Then
                    sFieldDelimiter = ""
                End If
                Print #iOutChan, vgaLibraryRecords(lLibraryRecord, iKeyField) & sFieldDelimiter;
            Next
            Print #iOutChan, ""  ' crlf
            lblpLoadingStatus.Caption = "Saving Record .. " & lLibraryRecord
            lblpLoadingStatus.Refresh
        Else
            Debug.Print "not saving blank record @ " & lLibraryRecord
        End If
    Next
    Close #iOutChan
    
    ' final status
    lblpLoadingStatus.Caption = "Saved (" & lgLibraryRecord & ") Records." & " on : " & Format(Now, "ddd-mmm-yyyy @ hh:mm")
    lblpLoadingStatus.Refresh

    SaveLibrary = True
    
   On Error GoTo 0
   Exit Function

SaveLibrary_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SaveLibrary of Module MediaLibarySample"

End Function


'---------------------------------------------------------------------------------------
' Procedure : FindMedia
' DateTime  : 29/04/2005 10:27
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function FindMedia(ByVal spKeyValue As String, ByRef lpDataIndex As Long)
 Dim bOk As Boolean
 Dim lFoundNode As Long
 
    lFoundNode = -1
    lpDataIndex = -1
    
    bOk = clsgBSTStore.SearchBSTN(spKeyValue, lFoundNode)

    If bOk = True Then
        lpDataIndex = clsgBSTStore.GetDataNodeIndex(lFoundNode)
    End If

FindMedia = bOk
End Function


'---------------------------------------------------------------------------------------
' Procedure : DeleteMedia
' DateTime  : 03/05/2005 18:39
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function DeleteMedia(ByVal spKeyValue As String, ByRef lpDataIndex As Long)
 Dim bOk As Boolean
 Dim lFoundNode As Long
 
   On Error GoTo DeleteMedia_Error

    lFoundNode = -1
    lpDataIndex = -1
    
    bOk = clsgBSTStore.SearchBSTN(spKeyValue, lFoundNode)

    If bOk = True Then
        lpDataIndex = clsgBSTStore.GetDataNodeIndex(lFoundNode)
        bOk = RemoveMediaDataRecord(lpDataIndex)   ' don't compress the library records or all data indexes are voided
        If bOk = True Then
            bOk = clsgBSTStore.RemoveNode(spKeyValue)
            If bOk = True Then
                MsgBox "Removed key from library store for key : " & spKeyValue, vbCritical
            Else
                MsgBox "Failed to remove key from library store for key : " & spKeyValue, vbCritical
            End If
        Else
            MsgBox "Failed to remove record data @ " & lpDataIndex & " from library store for key : " & spKeyValue, vbCritical
        End If
    End If

DeleteMedia = bOk

   On Error GoTo 0
   Exit Function

DeleteMedia_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure DeleteMedia of Module MediaLibarySample"
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetMediaRecordDataValues
' DateTime  : 29/04/2005 11:08
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function GetMediaRecordDataValues(ByVal lpLibraryRecord As Long, spaMediaDataValues() As String) As Boolean
 Dim iKeyField As Integer
 
   On Error GoTo GetMediaRecordDataValues_Error

    GetMediaRecordDataValues = False
    Erase spaMediaDataValues
    
    ' library records bounds check
    If (lpLibraryRecord < 0) Or (lpLibraryRecord > lgLibraryRecord) Then
        Exit Function
    End If

    ReDim spaMediaDataValues(igKeyFields)

    ' load the passed array with values from the Library record
    For iKeyField = 0 To igKeyFields
          spaMediaDataValues(iKeyField) = vgaLibraryRecords(lpLibraryRecord, iKeyField)
    Next
  
GetMediaRecordDataValues = True

   On Error GoTo 0
   Exit Function

GetMediaRecordDataValues_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetMediaRecordDataValues of Module MediaLibarySample"
End Function


'---------------------------------------------------------------------------------------
' Procedure : UpdateMediaToLibrary
' DateTime  : 05/05/2005 09:16
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function UpdateMediaToLibrary(ByVal spKeyValue As String, spaMediaDataValues() As String, ByRef lblpStatus As Label) As Boolean
 Dim iKeyField As Integer
 Dim iKeyFields As Integer
 Dim iMediaFields As Integer
 Dim bOk As Boolean
 Dim lpFoundNode As Long
 Dim lLibraryRecord As Long
     
     UpdateMediaToLibrary = False
      
   On Error GoTo UpdateMediaToLibrary_Error
 
    lblpStatus.Caption = ""
    lblpStatus.Refresh
    
    bOk = clsgBSTStore.SearchBSTN(spKeyValue, lpFoundNode)
    If bOk = False Then
        MsgBox "Error Update Record to Library failed key not found for value: " & spKeyValue, vbCritical
        'Err.Raise vbObjectError + 394035, "", ""
        Exit Function
    End If
    
    lLibraryRecord = clsgBSTStore.GetDataNodeIndex(lpFoundNode)
    
    iMediaFields = UBound(spaMediaDataValues())
    ' load the passed array with values from the Library record
    For iKeyField = 0 To igKeyFields
        If iKeyField <= iMediaFields Then
            vgaLibraryRecords(lLibraryRecord, iKeyField) = spaMediaDataValues(iKeyField)
        Else
            vgaLibraryRecords(lLibraryRecord, iKeyField) = ""
        End If
    Next
  
    ' final status
    lblpStatus.Caption = "Added (" & lgLibraryRecord & ") Media Record Key: " & spKeyValue & " on : " & Format(Now, "ddd-mmm-yyyy @ hh:mm")
    lblpStatus.Refresh

    UpdateMediaToLibrary = True

   On Error GoTo 0
   Exit Function

UpdateMediaToLibrary_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure UpdateMediaToLibrary of Module MediaLibarySample"
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : AddMediaToLibrary
' DateTime  : 29/04/2005 18:56
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function AddMediaToLibrary(ByVal spKeyValue As String, spaMediaDataValues() As String, ByRef lblpStatus As Label) As Boolean
 Dim iKeyField As Integer
 Dim iKeyFields As Integer
 Dim iMediaFields As Integer
 Dim bOk As Boolean
 Dim lpFoundNode As Long
 
    AddMediaToLibrary = False
    
   On Error GoTo AddMediaToLibrary_Error
    
    lblpStatus.Caption = ""
    lblpStatus.Refresh
    
    bOk = clsgBSTStore.SearchBSTN(spKeyValue, lpFoundNode)
    If bOk = True Then
        MsgBox "Error Add Record to Library failed duplicate key/index for value: " & spKeyValue, vbCritical
        'Err.Raise vbObjectError + 394035, "", ""
        Exit Function
    End If
    
    bOk = AddMediaDataRecord()
    
    If bOk = False Then
        MsgBox "Error Add Record to Library failed increasing record media store for key: " & spKeyValue, vbCritical
        Err.Raise vbObjectError + 394034, "", ""
        Exit Function
    End If


    bOk = clsgBSTStore.InsertBSTN(spKeyValue, lpFoundNode)

    If bOk = False Then
        MsgBox "Error Add Record to Library failed setting key/index for value: " & spKeyValue, vbCritical
        Err.Raise vbObjectError + 394035, "", ""
        Exit Function
    End If

    bOk = clsgBSTStore.SetDataNodeIndex(lpFoundNode, lgLibraryRecord)

    If bOk = False Then
        MsgBox "Error Add Record to Library failed setting key/index for value: " & spKeyValue, vbCritical
        Err.Raise vbObjectError + 394036, "", ""
        Exit Function
    End If
    iMediaFields = UBound(spaMediaDataValues())
    ' load the passed array with values from the Library record
    For iKeyField = 0 To igKeyFields
        If iKeyField <= iMediaFields Then
            vgaLibraryRecords(lgLibraryRecord, iKeyField) = spaMediaDataValues(iKeyField)
        Else
            vgaLibraryRecords(lgLibraryRecord, iKeyField) = ""
        End If
    Next
  
    ' final status
    lblpStatus.Caption = "Added (" & lgLibraryRecord & ") Media Record Key: " & spKeyValue & " on : " & Format(Now, "ddd-mmm-yyyy @ hh:mm")
    lblpStatus.Refresh

    AddMediaToLibrary = True

   On Error GoTo 0
   Exit Function

AddMediaToLibrary_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure AddMediaToLibrary of Module MediaLibarySample"
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : AddMediaDataRecord
' DateTime  : 02/05/2005 11:06
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function AddMediaDataRecord() As Boolean
 Dim iKeyField As Integer
 Dim lRow As Long
       
    AddMediaDataRecord = False
      
   On Error GoTo AddMediaDataRecord_Error

    ' rebuild the record store from the shadow, having adding a brefore the copy
    ReDim vgaLibraryRecordsShadow(lgLibraryRecord + 1, igKeyFields)
    For lRow = 0 To lgLibraryRecord
        ' load the original record store array with values from the Library shadow store
        For iKeyField = 0 To igKeyFields
              vgaLibraryRecordsShadow(lRow, iKeyField) = vgaLibraryRecords(lRow, iKeyField)
        Next
    Next
    
    ' increase record store by one record
    lgLibraryRecord = lgLibraryRecord + 1
     
    ' rebuild the record store from the shadow, having added an empty record during the copy
    ReDim vgaLibraryRecords(lgLibraryRecord, igKeyFields)
    For lRow = 0 To lgLibraryRecord
        ' load the original record store array with values from the Library shadow store
        For iKeyField = 0 To igKeyFields
              vgaLibraryRecords(lRow, iKeyField) = vgaLibraryRecordsShadow(lRow, iKeyField)
        Next
    Next
        
    Erase vgaLibraryRecordsShadow()

    AddMediaDataRecord = True

   On Error GoTo 0
   Exit Function

AddMediaDataRecord_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure AddMediaDataRecord of Module MediaLibarySample"
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : RemoveMediaDataRecord
' DateTime  : 03/05/2005 20:25
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'don't compress the library or all data indexes are voided
Public Function RemoveMediaDataRecord(ByVal lpLibraryRecord As Long) As Boolean

 Dim iKeyField As Integer
 Dim lRow As Long
 Dim iOffset As Integer
 Dim iCol As Integer
 Dim lDataPos As Long
 Dim iAdjust As Integer
 Dim lIndirectRow As Long
 
RemoveMediaDataRecord = False

   On Error GoTo RemoveMediaDataRecord_Error

    RemoveMediaDataRecord = False
         
    ' library records bounds check
    If (lpLibraryRecord < 0) Or (lpLibraryRecord > lgLibraryRecord) Then
        Exit Function
    End If
    
    For iKeyField = 0 To igKeyFields
          vgaLibraryRecords(lpLibraryRecord, iKeyField) = ""
    Next

RemoveMediaDataRecord = True

   On Error GoTo 0
   Exit Function

RemoveMediaDataRecord_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure RemoveMediaDataRecord of Module MediaLibarySample"

End Function


' eom
