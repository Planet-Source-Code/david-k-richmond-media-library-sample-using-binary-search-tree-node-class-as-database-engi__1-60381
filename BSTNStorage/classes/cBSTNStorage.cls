VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBSTNStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' BUG FIX   :
' Version   : 1.0 to 1.1
' Details   :  Add optional lpdataindex to allow tree rebalance to keep the
'              new root key dataindex else rebalanced key record data index is lost
'              causing orphaned or mixed up data record, if data index
'              being used to referenc external data records.
'
' Other changes:
'               tidy up nomenclature for parameters passed
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'
' NEWS:  April 2005
'        There is a beta media library sample in development using this class as it's database engine.
'        Other p-s-c visitors offered their help in testing the beta sample, checking the adds, removes,
'        ordered lists etc. and it's overall general usage.
'        The sample maintains details of media such as CDs, DVDs, Audio Tapes, VHS Videos etc.
'        Once testing and functionality has been completed, the sample will be published on p-s-c.
'        Anticipated to be published by middle of May 2005 (or sooner!)
'
'---------------------------------------------------------------------------------------
' Project   : cBSTNStorage
' DateTime  : 25/04/2005 20:37
' Author    : D K Richmond
' Status    : BETA RELEASE 1.1
' Version   : Visual Basic 6 (SP6)
' User Type : Beginner/Intermediate/Advanced
' Applicable: Database/Data Sorting/Key Retrieval/Binary Trees/Node Traversals
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
'  I have used a test harness (not included) to prove these functions work as I expect them to work.
'  I can not release the test harness, it's not pretty, there are lots of non-modal controls and really
'  it would take me weeks to make user-friendly and provide documentation on.  You will not
'  benefit from having it as it was only used to test randomly as many the possible combinations of trees.
'  As I have spent many many hours (that's more than 2 weeks several hours a day), to test and
'  refine the routines to give the best performance and reliablity, there is little testing left to do.
'  I have deleted many lines of code either because they were non-sensical, confusing or poor.  In some
'  cases where I wanted to experiment with different ideas, where I felt the normal established nodal functions
'  were too complicated or too unreliable or too complex to test for 100% data integrity.
'  The most complicated routines in the whole class were the RemoveNode and the
'  GetPrevNodeInSortedOrder member functions.  I spent many hours refining and checking
'  these along with others as they are interelated and relate to the correct removal of a node
'  with two sub-branches.
'
' OVERVIEW:
' -----------
'   Things to be aware of:
'      All entries into the Binary Tree structure are via the root node unless the member
'      function relies on a previous action and the next action is incremental or
'      decrimental e.g. GetPrevNodeInSortedOrder
'      To start a tree you will need to call InitRootNode passing the string key
'      you want to assign to the root node.  Now ideally this root node needs to be the
'      nearest to the middle of the sorted data as possible, but don't worry too much about
'      this as it won't affect anything expect a very small performance hit on searches, inserts and deletions.
'      Having entered a root node the next thing is to do is add the other nodes (any order), using InsertBSTN.
'      The most balanced tree will be from completeley random data inserted sequentially.
'      If you insert sorted data you will end up with a tree with a very long branch or very long branches.
'      This means that the searching will not be the fastest it could be.  Don't worry
'      if you have sorted data, just set the root node, read the data and then before using
'      it try a RebalanceTree passing the middle value of the sorted data.  The searches
'      will then be working on a reasonably balanced tree, i.e faster.
'   Where is the tree store?
'      There is a linear array, which is allocated up by one index for each node added.
'      As each item in the array is of type TBSTNStore is has the node values to allow Binary Tree access.
'   Why can't I find some things referenced in the project:
'       Generally anything included in this release should be of publication standard, e.g. be typed and error checked
'       correctly, laid out properly and essentially reliable and of course bug free!  This project is changing weekly
'       and a release will include only those things worthy of publication.  If after completing a major change,
'       I need to publish the code and an non-publishable item is critical to the exection/compilation I am forced to include it
'       as is.  Further releases will either include the item to the correct standard or it's dependencies removed and
'       itself removed from the project.  These are going to be small items and I hope those of you that may have already built
'       applications based on earlier versions of this code will be impacted heavily by any of these necessary changes.
'       In the next few months the core functionality should stabilise and changes we will see will be limited to additional
'       member functions.  Things like tree rotation, maybe an extra new style tree re-balance etc.  Maybe even some more
'       everyday practical examples to demonstrate how these functions are used.
'
'
' DEBUGGING:
' ---------
' In Class Initialise :  gbDebug = False   '  change this to true to see lots of debug message boxes
'
' SURPRISES:
'  I have left some msgbox's and some redundant error trapping in some routines, this will be taken out real soon.
'  The are some extra definitions that I have left for the future.

' FUTURE:
' -------
' There is a lot of scope to improve this class.
'
' o     Ensuring the tree is always the same level deep
' o     Auto balancing the tree
' o     Dealing with garbage collection
' o     Checking garbage collection thresholds
' o     Tree rotation techniques
' o     Using the Binary Trees for real life applications/examples
' o     Providing some simple test harnesses to test and
'       demonstrate the member functions of the class and the usage
' o     performance counters/traversal statistics

' I will be updating this class, adding extra member functions which may address the above future items or other
' items more pressing or necessary for using this class in other projects.  I have some projects that will make
' use of this class immediately.  If you have ideas, want help in understanding things I have ommitted from the
' overview, usage etc, let me know on the above email. Remember the Binary Tree principals are on the internet.
' Try google.com and use the phrase "Binary Trees" with quotes. Please try to use the example below and have a
' try using the functions yourself before contacting me, the rewards will be better.  Also if for example
' 50+ people contact me I am going to struggle to get a reply off to everyone and have time to improve the
' class. Otherwise look for the next update and I hope to include some interesting and hopefully practical
' examples of using the class to the MAX!!
'
'
' Key Compares.
' -----------
'  Use CompareKeys to check if keys are less, equal or more than each other.  This is because
'  there is a lexical member value.  This makes each key lexical compared, so it makes more
'  sense to human readability.  If you change the lexical to non-lexical you will get true
'  computer sorted keys and hence if you use this you will inherit the same comparisons.
'  If you don't and then later change it and you don't change to code outside the class
'  you will get data inconsistent results.
'
' DONT:  Try this with non-unique keys.  For example A,A,B,C,D,D,E,E
'--------------------------------------------------------------------
' If you do try adding the same key more than once, YOU WILL GENERATE an ERROR.
' This is deliberate.  About 10+ years ago, I wrote a procedure (in C) to allow duplicate
' keys in a Binary tree.  It took me ages to perfect, it was always needing updates/tweaks
' and many coding changes.  I know these days, I could add this quite easily but it's a
' lot more testing! So it has not been included in this release. When I have tested
' using duplicate keys in the same tree and it's reliable and allows correct nodal deletes,
' I will then remove the error trap, make it conditional based on global value and thus allow
' duplicate keys to be stored in the tree.  Of course those brave people reading this
' are welcome to take the class and try with duplicate keys, you don't have to wait for  =[8-)
'---------------------------------------------------------------------------------------
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
'
' NOTES for beginners:
' ---------------------
' The class members are protected from external alteration by being made private.  This protects the integrity
' of the data in the binary tree.   If you end up making the private's into public's you need to re-think why you need
' to do so.   There are public property's to allow Read/Write access where applicable to external programs.
' Please try to keep the private members private and if you need to get the value in a private member
' use a public get property.  Try not to go wild here!
' Private Const are used to protect constant values for being changed during runti
'
' ##########################################################################################################
'
' EXAMPLE USAGE:
'
' Create a new project (standard)
'
' add the following controls:
' o     a listbox named List1
' o     a command button named cmdLoadFile caption "Load File"
' o     a command button named cmdShowSorted caption "Show Sorted List"
'
' then add this code to the form module code
' remember to add both the classes using project menu, add class.
' class:   cBSTNStorage
' and
' class:   cList
'
' now paste this section of example code to the form module code window then
' remove the comment prefix using the ide toolbar function
'
' Run the app, click on the cmdShowSorted button and the letters in the Form_Load
' will appear sorted.
' Now create the file as show in the cmdLoadFile comment for DOS
' click on the cmdLoadFile and see a list of alphabetically sorted filenames
' in the listbox.
'
'  As an example for you to see how easy is it with a binary tree class to find items,
'  write a routine where the item clicked in the list is extracted from the list
'  into a string named sKey. Then use the SearchBSTN passing sKey to located the item in the tree.
'  Use the public property's to retrieve the values of the node Left, Right, Parent and DataIndex.
'
'  Next, I expect having seen some potential, you're thinking, yep, okay, not bad, but what's next?
'  Well, you can use SetDataNodeIndex to assign an index (data type long) to the node.
'  So if you have an array of values elsewhere, you can then assign that index to the node's DataIndex value.
'  Whenever the tree is re-organised the DataIndex value moves with the node.
'  So you can store a whole record, keep the index value and assign it to a key node.
'  You can then perform fast searches and use the GetDataNodeIndex to retrieve the offset to the record.
'
' In SUMMARY: Does that sound similar to a key search on a database record or not?

''  #### start ################### EXAMPLE CODE TO PASTE INTO THE MAIN FORM ##########################
'Dim clsBSTStore As cBSTNStorage
'
'
' Sub Form_Load()
'
'    Set clsBSTStore = New cBSTNStorage
'
'    clsBSTStore.InitRootNode "K"
'
'    clsBSTStore.InsertBSTN "H"
'    clsBSTStore.InsertBSTN "V"
'    clsBSTStore.InsertBSTN "L"
'    clsBSTStore.InsertBSTN "Z"
'    clsBSTStore.InsertBSTN "M"
'    clsBSTStore.InsertBSTN "O"
'
'End Sub
'
'
'Private Sub Form_Unload(Cancel As Integer)
'    Set clsBSTStore = Nothing
'End Sub
'
'
' Sub cmdLoadFile_Click()
'  'use DOS command line:   C:\> dir /b >c:\temp\dir.txt
'    clsBSTStore.InitRootNode "."
'     Open "C:\temp\dir.txt" For Input As #12
'    While Not EOF(12)
'        Line Input #12, sLine
'        clsBSTStore.InsertBSTN sLine
'    Wend
'    Close #12
' End Sub
'
'
' Sub cmdShowSorted_Click()
'
' Dim lNextNode As Long
' Dim sKey As String
'
'    List1.Clear
'
'    clsBSTStore.ResetToRootNode
'
'    Do
'        If clsBSTStore.GetNextNodeInSortedOrder(lNextNode) = True Then
'            Debug.Print String(40, "-")
'            Debug.Print "SO_BTNS[" & lNextNode & "] P=" & clsBSTStore.GetParentNodeOfIndex(lNextNode)
'            Debug.Print "SO_BTNS[" & lNextNode & "] L=" & clsBSTStore.GetLeftNodeOfIndex(lNextNode)
'            Debug.Print "SO_BTNS[" & lNextNode & "] R=" & clsBSTStore.GetRightNodeOfIndex(lNextNode)
'            Debug.Print "SO_BTNS[" & lNextNode & "] K=" & clsBSTStore.GetKeyAtNodeIndex(lNextNode)
'            Debug.Print String(40, "=")
'        Else
'            Exit Do
'        End If
'
'        DoEvents
'
'        If (clsBSTStore.TopOfTreeReached = False) Then
'            sKey = clsBSTStore.GetKeyAtNodeIndex(lNextNode)
'            List1.AddItem sKey
'        End If
'    Loop
'
'End Sub
'
'
''  #### end ################### EXAMPLE CODE TO PASTE INTO THE MAIN FORM ##########################

  
Option Explicit
Dim gbDebug As Boolean

'---------------------------------------------------------------------------------------
' Module    : cBSTNStorage
' DateTime  : 17/04/2005 10:01
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------

Private BSTNStore() As TBSTNStore     ' this is the tree node storage data

Public Enum NodeTypes
    m_ePNODE_TYPE = 1       ' Parent node type
    m_eLNODE_TYPE = 2       ' Left node type
    m_eRNODE_TYPE = 3       ' Right node type
End Enum

' IMPORTANT NOTE:
' if you add any extra items to this type you will need to add them to any routines
' that maintain them during tree re-organisations
Private Type TBSTNStore
    m_lParent As Long       ' Parent Node Link
    m_lLeft As Long         ' Left Node Link
    m_lRight As Long        ' Right Node Link
    m_sNodeKey As String    ' Node Key Value
    m_lNodeDataIndex As Long   ' provides index into external data and this is maintained during tree reorganisations
End Type

' A pretty unique garbage identity key.  If your key value is this then the node has been deleted
Private m_sGarbageIdentityKey  ' = Chr(128) & chr(127) x 5
            
' as it's a percentage of those nodes also allocated on top of store this needs to be reasonably low
Private Const m_cgiGarbageCollectThresholdPercentage = 30

' lexical text matching if set true . ie . a and A are equal value for compares
Private Const m_cbLexical = True
 
Public Enum TraverseOpCodes
    m_eiUNKNOWN = -1
    m_eiGOFARLEFT = 1
    m_eiGOright = 2
    m_eiGOUPVISIT = 3
    m_egiGOUP = 4
    m_eiGOFARRIGHT = 5
    m_eiGOLEFT = 6
    m_eiTOPOFTREE = 99
End Enum

' consts TraverseOpTexts
Private Const m_csGOFARLEFT = "GOFARLEFT"
Private Const m_csGORIGHT = "GORIGHT"
Private Const m_csGOUPVISIT = "GOUPVISIT"
Private Const m_csiGOUP = "GOUP"
Private Const m_csTOPOFTREE = "TOPOFTREE"
' end consts

' tree span counters, watchers
Private m_lNodeHitsCount As Long
Private m_lOffsetLeft As Long
Private m_lOffsetRight As Long
Private m_lNodeLevel As Long
Private m_lCurrentNode As Long
Private m_iOperation As Integer
Private m_lRemovedNodesCount As Long

Private Const m_cNoLink = -1

'---------------------------------------------------------------------------------------
' Procedure : GetDataNodeIndex
' DateTime  : 25/04/2005 19:32
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function GetDataNodeIndex(ByVal lpBSTNNodeIndex As Long) As Long

   On Error GoTo GetDataNodeIndex_Error

    If (lpBSTNNodeIndex >= 0) And (lpBSTNNodeIndex <= GetNodesMaxIndex()) Then
        GetDataNodeIndex = BSTNStore(lpBSTNNodeIndex).m_lNodeDataIndex
    Else
        GetDataNodeIndex = -1
        Err.Raise vbObjectError + 39049, "", ""
        Exit Function
    End If

   On Error GoTo 0
   Exit Function

GetDataNodeIndex_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetDataNodeIndex of Class Module cBSTNStorage"

End Function

'---------------------------------------------------------------------------------------
' Procedure : SetDataNodeIndex
' DateTime  : 25/04/2005 19:32
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function SetDataNodeIndex(ByVal lpBSTNNodeIndex As Long, ByVal lpObjectNodeIndex As Long) As Boolean

SetDataNodeIndex = False
    
    On Error GoTo SetDataNodeIndex_Error

    If (lpBSTNNodeIndex >= 0) And (lpBSTNNodeIndex <= GetNodesMaxIndex()) Then
        BSTNStore(lpBSTNNodeIndex).m_lNodeDataIndex = lpObjectNodeIndex
    Else
        Err.Raise vbObjectError + 39049, "", ""
        Exit Function
    End If

SetDataNodeIndex = True

   On Error GoTo 0
   Exit Function

SetDataNodeIndex_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SetDataNodeIndex of Class Module cBSTNStorage"
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetGarbageCollectThresholdPercentage
' DateTime  : 25/04/2005 19:32
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get GetGarbageCollectThresholdPercentage() As Integer
    GetGarbageCollectThresholdPercentage = m_cgiGarbageCollectThresholdPercentage
End Property

'---------------------------------------------------------------------------------------
' Procedure : GetTraverseOperationCode
' DateTime  : 25/04/2005 19:32
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get GetTraverseOperationCode() As Integer
    GetTraverseOperationCode = m_iOperation
End Property

'---------------------------------------------------------------------------------------
' Procedure : GetTraverseOperationText
' DateTime  : 25/04/2005 19:32
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get GetTraverseOperationText() As String
    GetTraverseOperationText = GetTraverseTextForOpcode(m_iOperation)
End Property

'---------------------------------------------------------------------------------------
' Procedure : GetTraverseCurrentNodeIndex
' DateTime  : 25/04/2005 19:32
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get GetTraverseCurrentNodeIndex() As Long
    GetTraverseCurrentNodeIndex = m_lCurrentNode
End Property

'---------------------------------------------------------------------------------------
' Procedure : GetTraverseCurrentNodeKey
' DateTime  : 25/04/2005 19:32
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get GetTraverseCurrentNodeKey() As Long
    GetTraverseCurrentNodeKey = BSTNStore(m_lCurrentNode).m_sNodeKey
End Property

'---------------------------------------------------------------------------------------
' Procedure : GetTraverseNodeLevel
' DateTime  : 25/04/2005 19:32
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get GetTraverseNodeLevel() As Long
    GetTraverseNodeLevel = m_lNodeLevel
End Property

'---------------------------------------------------------------------------------------
' Procedure : GetTraverseOffsetLeft
' DateTime  : 25/04/2005 19:32
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get GetTraverseOffsetLeft() As Long
    GetTraverseOffsetLeft = m_lOffsetLeft
End Property

'---------------------------------------------------------------------------------------
' Procedure : GetTraverseOffsetRight
' DateTime  : 25/04/2005 19:32
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get GetTraverseOffsetRight() As Long
    GetTraverseOffsetRight = m_lOffsetRight
End Property

'---------------------------------------------------------------------------------------
' Procedure : TopOfTreeReached
' DateTime  : 25/04/2005 19:32
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function TopOfTreeReached() As Boolean
    TopOfTreeReached = (m_iOperation = m_eiTOPOFTREE)
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetCurrentNodeIndex
' DateTime  : 25/04/2005 19:32
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get GetCurrentNodeIndex() As Long
    GetCurrentNodeIndex = m_lCurrentNode
End Property


'---------------------------------------------------------------------------------------
' Procedure : GetKeyAtNodeIndex
' DateTime  : 25/04/2005 19:32
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function GetKeyAtNodeIndex(ByVal lpNodeIndex As Long) As String
    
   On Error GoTo GetKeyAtNodeIndex_Error

    If (lpNodeIndex >= 0) And (lpNodeIndex <= GetNodesMaxIndex()) Then
        GetKeyAtNodeIndex = BSTNStore(lpNodeIndex).m_sNodeKey
    Else
        Err.Raise vbObjectError + 39009, "GetKeyAtNodeIndex", "GetKeyAtNodeIndex"
        Exit Function
    End If

   On Error GoTo 0
   Exit Function

GetKeyAtNodeIndex_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetKeyAtNodeIndex of Class Module cBSTNStorage"
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetLeftNodeOfIndex
' DateTime  : 25/04/2005 19:32
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function GetLeftNodeOfIndex(ByVal lpNodeIndex As Long) As Long

   On Error GoTo GetLeftNodeOfIndex_Error

    If (lpNodeIndex >= 0) And (lpNodeIndex <= GetNodesMaxIndex()) Then
        GetLeftNodeOfIndex = BSTNStore(lpNodeIndex).m_lLeft
    Else
        Err.Raise vbObjectError + 39041, "GetLeftNodeOfIndex", "GetLeftNodeOfIndex"
        Exit Function
    End If

   On Error GoTo 0
   Exit Function

GetLeftNodeOfIndex_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetLeftNodeOfIndex of Class Module cBSTNStorage"

End Function

'---------------------------------------------------------------------------------------
' Procedure : GetRightNodeOfIndex
' DateTime  : 25/04/2005 19:32
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function GetRightNodeOfIndex(ByVal lpNodeIndex As Long) As Long
    
   On Error GoTo GetRightNodeOfIndex_Error

    If (lpNodeIndex >= 0) And (lpNodeIndex <= GetNodesMaxIndex()) Then
        GetRightNodeOfIndex = BSTNStore(lpNodeIndex).m_lRight
    Else
        Err.Raise vbObjectError + 39049, "GetRightNodeOfIndex", "GetRightNodeOfIndex"
        Exit Function
    End If

   On Error GoTo 0
   Exit Function

GetRightNodeOfIndex_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetRightNodeOfIndex of Class Module cBSTNStorage"
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetParentNodeOfIndex
' DateTime  : 25/04/2005 19:32
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function GetParentNodeOfIndex(ByVal lpNodeIndex As Long) As Long

   On Error GoTo GetParentNodeOfIndex_Error

    If (lpNodeIndex >= 0) And (lpNodeIndex <= GetNodesMaxIndex()) Then
        GetParentNodeOfIndex = BSTNStore(lpNodeIndex).m_lParent
    Else
        Err.Raise vbObjectError + 39043, "GetParentNodeOfIndex", "GetParentNodeOfIndex"
        Exit Function
    End If

   On Error GoTo 0
   Exit Function

GetParentNodeOfIndex_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetParentNodeOfIndex of Class Module cBSTNStorage"
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetNodesMaxIndex
' DateTime  : 25/04/2005 19:33
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function GetNodesMaxIndex() As Long
    GetNodesMaxIndex = UBound(BSTNStore())
End Function


'---------------------------------------------------------------------------------------
' Procedure : Clear
' DateTime  : 18/04/2005 08:17
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function Clear()

   On Error GoTo Clear_Error

    ReDim BSTNStore(0)
        
    BSTNStore(0).m_lParent = m_cNoLink
    BSTNStore(0).m_lLeft = m_cNoLink
    BSTNStore(0).m_lRight = m_cNoLink
    BSTNStore(0).m_sNodeKey = m_sGarbageIdentityKey

    m_lRemovedNodesCount = 0

   On Error GoTo 0
   Exit Function

Clear_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Clear of Class Module cBSTNStorage"

End Function


'---------------------------------------------------------------------------------------
' Procedure : InitRootNode
' DateTime  : 18/04/2005 08:17
' Author    : D K Richmond
' Purpose   :
' BUG FIX   : add optional lpdataindex to allow rebalance to keep the new root key dataindex
'---------------------------------------------------------------------------------------
'
Public Function InitRootNode(ByVal spNodeKey As String, Optional ByVal lpDataIndex As Variant)

   On Error GoTo InitRootNode_Error

    ReDim BSTNStore(0)
        
    BSTNStore(0).m_lParent = m_cNoLink
    BSTNStore(0).m_lLeft = m_cNoLink
    BSTNStore(0).m_lRight = m_cNoLink
    BSTNStore(0).m_sNodeKey = spNodeKey
' BUG FIX   : add optional lpdataindex to allow rebalance to keep the new root key dataindex
    If IsMissing(lpDataIndex) Then
        BSTNStore(0).m_lNodeDataIndex = -1
    Else
        BSTNStore(0).m_lNodeDataIndex = lpDataIndex
    End If
    m_lRemovedNodesCount = 0

   On Error GoTo 0
   Exit Function

InitRootNode_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure InitRootNode of Class Module cBSTNStorage"

End Function

'---------------------------------------------------------------------------------------
' Procedure : InitNodeTree
' DateTime  : 25/04/2005 19:33
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Public Function InitNodeTree(ByVal spNodeKey As String)
'
'    ReDim BSTNStore(0)
'
'    BSTNStore(0).m_lParent = m_cNoLink
'    BSTNStore(0).m_lLeft = m_cNoLink
'    BSTNStore(0).m_lRight = m_cNoLink
'    BSTNStore(0).m_sNodeKey = ""
'BSTNStore(0).m_lNodeDataIndex = -1
'
'End Function


'---------------------------------------------------------------------------------------
' Procedure : GetTraverseTextForOpcode
' DateTime  : 17/04/2005 10:12
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function GetTraverseTextForOpcode(ByVal ipTraverseOpCode As Integer) As String
 Dim sOPeration As String
 
    Select Case ipTraverseOpCode
        Case TraverseOpCodes.m_eiGOFARLEFT:
            sOPeration = m_csGOFARLEFT
        Case TraverseOpCodes.m_eiGOright:
            sOPeration = m_csGORIGHT
        Case TraverseOpCodes.m_eiGOUPVISIT:
            sOPeration = m_csGOUPVISIT
        Case TraverseOpCodes.m_egiGOUP:
            sOPeration = m_csiGOUP
        Case TraverseOpCodes.m_eiTOPOFTREE:
            sOPeration = m_csTOPOFTREE
        Case Else
            sOPeration = "?? unknown ??"
     End Select

GetTraverseTextForOpcode = sOPeration
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetNextNodeInSortedOrder
' DateTime  : 17/04/2005 10:01
' Author    : D K Richmond
' Purpose   :
' Notes     : pass thru detects when the node traversal ascends thru a specific node and sets flag
'---------------------------------------------------------------------------------------
'
Public Function GetNextNodeInSortedOrder(ByRef lpNextNode As Long, Optional ByRef lpoCurrentNode As Long = -1, Optional ByRef ipoOperation As Integer = m_eiUNKNOWN, Optional bpoPassThru As Boolean, Optional lpoPassThruNode As Long = -1) As Boolean
 Dim lLeft As Long
 Dim lRight As Long
 Dim lParent As Long
 Dim bMoved As Boolean
 
 GetNextNodeInSortedOrder = False
 
   On Error GoTo GetNextNodeInSortedOrder_Error

    ' @node
    'IF LEFT exist
    '    GO FAR LEFT
    '    next op is go right
    'End If
    '
    'if NOT LEFT exist
    '    if RIGHT exist
    '        GO RIGHT ONE
    '        GO FAR LEFT
    '    End If
    '    ' no left or right
    '    if NOT RIGHT exist
    '        go up
    '        if ascend from right i.e if parent's right is same as current
    '            go up
    '        End If
    '        if ascend from left i.e if parent's left is same as current
    '            at Node
    '        End If
    '    End If
    'End If

    If lpoCurrentNode <> -1 Then
        m_lCurrentNode = lpoCurrentNode
    End If
    If ipoOperation <> m_eiUNKNOWN Then
        m_iOperation = ipoOperation
    End If
    
    bpoPassThru = False
    lpNextNode = m_lCurrentNode
    bMoved = False
    If m_iOperation = TraverseOpCodes.m_eiGOFARLEFT Then
        Do
            lLeft = BSTNStore(lpNextNode).m_lLeft
            If lLeft = m_cNoLink Then
                Exit Do
            Else
                lpNextNode = lLeft
            End If
            If lpoPassThruNode = lpNextNode Then
                bpoPassThru = True
            End If
            m_lNodeLevel = m_lNodeLevel + 1
            m_lOffsetLeft = m_lOffsetLeft + 1
            m_lOffsetRight = m_lOffsetRight - 1
            bMoved = True
        Loop
        If bMoved = True Then
            lRight = BSTNStore(lpNextNode).m_lRight
            If lRight <> m_cNoLink Then
                m_iOperation = TraverseOpCodes.m_eiGOright
            Else
                m_iOperation = TraverseOpCodes.m_egiGOUP
            End If
            m_lNodeLevel = m_lNodeLevel + 1
            m_lCurrentNode = lpNextNode
            GetNextNodeInSortedOrder = bMoved
            Exit Function
        Else
            m_iOperation = TraverseOpCodes.m_eiGOright
            m_lCurrentNode = lpNextNode
            GetNextNodeInSortedOrder = True
            Exit Function
        End If
    End If
    
    'DEBUG
    Debug.Print BSTNStore(m_lCurrentNode).m_sNodeKey
    
    If m_iOperation = TraverseOpCodes.m_eiGOright Then
        lRight = BSTNStore(m_lCurrentNode).m_lRight
        If lRight <> m_cNoLink Then
            lpNextNode = lRight
            If lpoPassThruNode = lpNextNode Then
                bpoPassThru = True
            End If
            bMoved = True
            m_lNodeLevel = m_lNodeLevel + 1
            m_lOffsetRight = m_lOffsetRight + 1
            m_lOffsetLeft = m_lOffsetLeft - 1
            Do  'GOFARLEFT
                lLeft = BSTNStore(lpNextNode).m_lLeft
                If lLeft = m_cNoLink Then
                    Exit Do
                End If
                lpNextNode = lLeft
                If lpoPassThruNode = lpNextNode Then
                    bpoPassThru = True
                End If
                m_lNodeLevel = m_lNodeLevel + 1
                m_lOffsetLeft = m_lOffsetLeft + 1
                m_lOffsetRight = m_lOffsetRight - 1
                bMoved = True
            Loop
            If bMoved = True Then
                lRight = BSTNStore(lpNextNode).m_lRight
                If lRight <> m_cNoLink Then
                    m_iOperation = TraverseOpCodes.m_eiGOright
                Else
                    m_iOperation = TraverseOpCodes.m_egiGOUP
                End If
                m_lCurrentNode = lpNextNode
                GetNextNodeInSortedOrder = bMoved
                Exit Function
            Else
                 m_iOperation = TraverseOpCodes.m_egiGOUP
            End If
            m_lCurrentNode = lpNextNode
            GetNextNodeInSortedOrder = bMoved
            Exit Function
         Else
            m_iOperation = TraverseOpCodes.m_egiGOUP
        End If
     End If
     
    If m_iOperation = TraverseOpCodes.m_egiGOUP Then
        lParent = BSTNStore(m_lCurrentNode).m_lParent
        If lParent <> m_cNoLink Then
            m_lNodeLevel = m_lNodeLevel - 1
            Do
                If BSTNStore(lParent).m_lParent = m_cNoLink Then
                   ' MsgBox "eek!"
                    'Err.Raise vbobjecterror +940945, "", ""
                    'Exit Function
                    ValidateTree  ' DEBUG
                End If
                If BSTNStore(lParent).m_lRight = m_lCurrentNode Then
                    m_lOffsetRight = m_lOffsetRight - 1
                    m_lOffsetLeft = m_lOffsetLeft + 1
                    ' ascending don't show node
                    m_lCurrentNode = lParent
                    If lpoPassThruNode = m_lCurrentNode Then
                        bpoPassThru = True
                    End If
                    lParent = BSTNStore(m_lCurrentNode).m_lParent
                    If lParent = m_cNoLink Then
                        m_iOperation = TraverseOpCodes.m_eiTOPOFTREE
                        lpNextNode = 0
                        m_lCurrentNode = lpNextNode
                        GetNextNodeInSortedOrder = False
                        Exit Function
                    Else
                        m_lNodeLevel = m_lNodeLevel - 1
                    End If
                Else
                    If BSTNStore(lParent).m_lLeft = m_lCurrentNode Then
                        m_lOffsetLeft = m_lOffsetLeft - 1
                        m_lOffsetRight = m_lOffsetRight + 1
                        ' ascending show node
                        m_lCurrentNode = lParent
                        lpNextNode = lParent
                        If lpoPassThruNode = lpNextNode Then
                            bpoPassThru = True
                        End If
                        bMoved = True
                        m_iOperation = TraverseOpCodes.m_eiGOright
                        m_lCurrentNode = lpNextNode
                        GetNextNodeInSortedOrder = bMoved
                        Exit Function
                    End If
                End If
            Loop
        Else
            m_iOperation = TraverseOpCodes.m_eiTOPOFTREE
        End If
    End If

GetNextNodeInSortedOrder = bMoved

   On Error GoTo 0
   Exit Function

GetNextNodeInSortedOrder_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetNextNodeInSortedOrder of Class Module cBSTNStorage"
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetNextAllocationNode
' DateTime  : 17/04/2005 10:01
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function GetNextAllocationNode() As Long
    
    GetNextAllocationNode = UBound(BSTNStore) + 1
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : SearchBSTN
' DateTime  : 17/04/2005 10:01
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function SearchBSTN(ByVal spKey As String, ByRef lpVisitedNode As Long, Optional ByVal lpCurrentNode As Long = 0) As Boolean
 Dim bNodeFound As Boolean
 Dim lLeft As Long
 Dim lRight As Long
 Dim sKey As String
 
    m_lNodeHitsCount = m_lNodeHitsCount + 1
 
    bNodeFound = False  ' default result is not found
    
    If lpCurrentNode = m_cNoLink Then
        SearchBSTN = False
        Exit Function
    End If
    
    lpVisitedNode = lpCurrentNode
    
    sKey = BSTNStore(lpCurrentNode).m_sNodeKey
    
    If m_cbLexical = True Then
        sKey = UCase(sKey)
        spKey = UCase(spKey)
    End If
    
    If spKey = sKey Then
        ' yes key found
        bNodeFound = True
    End If
    If spKey < sKey Then
        ' check left branch
        lLeft = BSTNStore(lpCurrentNode).m_lLeft
        If lLeft <> m_cNoLink Then
            bNodeFound = SearchBSTN(spKey, lpVisitedNode, lLeft)
        Else
            Exit Function
        End If
    End If
    If spKey > sKey Then
        ' check right branch
        lRight = BSTNStore(lpCurrentNode).m_lRight
        If lRight <> m_cNoLink Then
            bNodeFound = SearchBSTN(spKey, lpVisitedNode, lRight)
        Else
            Exit Function
        End If
    End If

SearchBSTN = bNodeFound
End Function


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' DateTime  : 17/04/2005 10:18
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()

    gbDebug = False   '  change this to true to see lots of debug message boxes
    
    m_sGarbageIdentityKey = String(5, Chr(128) & Chr(127))
    
    InitRootNode m_sGarbageIdentityKey
    
    m_lRemovedNodesCount = 0
    
End Sub


Public Function ResetToRootNode() As Boolean

    ' reset counters, trackers.
    m_lOffsetLeft = 0
    m_lOffsetRight = 0
    m_lNodeLevel = 0
    
    ' reset traversal params
    m_lCurrentNode = 0
    m_iOperation = TraverseOpCodes.m_eiGOFARLEFT

ResetToRootNode = True
End Function


'---------------------------------------------------------------------------------------
' Procedure : InsertBSTN
' DateTime  : 17/04/2005 11:27
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function InsertBSTN(ByVal spKey As String, Optional ByRef plAllocatedNode) As Boolean
 Dim sKey As String
 Dim bNodeFound As Boolean
 Dim lVisitedNode As Long
 'Dim lAllocatedNode As Long
 Dim sVisited As String
 Dim sAllocated As String
 
    InsertBSTN = False
    plAllocatedNode = -1
    m_lNodeHitsCount = 0
 
    If BSTNStore(0).m_sNodeKey = m_sGarbageIdentityKey Then
        MsgBox "Tree is empty going to replace root node with " & spKey
        BSTNStore(0).m_sNodeKey = spKey
        InsertBSTN = True
        plAllocatedNode = 0
        Exit Function
    End If
    
    bNodeFound = SearchBSTN(spKey, lVisitedNode)

    ' prevent duplicates for now
    If bNodeFound = False Then
        plAllocatedNode = GetNextAllocationNode()
        
        ReDim Preserve BSTNStore(plAllocatedNode)
        
        BSTNStore(plAllocatedNode).m_sNodeKey = spKey
        
        sAllocated = BSTNStore(plAllocatedNode).m_sNodeKey
        sVisited = BSTNStore(lVisitedNode).m_sNodeKey
        If m_cbLexical = True Then
            sAllocated = UCase(sAllocated)
            sVisited = UCase(sVisited)
        End If

        ' traversals must end on a leaf of some kind so this is valid
        ' no checking of m_cNoLink required.
        If sAllocated < sVisited Then
            BSTNStore(lVisitedNode).m_lLeft = plAllocatedNode
        Else
            BSTNStore(lVisitedNode).m_lRight = plAllocatedNode
        End If
        
        BSTNStore(plAllocatedNode).m_lParent = lVisitedNode
        BSTNStore(plAllocatedNode).m_lLeft = m_cNoLink
        BSTNStore(plAllocatedNode).m_lRight = m_cNoLink
        BSTNStore(plAllocatedNode).m_lNodeDataIndex = m_cNoLink
        InsertBSTN = True

    End If
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : ReleaseRootNode
' DateTime  : 17/04/2005 13:37
' Author    : D K Richmond
' Purpose   : used to release the root node prior to node removal
'---------------------------------------------------------------------------------------
'
Public Function ReleaseRootNode() As Boolean
 Dim ListX As cList
 Dim lNextNode As Long
 Dim iOperation As Integer
 Dim sKey As String
 Dim sOPeration As String
 Dim lIndex As Long
 Dim lMax As Long
 Dim lLeft As Long
 Dim lRight As Long
 Dim lParent As Long
 Dim sRebalanceKey As String
 Dim sRootNodeKey As String
 Dim iRandom As Integer
 Dim bOk As Boolean
 Dim lDataIndex As Long
 
ReleaseRootNode = False
 
    Set ListX = New cList
    
    ' use list to rebuild keys held in the tree one side of root node key
    ListX.Clear
 
    lParent = GetParentNodeOfIndex(0)
    lLeft = GetLeftNodeOfIndex(0)
    lRight = GetRightNodeOfIndex(0)
    sRootNodeKey = GetKeyAtNodeIndex(0)
    sRebalanceKey = ""  ' if fails nothing happens

    If (lLeft = m_cNoLink) And (lRight <> m_cNoLink) Then
        ' rebalance on right
        sRebalanceKey = GetKeyAtNodeIndex(lRight)
        lDataIndex = GetDataNodeIndex(lRight)
    End If
    If (lLeft <> m_cNoLink) And (lRight = m_cNoLink) Then
        ' rebalance on left
        sRebalanceKey = GetKeyAtNodeIndex(lLeft)
        lDataIndex = GetDataNodeIndex(lLeft)
    End If
    If (lLeft <> m_cNoLink) And (lRight <> m_cNoLink) Then
        iRandom = Int(Rnd(1) * 2)
        If (iRandom \ 1) = 1 Then
            sRebalanceKey = GetKeyAtNodeIndex(lLeft)
            lDataIndex = GetDataNodeIndex(lLeft)
        Else
            sRebalanceKey = GetKeyAtNodeIndex(lRight)
            lDataIndex = GetDataNodeIndex(lRight)
        End If
    End If
    
    bOk = RebalanceTree(sRebalanceKey, False, lDataIndex)
    
    If bOk = False Then
        MsgBox "Problem with Rebalance in ReleaseRoot Node", vbCritical
        Err.Raise vbObjectError + 4054095, "", ""
    End If
    
ReleaseRootNode = bOk
End Function

'---------------------------------------------------------------------------------------
' Procedure : WhatGarbage
' DateTime  : 17/04/2005 20:31
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
' notes: 'there is a time penalty for each time to call this if dynamic used
Public Function WhatGarbage(ByVal bAutoGarbageCollect As Boolean, Optional ByVal bDynamic As Boolean = False) As Double
Dim lMax As Long
Dim lIndex As Long
Dim lCount As Long
Dim dPercentGarbage As Double
Dim bOk As Boolean

    lMax = GetNodesMaxIndex()
    If bDynamic = True Then
        lCount = 0
        For lIndex = 0 To lMax
            If GetKeyAtNodeIndex(lIndex) = m_sGarbageIdentityKey Then
                lCount = lCount + 1
            End If
        Next
    Else
        lCount = m_lRemovedNodesCount
    End If
    
    If (lCount > 0) And (lMax > 0) Then
        dPercentGarbage = lCount / lMax * 100
    Else
        dPercentGarbage = 0
    End If
    
    If dPercentGarbage > m_cgiGarbageCollectThresholdPercentage Then
        If bAutoGarbageCollect Then
            MsgBox "Garbage Collection Activated G%=" & dPercentGarbage & " >  GMAX%=" & m_cgiGarbageCollectThresholdPercentage
            bOk = RebalanceTree()
            If bOk = False Then
                MsgBox "Garbage Collection Failed!"
                Err.Raise vbObjectError + 3490394, "", ""
            End If
        End If
    End If

WhatGarbage = dPercentGarbage
End Function

'---------------------------------------------------------------------------------------
' Procedure : RebalanceRootNode
' DateTime  : 17/04/2005 20:36
' Author    : D K Richmond
' Purpose   : Used to rebalance the tree around a central key value - may promote a need
'               to call Garbage Collection
' Notes     : uses pseudo non-gui list class
'
' BUG FIX   : add optional lpdataindex to allow rebalance to keep the new root key dataindex
' NOTE: if around middle is used the local lpdataindex is recalculated else the data index passed is used
'         for the new root node
'---------------------------------------------------------------------------------------
'
Public Function RebalanceTree(Optional ByVal spKey As String, Optional ByVal bpoAroundMiddle As Boolean = False, Optional lpoDataIndex As Long) As Boolean
 Dim lAllocatedNode As Long
 Dim lNextNode As Long
 Dim sKey As String
 Dim lIndex As Long
 Dim lMax As Long
 Dim sRebalanceKey As String
 Dim lRebalanceDataIndex As Long
 Dim sNodeKey As String
 Dim lNodeCount As Long
 Dim lNodesMiddle As Long
 Dim List4 As cList
 Dim List6 As cList
 Dim bOk As Boolean

RebalanceTree = False

    If spKey = "" Then
        sNodeKey = GetKeyAtNodeIndex(0)
    Else
        sNodeKey = spKey
    End If
    
    If sNodeKey <> m_sGarbageIdentityKey Then
        sRebalanceKey = sNodeKey
        lRebalanceDataIndex = lpoDataIndex
    Else
        MsgBox "Tree is empty!", vbCritical
        Exit Function
    End If
    
    Set List4 = New cList
    Set List6 = New cList
    ' use list to rebuild keys held in the tree one side of root node key
    List4.Clear
    List6.Clear
    
    lNodeCount = 0
    lNodesMiddle = CLng((GetNodesMaxIndex() + 1) / 2)
    ResetToRootNode
    Do
        If GetNextNodeInSortedOrder(lNextNode) = True Then
            Debug.Print String(40, "-")
            Debug.Print "SO_BSTN[" & lNextNode & "] P=" & GetParentNodeOfIndex(lNextNode)
            Debug.Print "SO_BSTN[" & lNextNode & "] L=" & GetLeftNodeOfIndex(lNextNode)
            Debug.Print "SO_BSTN[" & lNextNode & "] R=" & GetRightNodeOfIndex(lNextNode)
            Debug.Print "SO_BSTN[" & lNextNode & "] C**=" & GetKeyAtNodeIndex(lNextNode)
            Debug.Print String(40, "=")
        Else
            Exit Do
        End If
        If (TopOfTreeReached = False) Then
            sKey = GetKeyAtNodeIndex(lNextNode)
            List4.AddItem sKey
            List6.AddItem BSTNStore(lNextNode).m_lNodeDataIndex
            lNodeCount = lNodeCount + 1
            If bpoAroundMiddle = True Then
                If lNodeCount >= lNodesMiddle Then
                    sRebalanceKey = sKey
                    lRebalanceDataIndex = BSTNStore(lNextNode).m_lNodeDataIndex
                    bpoAroundMiddle = False
                End If
            End If
        End If
        If (TopOfTreeReached = True) Then
            Exit Do
        End If
    Loop
    
    Debug.Print String(40, "*")
    
    lMax = List4.ListCount - 1
    For lIndex = 0 To lMax
        Debug.Print "SO_BSTN[" & lIndex & "] C**=" & GetKeyAtNodeIndex(lIndex)
    Next
    Debug.Print String(40, "=")
    Debug.Print "SO_BSTN Root [" & 0 & "] C**=" & sNodeKey
    
    If sRebalanceKey = "" Then
        InitRootNode m_sGarbageIdentityKey
        MsgBox "Tree is empty", vbInformation
        'Err.Raise vbobjecterror + 3490394, "RootNodeKey null error!", ""
        Exit Function
    End If
    
    InitRootNode sRebalanceKey
' BUG FIX   : add optional lpdataindex to allow rebalance to keep the new root key dataindex
    ' preserve the data index - root is always node 0
    ' index is passed or set from middle item
    SetDataNodeIndex 0, lRebalanceDataIndex

    ' needs to rebuild randomly in order to have a reasonable tree shape else all left or right nodes
    While List4.ListCount > 0  ' allow for rebalance key
        lIndex = Int(Rnd(1) * (List4.ListCount - 1))
        sKey = List4.List(lIndex)
        If sKey <> sRebalanceKey Then
            Debug.Print "Replay Key: " & sKey
            bOk = InsertBSTN(sKey, lAllocatedNode)
            If bOk = False Then
                MsgBox "Failed to replay key : " & sKey & " DATASTORE CORRUPTED!"
                Err.Raise vbObjectError + 9393390, "", ""
                Exit Function
            Else
                BSTNStore(lAllocatedNode).m_lNodeDataIndex = CLng(List6.List(lIndex))
                List4.RemoveItem lIndex
                List6.RemoveItem lIndex
            End If
        Else
            List4.RemoveItem lIndex
            List6.RemoveItem lIndex
        End If
    Wend
    
    
Set List4 = Nothing
Set List6 = Nothing
RebalanceTree = True
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetPrevNodeInSortedOrder
' DateTime  : 25/04/2005 20:48
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function GetPrevNodeInSortedOrder(ByVal spKey As String, ByRef lpVisitedNode As Long, ByRef spVisitedKey As String) As Boolean
 Dim lCurrentNode As Long
 Dim lNextNode As Long
 Dim sKey As String
 Dim bOk As Boolean
 Dim lLeft As Long
 Dim sOrgKey As String
 Dim lParent As Long
 Dim lRight As Long
 Dim bNodeFound As Boolean
 Dim bPassThru As Boolean
 Dim lPassThru As Long
 Dim bLeftie As Boolean
 
 
 GetPrevNodeInSortedOrder = False
 
    spVisitedKey = ""
    lpVisitedNode = -1
       
    sOrgKey = spKey

    bOk = SearchBSTN(sOrgKey, lCurrentNode)
    If bOk = False Then
        Exit Function
    End If
    
    lPassThru = lCurrentNode
    bNodeFound = False
    ' first find a parent that is a right node of current
    bLeftie = False
    
    lLeft = GetLeftNodeOfIndex(lCurrentNode)
    If lLeft <> m_cNoLink Then
        MsgBox "leftdec!"
        lCurrentNode = lLeft
        bNodeFound = GetNextNodeInSortedOrder(lNextNode, lCurrentNode, m_eiGOFARLEFT, bPassThru, lPassThru)  ' GOFARLEFT
    Else
        Do
            lParent = GetParentNodeOfIndex(lCurrentNode)
            Debug.Print "SO_BTNS[" & lCurrentNode & "] C**=" & GetKeyAtNodeIndex(lCurrentNode)
                
            If lParent = m_cNoLink Then
                MsgBox "no parent!"
                Exit Do
            End If
            
            lLeft = GetLeftNodeOfIndex(lParent)
            If lLeft = lCurrentNode Then
                ' ascended to a right parent
                MsgBox "ascended to a right parent"
                lCurrentNode = lParent
            End If
            lRight = GetRightNodeOfIndex(lParent)
            If lRight = lCurrentNode Then
                ' ascended to a left parent
                bLeftie = True
                MsgBox "ascended to a right parent"
            End If
            
            Debug.Print "SOP_BTNS[" & lCurrentNode & "] C**=" & GetKeyAtNodeIndex(lParent)
          
            If bLeftie = True Then
                    lCurrentNode = lParent
                    bNodeFound = GetNextNodeInSortedOrder(lNextNode, lCurrentNode, m_eiGOFARLEFT, bPassThru, lPassThru)  ' GOFARLEFT
                    Exit Do
                End If
        Loop
    End If
    
    If bNodeFound = False Then
        MsgBox "start of alpha list"
        Exit Function
    End If
    
    Do
        If bPassThru = True Then
            Exit Do
        End If
        
        If bNodeFound = True Then
            Debug.Print String(40, "-")
            Debug.Print "SO_BTNS[" & lNextNode & "] P=" & GetParentNodeOfIndex(lNextNode)
            Debug.Print "SO_BTNS[" & lNextNode & "] L=" & GetLeftNodeOfIndex(lNextNode)
            Debug.Print "SO_BTNS[" & lNextNode & "] R=" & GetRightNodeOfIndex(lNextNode)
            Debug.Print "SO_BTNS[" & lNextNode & "] C**=" & GetKeyAtNodeIndex(lNextNode)
            Debug.Print String(40, "=")
        Else
            Exit Do
        End If
        
        If (TopOfTreeReached = False) Then
            sKey = GetKeyAtNodeIndex(lNextNode)
        End If
     
        bNodeFound = GetNextNodeInSortedOrder(lNextNode, bpoPassThru:=bPassThru, lpoPassThruNode:=lPassThru)
        
    Loop
  

MsgBox "Previous from : " & sOrgKey & " is " & sKey
 

lpVisitedNode = lCurrentNode
spVisitedKey = GetKeyAtNodeIndex(lpVisitedNode)  'sKey
GetPrevNodeInSortedOrder = True
End Function

'------------------------------------------
' -- Function: GetNodeType
' -- Desc    : determine from parent and current details if node is left,right or parent
' -- Passed  :
' -- Uses    :
' -- Calls   : other functions
' -- Returns : LEFT,RIGHT OR PARENT type node
' -- DateTime: Org-19:54 22/05/00
'------------------------------------------
Function GetNodeType(ByVal lpCurrentNode As Long) As Integer
 Dim lRight As Long
 Dim lParentNode  As Long
 Dim lLeft As Long
 Dim lCurrentNode As Long
 
    lCurrentNode = lpCurrentNode
    
    ' to find out what type of node we are on we need to interogate the parent node
    ' check the value of the right and left node with current node
    GetNodeType = m_ePNODE_TYPE  ' default with no siblings
    lParentNode = GetParentNodeOfIndex(lpCurrentNode)
    If lParentNode <> m_cNoLink Then
        lLeft = GetLeftNodeOfIndex(lParentNode)
        If lCurrentNode = lLeft Then
            GetNodeType = m_eLNODE_TYPE
        End If
        ' only really need to do this if no left node found
        lRight = GetRightNodeOfIndex(lParentNode)
        If lCurrentNode = lRight Then
            GetNodeType = m_eRNODE_TYPE
        End If
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : CompKeys
' DateTime  : 20/04/2005 13:18
' Author    : D K Richmond
' Purpose   : use this for all key comparisons as it allow for lexical or non-lexical
'---------------------------------------------------------------------------------------
'
Public Function CompareKeys(ByVal spKey1 As String, spKey2 As String) As Integer

    ' if key1 < key 2 return -1
    ' if key1 = key 2 return 0
    ' if key1 > key 2 return +1

    If m_cbLexical = True Then
        spKey1 = UCase(spKey1)
        spKey2 = UCase(spKey2)
    End If
    
    If spKey1 = spKey2 Then
        CompareKeys = 0
    End If
    If spKey1 < spKey2 Then
        CompareKeys = -1
    End If
    If spKey1 > spKey2 Then
        CompareKeys = 1
    End If
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : RemoveNode
' DateTime  : 25/04/2005 20:20
' Author    : D K Richmond
' Purpose   : remove a node whether left, right or both sub branches, calls
'   RemoveNodeLeftOrRightOnly to deal with nodes with two sub branches
'    as well as ones with either left or right nodes
'---------------------------------------------------------------------------------------
'
Public Function RemoveNode(ByVal spKey As String) As Boolean
 Dim lCurrentNode As Long
 Dim bOk As Boolean
 Dim iOperation As Integer
 Dim sOrgKey As String
 Dim sPrevKey As String
 Dim bFound As Boolean
 Dim sHighestKey As String
 Dim lOtherNode As Long
 Dim lLeft As Long
 Dim lRight As Long
 Dim lParent As Long
 Dim sKey As String
 Dim lNextNode As Long
 Dim bNodeFound As Boolean
 Dim sLastkeyVisited As String
 Dim sNextKey As String
 Dim sCurrentKey As String
 Dim sOtherKey As String
 Dim iR As Integer
 
 ' check if root node is target for deletion and if so release it
 ' this rebalances the tree and hence the root node will have moved
 ' this means a new search .. it may end up as a leaf or a node with
 ' two sub branches
 ' check the target node for deletion has one, or two branches
 ' if it has two then .. then we need to move a node into it's position
 ' Search down the first left then all right or first right then all left
 ' randomly choose the left or the right to try not to unbalance the tree.
  
RemoveNode = False
                
    sKey = spKey
    bOk = SearchBSTN(sKey, lCurrentNode)
                
    If lCurrentNode = 0 Then   'if root node.. rebalance means location moves
        MsgBox "Deletion or root node means rebalance"
        bOk = ReleaseRootNode()
        If bOk = False Then
            Exit Function
        Else
            bOk = SearchBSTN(sKey, lCurrentNode)
        End If
    End If

    sOrgKey = GetKeyAtNodeIndex(lCurrentNode)
    lLeft = GetLeftNodeOfIndex(lCurrentNode)
    lRight = GetRightNodeOfIndex(lCurrentNode)
    lOtherNode = lCurrentNode
    
    ' ############# check for deletion of node having two child branches #############
    bFound = False
    If (lLeft <> m_cNoLink) And (lRight <> m_cNoLink) Then

        If gbDebug = True Then
            MsgBox "has left and right branches", vbInformation
        End If
        If bOk = True Then
            ' find largest key lower than current
            If bOk = True Then
                bOk = GetPrevInSortedOrder(lCurrentNode, lNextNode)
                If bOk = True Then
                    sPrevKey = GetKeyAtNodeIndex(lNextNode)
                    'MsgBox "Previous of " & sKey & " is " & sPrevKey
                Else
                    MsgBox "Failed!"
                    Err.Raise vbObjectError + 93049404, "", ""
                    Exit Function
                End If
                If gbDebug = True Then
                    MsgBox "**Highest left leaf from : " & sOrgKey & " is " & sPrevKey, vbInformation
                End If
                Debug.Print "**Highest left leaf from : " & sOrgKey & " is " & sPrevKey, vbInformation
            End If
        Else
            MsgBox "Failed!"
            Err.Raise vbObjectError + 93049404, "", ""
            Exit Function
        End If
        
        'If bFound = False Then
        If 1 > 0 Then
            ' find smallest key higher than current
            sKey = spKey
            bOk = SearchBSTN(sKey, lCurrentNode)
            sCurrentKey = GetKeyAtNodeIndex(lCurrentNode)
            Debug.Print "CK: " & sCurrentKey
            If bOk = True Then
                bOk = GetNextInSortedOrder(lCurrentNode, lNextNode)
                If bOk = True Then
                    sNextKey = GetKeyAtNodeIndex(lNextNode)
                Else
                    MsgBox "Failed!"
                    Err.Raise vbObjectError + 93049404, "", ""
                    Exit Function
                End If
                'MsgBox "Next of " & sKey & " is " & sNextKey
            End If
    
            If gbDebug = True Then
                MsgBox "**Lowest right leaf from : " & sOrgKey & " is " & sNextKey, vbInformation
            End If
            Debug.Print "**Lowest right leaf from : " & sOrgKey & " is " & sNextKey, vbInformation
        Else
            MsgBox "Failed!"
        End If
        
        'randomly decide on left or right substitution on deleted node position
        iR = Int(Rnd(1) * 2)
        If (iR = 1) Then
            sOtherKey = sNextKey
            If gbDebug = True Then
                MsgBox "random right" & "Key: " & sOrgKey & " Moving " & sOtherKey
            End If
            Debug.Print "random right: " & "Key: " & sOrgKey & " Moving " & sOtherKey
        Else
            sOtherKey = sPrevKey
            If gbDebug = True Then
                MsgBox "random left" & "Key: " & sOrgKey & " Moving " & sOtherKey
            End If
            Debug.Print "random left: " & "Key: " & sOrgKey & " Moving " & sOtherKey
        End If
        
        ' because we found the largest left or the smallest right
        ' it CANNOT have two branches as one of them would have to be
        ' the key we are trying to delete.  Hence we only have to deal
        ' with one of the branches when we move the node to the current
        ' key node being deleted
        
        ' now get the node index for the other node
        ' #### this section is to cater for things as yet unknown
        '       so far nothing has been discovered and soon this will
        '         disappear   #################
        bOk = SearchBSTN(sOtherKey, lOtherNode)
        If bOk = False Then
            MsgBox "Failed!"
            Err.Raise vbObjectError + 93049404, "", ""
            Exit Function
        End If
            
        If lCurrentNode = lOtherNode Then
          MsgBox "no can do!"
          Exit Function
          Err.Raise vbObjectError + 3940349, "", ""
        End If

        If lCurrentNode = 0 Then
          MsgBox "no can do on root moving!"
          Exit Function
          Err.Raise vbObjectError + 3940349, "", ""
        End If

        If lOtherNode = 0 Then
          MsgBox "no can do on root!"
          Exit Function
          Err.Raise vbObjectError + 3940349, "", ""
      End If
    
      ' end section #################
        
      If gbDebug = True Then
        MsgBox "only check for either left or right branches for the moving node", vbInformation
      End If
      
      bOk = RemoveNodeLeftOrRightOnly(sOtherKey, lCurrentNode)
      If bOk = False Then
         MsgBox "failed to remove node to be moved!"
         Err.Raise vbObjectError + 30493, "", ""
      End If
    
      ' this moves the node from the chosen to the deleted node key
      '     L & R & P links remain the same
      BSTNStore(lCurrentNode).m_sNodeKey = sOtherKey
      
      RemoveNode = True
    
    Else
        
        If gbDebug = True Then
            MsgBox "only check for either left or right branches", vbInformation
        End If
        
        bOk = RemoveNodeLeftOrRightOnly(sKey, lCurrentNode)
        
        If bOk = False Then
            MsgBox "failed to remove node to be moved!"
            Err.Raise vbObjectError + 30493, "", ""
        Else
            RemoveNode = True
        End If
End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : RemoveNodeLeftOrRightOnly
' DateTime  : 25/04/2005 20:29
' Author    : D K Richmond
' Purpose   : private - only usually call from RemoveNode
'---------------------------------------------------------------------------------------
'
Private Function RemoveNodeLeftOrRightOnly(ByVal spKey As String, ByVal plCurrent As Long) As Boolean
 Dim lCurrentNode As Long
 Dim bOk As Boolean
 Dim sOrgKey As String
 Dim bFound As Boolean
 Dim lOtherNode As Long
 Dim lLeft As Long
 Dim lRight As Long
 Dim lParent As Long
 Dim sKey As String
 Dim sOtherKey As String

 ' check if node for removal is a left or right sub branch and act accordingly
 
RemoveNodeLeftOrRightOnly = False

    lCurrentNode = plCurrent
    sKey = spKey
    bOk = SearchBSTN(sKey, lCurrentNode)
                
    If lCurrentNode = 0 Then   'if root node.. rebalance means location moves
        MsgBox "Deletion or root node means rebalance"
        bOk = ReleaseRootNode
        If bOk = False Then
            Exit Function
        Else
            ' rescan to find current node that moved during re-balance
            bOk = SearchBSTN(sKey, lCurrentNode)
        End If
    End If

    sOrgKey = GetKeyAtNodeIndex(lCurrentNode)
    lLeft = GetLeftNodeOfIndex(lCurrentNode)
    lRight = GetRightNodeOfIndex(lCurrentNode)
    lOtherNode = lCurrentNode
    
    ' ############# check for deletion node having two child branches #############
    bFound = False
    If (lLeft <> m_cNoLink) And (lRight <> m_cNoLink) Then

        MsgBox "has left and right branches this routine is not used for those nodes !!", vbInformation
        Err.Raise vbObjectError + 3494, "", ""
        Exit Function
    End If
    
    If gbDebug = True Then
        MsgBox "only check for either left or right branches", vbInformation
    End If
        
    ' ############# check for deletion node having left child branches #############
    If (lLeft <> m_cNoLink) And (lRight = m_cNoLink) Then
        ' deletion of node with left branch
        ' NOTE: ORDER OF OPERATIONS CRITICAL
        ' if CP.Left is Cn then CP.Left = On   or..
        ' if CP.Right is Cn then CP.Right = On
        ' if OP <> m_cNoLink then OP = CP
        lOtherNode = lLeft
        ' sort out node we are trying to remove
        sKey = GetKeyAtNodeIndex(lOtherNode)  ' debug
        sKey = GetKeyAtNodeIndex(lCurrentNode) ' debug
        ' if CP.Left is Cn then CP.Left = On    or..
        ' if CP.Right is Cn then CP.Right = On
        lParent = BSTNStore(lCurrentNode).m_lParent
        If lParent <> -1 Then Debug.Print "cP=" & GetKeyAtNodeIndex(lParent)
        If lParent <> m_cNoLink Then
            If BSTNStore(lParent).m_lLeft = lCurrentNode Then
                BSTNStore(lParent).m_lLeft = lOtherNode
            End If   ' or ...
            If BSTNStore(lParent).m_lRight = lCurrentNode Then
                BSTNStore(lParent).m_lRight = lOtherNode
            End If
            ' OP = CP
            BSTNStore(lOtherNode).m_lParent = lParent
        End If
        
        BSTNStore(lCurrentNode).m_lLeft = m_cNoLink  ' reset n's left branch
        BSTNStore(lCurrentNode).m_lRight = m_cNoLink   ' reset n's right branch
        BSTNStore(lCurrentNode).m_lParent = m_cNoLink   ' reset n's parent
        BSTNStore(lCurrentNode).m_sNodeKey = m_sGarbageIdentityKey  ' reset n's key to empty
        
        RemoveNodeLeftOrRightOnly = True
      End If
      
      ' ############# check for deletion node having right child branches #############
      If (lLeft = m_cNoLink) And (lRight <> m_cNoLink) Then
            ' deletion of node with left branch
            ' NOTE: ORDER OF OPERATIONS CRITICAL
            ' if CP.Left is Cn then CP.Left = On   or..
            ' if CP.Right is Cn then CP.Right = On
            ' if OP <> m_cNoLink then OP = CP
            lOtherNode = lRight
            ' sort out node we are trying to remove
            sKey = GetKeyAtNodeIndex(lOtherNode)  ' debug
            sKey = GetKeyAtNodeIndex(lCurrentNode) ' debug
            ' if CP.Left is Cn then CP.Left = On    or..
            ' if CP.Right is Cn then CP.Right = On
            lParent = BSTNStore(lCurrentNode).m_lParent
            If lParent <> m_cNoLink Then Debug.Print "cP=" & GetKeyAtNodeIndex(lParent)
            If lParent <> m_cNoLink Then
                If BSTNStore(lParent).m_lLeft = lCurrentNode Then
                    BSTNStore(lParent).m_lLeft = lOtherNode
                End If   ' or ...
                If BSTNStore(lParent).m_lRight = lCurrentNode Then
                    BSTNStore(lParent).m_lRight = lOtherNode
                End If
                ' OP = CP
                BSTNStore(lOtherNode).m_lParent = lParent
            End If
            
            BSTNStore(lCurrentNode).m_lLeft = m_cNoLink  ' reset n's left branch
            BSTNStore(lCurrentNode).m_lRight = m_cNoLink   ' reset n's right branch
            BSTNStore(lCurrentNode).m_lParent = m_cNoLink   ' reset n's parent
            BSTNStore(lCurrentNode).m_sNodeKey = m_sGarbageIdentityKey  ' reset n's key to empty
            
            RemoveNodeLeftOrRightOnly = True
      End If
      
      ' ############# check for deletion node being a leaf #############
      If (lLeft = m_cNoLink) And (lRight = m_cNoLink) Then
            ' deletion of node with left branch
            ' NOTE: ORDER OF OPERATIONS CRITICAL
            ' if OP.Left is Cn then OP.Left = m_cNoLink    or..
            ' if OP.Right is Cn then OP.Right = m_cNoLink
            lOtherNode = BSTNStore(lCurrentNode).m_lParent
            ' sort out node we are trying to remove
            If lOtherNode <> m_cNoLink Then sKey = GetKeyAtNodeIndex(lOtherNode)  ' debug
            If lOtherNode <> m_cNoLink Then sKey = GetKeyAtNodeIndex(lCurrentNode) ' debug
            ' if OP.Left is Cn then OP.Left = m_cNoLink    or..
            ' if OP.Right is Cn then OP.Right = m_cNoLink
            If lOtherNode <> m_cNoLink Then Debug.Print "cP=" & GetKeyAtNodeIndex(lParent)
            If lOtherNode <> m_cNoLink Then
                If BSTNStore(lOtherNode).m_lLeft = lCurrentNode Then
                    BSTNStore(lOtherNode).m_lLeft = m_cNoLink
                End If   ' or ...
                If BSTNStore(lOtherNode).m_lRight = lCurrentNode Then
                    BSTNStore(lOtherNode).m_lRight = m_cNoLink
                End If
            End If
            
            BSTNStore(lCurrentNode).m_lLeft = m_cNoLink  ' reset n's left branch
            BSTNStore(lCurrentNode).m_lRight = m_cNoLink   ' reset n's right branch
            BSTNStore(lCurrentNode).m_lParent = m_cNoLink   ' reset n's parent
            BSTNStore(lCurrentNode).m_sNodeKey = m_sGarbageIdentityKey  ' reset n's key to empty
            
            RemoveNodeLeftOrRightOnly = True
        
     End If

End Function

'---------------------------------------------------------------------------------------
' Procedure : ValidateTree
' DateTime  : 25/04/2005 20:31
' Author    : D K Richmond
' Purpose   : for debugging
'---------------------------------------------------------------------------------------
'
Public Function ValidateTree()
 Dim lMax As Long
 Dim lIndex As Long
 Dim lLeft As Long
 Dim lRight As Long
 Dim lParent As Long
 Dim sKey As String
 Dim sCurrentKey As String
 
   On Error GoTo ValidateTree_Error

    lMax = GetNodesMaxIndex() - 1
    For lIndex = 0 To lMax
       sCurrentKey = GetKeyAtNodeIndex(lIndex)
       lLeft = BSTNStore(lIndex).m_lLeft
       lRight = BSTNStore(lIndex).m_lRight
       lParent = BSTNStore(lIndex).m_lParent
       If lLeft <> m_cNoLink Then
           If BSTNStore(lLeft).m_lParent <> lIndex Then
               sKey = GetKeyAtNodeIndex(lLeft)
               MsgBox "Error in Left Key: " & sKey
           End If
       End If
       If lRight <> m_cNoLink Then
           If BSTNStore(lRight).m_lParent <> lIndex Then
               sKey = GetKeyAtNodeIndex(lRight)
               MsgBox "Error in Right Key: " & sKey
           End If
       End If
    Next
    

   On Error GoTo 0
   Exit Function

ValidateTree_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ValidateTree of Class Module cBSTNStorage"

End Function

' ##############################################################################################################
' ############################# functions below here are in re-design in order to be removed ###################
' #############  or be re-established as part of the main project ##############################################
' ##############################################################################################################


Function GetPrevInSortedOrder(ByVal lpCurrentNode As Long, ByRef lpNextNode As Long, Optional bpoPassThru As Boolean = False, Optional lpoPassThruNode As Long = -1) As Boolean
' PLEASE DON'T MAKE TO MANY REFERENCES TO THIS MEMBER FUNCTION AS IT IS LIKELY TO BE SUPERSEEDED IN LATER RELEASES.
'------------------------------------------
' -- Function: GetPrevInSortedOrder
' -- Desc    : Get Previous Node In Sorted Order
' -- Passed  : nothing
' -- Uses    :
' -- Calls   : other functions
' -- Returns : 0 success, -1 fail
' -- DateTime: Org - 20:17 24/05/00
'------------------------------------------

 Dim wNodeType As Integer
 Dim wPrevNodeType As Integer
 Dim lRight As Long
 Dim lParent As Long
 Dim lLeft As Long
 
    GetPrevInSortedOrder = False

    wNodeType = GetNodeType(lpCurrentNode)
 
    If (wNodeType = m_eRNODE_TYPE) Then
        lLeft = GetLeftNodeOfIndex(lpCurrentNode)
        If (lLeft <> m_cNoLink) Then
            ' left exists
            lpCurrentNode = lLeft
            If lpoPassThruNode = lpCurrentNode Then _
                                            bpoPassThru = True
            Do
            lRight = GetRightNodeOfIndex(lpCurrentNode)
            If (lRight = m_cNoLink) Then
                Exit Do
            Else
                ' right exists
                lpCurrentNode = lRight
                If lpoPassThruNode = lpCurrentNode Then _
                                            bpoPassThru = True
            End If
            Loop
        Else
            lParent = GetParentNodeOfIndex(lpCurrentNode)
            If (lParent <> m_cNoLink) Then
                lpCurrentNode = lParent
                If lpoPassThruNode = lpCurrentNode Then _
                            bpoPassThru = True
            End If
        End If
    Else
     If (wNodeType = m_eLNODE_TYPE) Or (wNodeType = m_ePNODE_TYPE) Then
        ' go left if any exist then all left
        lLeft = GetLeftNodeOfIndex(lpCurrentNode)
        If (lLeft <> m_cNoLink) Then
            lpCurrentNode = lLeft
            If lpoPassThruNode = lpCurrentNode Then _
                            bpoPassThru = True
            Do
                lRight = GetRightNodeOfIndex(lpCurrentNode)
                If (lRight <> m_cNoLink) Then
                    ' right exists
                    lpCurrentNode = lRight
                    If lpoPassThruNode = lpCurrentNode Then _
                            bpoPassThru = True
                Else
                    Exit Do
                End If
            Loop
        Else
           Do
                wNodeType = GetNodeType(lpCurrentNode)
                lParent = GetParentNodeOfIndex(lpCurrentNode)
                If (lParent <> m_cNoLink) Then
                    lpCurrentNode = lParent
                    If lpoPassThruNode = lpCurrentNode Then _
                            bpoPassThru = True
                    wPrevNodeType = wNodeType
                Else
                    Exit Do
                End If
           Loop Until wNodeType <> m_eLNODE_TYPE
           '  If dw_GetCurrentPNodeRecNo() = 0 Then         ' at root
           '     MsgBox "underrun sot!"
           '     wResult = -98
           ' End If
           
        End If
     End If
     
        If (lpCurrentNode = 0) And (wNodeType = m_ePNODE_TYPE) Then
            MsgBox "Gone to root!"
              GetPrevInSortedOrder = False
              Exit Function
        End If
        
'     If dw_GetCurrentPNodeRecNo() = 0 And wPrevNodeType = gcLNODE_TYPE Then         ' at root
 '       MsgBox "underrun sot!"
  '      wResult = -98
  '   End If
  End If
lpNextNode = lpCurrentNode
  
  GetPrevInSortedOrder = True

End Function

'------------------------------------------
' -- Function: GetNextInSortedOrder
' -- Passed  : nothing
' -- Uses    :
' -- Calls   : other functions using and updating globals
' -- Returns : 0 success, -1 fail
' -- DateTime: 20:17 24/05/00
'------------------------------------------
' PLEASE DON'T MAKE TO MANY REFERENCES TO THIS MEMBER FUNCTION AS IT IS LIKELY TO BE SUPERSEEDED IN LATER RELEASES.

Function GetNextInSortedOrder(ByVal lpCurrentNode As Long, ByRef lpNextNode As Long) As Boolean
 
 Dim lRight As Long
 Dim lParentNode  As Long
 Dim lParent As Long
 Dim lLeft As Long
 Dim wPrevNodeType As Integer
 Dim wNodeType As Integer
 Dim lPrevNode As Long
 
    GetNextInSortedOrder = False

    ' next in sorted alpha data order
    lPrevNode = -1
    ' go right if any exist then all left
    lRight = GetRightNodeOfIndex(lpCurrentNode)
    If (lRight <> m_cNoLink) Then
        ' right exists
        lpCurrentNode = lRight
        Do
            lLeft = GetLeftNodeOfIndex(lpCurrentNode)
            If (lLeft <> m_cNoLink) Then
                lpCurrentNode = lLeft
            Else
                Exit Do
            End If
        Loop
    Else
        Do
            wNodeType = GetNodeType(lpCurrentNode)
            lParent = GetParentNodeOfIndex(lpCurrentNode)
            If (lParent <> m_cNoLink) Then
                lPrevNode = lpCurrentNode
                lpCurrentNode = lParent
                wPrevNodeType = wNodeType
            Else
                Exit Do
            End If
        Loop Until wNodeType <> m_eRNODE_TYPE
        wNodeType = GetNodeType(lPrevNode)
        If (lpCurrentNode = 0) And (wNodeType = m_eRNODE_TYPE) Then
            MsgBox "Gone to root!"
              GetNextInSortedOrder = False
              Exit Function
        End If

    '     If dw_GetCurrentPNodeRecNo() = 0 Then    ' at root
    '        MsgBox "overrun eot!"
    '        wResult = -99 ' gcFail
    '     End If
    End If

 lpNextNode = lpCurrentNode
  GetNextInSortedOrder = True

End Function


' <eoc>
