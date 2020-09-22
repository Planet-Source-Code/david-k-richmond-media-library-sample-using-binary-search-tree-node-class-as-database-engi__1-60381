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

