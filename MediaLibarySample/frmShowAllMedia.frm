VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmShowAllMedia 
   Caption         =   "MediaLibraryShow All"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12300
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   12300
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRefreshView 
      Caption         =   "Refresh View"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdShowFind 
      Caption         =   "&Find"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   6960
      Width           =   1455
   End
   Begin VB.TextBox txtFindKeyValue 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2040
      TabIndex        =   2
      Text            =   "-- enter media key --"
      Top             =   6960
      Width           =   5535
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   10440
      TabIndex        =   1
      Top             =   7080
      Width           =   1695
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6255
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   11033
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "frmShowAllMedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmShowAll
' DateTime  : 04/05/2005 08:30
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

Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdRefreshView_Click
' DateTime  : 04/05/2005 07:29
' Author    : D K Richmond
' Purpose   : BUG fixed where rebalance caused null key in listview due to index being lost in InitrootNode
'---------------------------------------------------------------------------------------
'
Private Sub cmdRefreshView_Click()
 Dim lNextNode As Long
 Dim sKey As String
 Dim sKeyValue As String
 Dim sTextValue As String
 Dim lDataIndex As Long
 Dim saDataValues() As String
 Dim bOk As Boolean
 Dim lviItem As MSComctlLib.ListItem
 Dim iFields As Integer
 Dim iField As Integer
 Dim bfirst As Boolean
 
   On Error GoTo cmdRefreshView_Click_Error

    ListView1.Visible = False
    ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    
    clsgBSTStore.ResetToRootNode
    bfirst = True
    Do
        If clsgBSTStore.GetNextNodeInSortedOrder(lNextNode) = True Then
            lDataIndex = clsgBSTStore.GetDataNodeIndex(lNextNode)
            bOk = GetMediaRecordDataValues(lDataIndex, saDataValues())
            If bOk = False Then
                MsgBox "Media @node " & lNextNode & " No Data Found!", vbExclamation
                Exit Sub
            Else
                sKeyValue = saDataValues(egMediaIdentifier)
                sTextValue = saDataValues(egMediaIdentifier)
                Set lviItem = ListView1.ListItems.Add(, "K" & sKeyValue, sTextValue)
                If Not lviItem Is Nothing Then
                    iFields = UBound(saDataValues())
                    For iField = 0 To iFields
                        If bfirst = True Then
                            ListView1.ColumnHeaders.Add iField + 1, "K" & iField, "" & iField
                            ListView1.ColumnHeaders(iField + 1).Width = (iField * igcColumnSizing)
                        End If
                        lviItem.ListSubItems.Add , "K" & saDataValues(iField) & Format(iField, "000000"), saDataValues(iField)
                    Next
                End If
            End If
        Else
            Exit Do
        End If

        DoEvents

        If (clsgBSTStore.TopOfTreeReached = True) Then
            Exit Do
        End If
        bfirst = False
    Loop

    ListView1.Visible = True

    bgRefreshView = False
    
   On Error GoTo 0
   Exit Sub

cmdRefreshView_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdRefreshView_Click of Form frmShowAllMedia"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdShowFind_Click
' DateTime  : 05/05/2005 19:57
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdShowFind_Click()
Dim lviItem As MSComctlLib.ListItem
Dim sKeyFindValue As String

   On Error GoTo cmdShowFind_Click_Error

    sKeyFindValue = txtFindKeyValue.Text
    
    For Each lviItem In ListView1.ListItems
        If lviItem.Text = sKeyFindValue Then
            lviItem.Selected = True
            lviItem.EnsureVisible
            Exit For
        End If
    Next

   On Error GoTo 0
   Exit Sub

cmdShowFind_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdShowFind_Click of Form frmShowAllMedia"
End Sub

Private Sub Form_GotFocus()

    If bgRefreshView = True Then
        cmdRefreshView_Click
    End If
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' DateTime  : 05/05/2005 19:57
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()

    cmdRefreshView_Click
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If bgRefreshView = True Then
        cmdRefreshView_Click
    End If

End Sub

'---------------------------------------------------------------------------------------
' Procedure : ListView1_Click
' DateTime  : 05/05/2005 19:57
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub ListView1_Click()
 Dim lviItem As MSComctlLib.ListItem
 Dim iFields As Integer

   On Error GoTo ListView1_Click_Error

    Set lviItem = ListView1.SelectedItem
    If Not lviItem Is Nothing Then
        frmMainMediaLibrary.txtMediaIdentifier.Text = lviItem.ListSubItems(egMediaIdentifier + 1).Text
        frmMainMediaLibrary.txtPublisher.Text = lviItem.ListSubItems(egMediaPublisher + 1).Text
        frmMainMediaLibrary.txtTitle.Text = lviItem.ListSubItems(egMediaTitle + 1).Text
        frmMainMediaLibrary.SetMediaTypeOptionBoxes lviItem.ListSubItems(egMediaType + 1).Text
        frmMainMediaLibrary.cmdMediaDelete.Enabled = True
        frmMainMediaLibrary.cmdMediaEdit.Enabled = True
        frmMainMediaLibrary.Show
    Else
        MsgBox "select something first!", vbExclamation
    End If

   On Error GoTo 0
   Exit Sub

ListView1_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ListView1_Click of Form frmShowAllMedia"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ListView1_ColumnClick
' DateTime  : 05/05/2005 19:57
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
   On Error GoTo ListView1_ColumnClick_Error

    ListView1.Sorted = True
    ListView1.SortKey = ColumnHeader.Index - 1
    If ListView1.SortOrder = lvwAscending Then
        ListView1.SortOrder = lvwDescending
    Else
        ListView1.SortOrder = lvwAscending
    End If

   On Error GoTo 0
   Exit Sub

ListView1_ColumnClick_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ListView1_ColumnClick of Form frmShowAllMedia"
    
End Sub


' eom

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   
    If bgRefreshView = True Then
        cmdRefreshView_Click
    End If
    
End Sub
