VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMainMediaLibrary 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Media Library (Beta) Not For Release (C) David K Richmond"
   ClientHeight    =   6180
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   10635
   Icon            =   "frmMainMediaLibrary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   10635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   8880
      TabIndex        =   30
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton cmdMediaDelete 
      Caption         =   "&Delete Media"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3960
      TabIndex        =   29
      Top             =   5160
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   9600
      Picture         =   "frmMainMediaLibrary.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   28
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdMediaEdit 
      Caption         =   "&Edit Media"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   27
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdShowAll 
      Caption         =   "Show All"
      Height          =   375
      Left            =   240
      TabIndex        =   26
      Top             =   4680
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   480
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdUpdateMedia 
      Caption         =   "&Update Media"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3960
      TabIndex        =   23
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdSearchMedia 
      Caption         =   "&Search Media"
      Height          =   375
      Left            =   240
      TabIndex        =   22
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton cmdAddMedia 
      Caption         =   "&Add Media"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   21
      Top             =   5160
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   2640
      TabIndex        =   20
      Top             =   4080
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   2640
      TabIndex        =   19
      Top             =   3480
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   2640
      TabIndex        =   18
      Top             =   2880
      Width           =   3615
   End
   Begin VB.Frame frMediaType 
      Caption         =   "Media Type"
      Enabled         =   0   'False
      Height          =   2295
      Left            =   6480
      TabIndex        =   7
      Top             =   2880
      Width           =   3735
      Begin VB.OptionButton OptMediaVHS 
         Caption         =   "VHS Video"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1320
         TabIndex        =   25
         Top             =   1680
         Width           =   1000
      End
      Begin VB.OptionButton OptMediaPC 
         Caption         =   "PC Stuff"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   2640
         TabIndex        =   14
         Top             =   960
         Width           =   1000
      End
      Begin VB.OptionButton OptMediaPS2Game 
         Caption         =   "PS2 Game"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1320
         TabIndex        =   13
         Top             =   960
         Width           =   1000
      End
      Begin VB.OptionButton OptMediaPS1Game 
         Caption         =   "PS1 Game"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   360
         TabIndex        =   12
         Top             =   960
         Width           =   1000
      End
      Begin VB.OptionButton OptMediaAudioTape 
         Caption         =   "Audio Tape"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   360
         TabIndex        =   11
         Top             =   1680
         Width           =   1000
      End
      Begin VB.OptionButton OptMediaBook 
         Caption         =   "Book"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   2640
         TabIndex        =   10
         Top             =   360
         Width           =   1000
      End
      Begin VB.OptionButton OptMediaDVD 
         Caption         =   "DVD"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   1320
         TabIndex        =   9
         Top             =   360
         Width           =   1000
      End
      Begin VB.OptionButton OptMediaCD 
         Caption         =   "CD"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Width           =   1000
      End
   End
   Begin VB.TextBox txtPublisher 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   2640
      TabIndex        =   6
      Top             =   2280
      Width           =   7575
   End
   Begin VB.TextBox txtTitle 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   2640
      TabIndex        =   5
      Top             =   1680
      Width           =   7575
   End
   Begin VB.TextBox txtMediaIdentifier 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   2640
      TabIndex        =   4
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Label lblLoadingStatus 
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   5760
      Width           =   10575
   End
   Begin VB.Label Label8 
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   17
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label Label7 
      Caption         =   "Supplier:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   16
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label Label6 
      Caption         =   "Cost:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   15
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Publisher:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Title:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Media Identifier:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Sample Media Library (Beta)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   240
      Width           =   4695
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Begin VB.Menu mnuNewLibrary 
         Caption         =   "New Library"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuLoadLibrary 
         Caption         =   "Load Library"
      End
      Begin VB.Menu mnuSaveLibrary 
         Caption         =   "Save Library"
      End
      Begin VB.Menu mnuSpacer2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuShow 
      Caption         =   "Show"
      Begin VB.Menu mnuShowAll 
         Caption         =   "Show All"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMainMediaLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmMainMediaLibrary
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

Option Explicit

' media types
Public Enum MediaTypes
    egMediaIdentifier = 0
    egMediaTitle = 1
    egMediaPublisher = 2
    egMediaType = 3
    egMediaEOD = 4
End Enum


'---------------------------------------------------------------------------------------
' Procedure : cmdAddMedia_Click
' DateTime  : 05/05/2005 19:55
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdAddMedia_Click()
 Dim saRecordFields() As String
 Dim bOk As Boolean
 Dim sKeyValue As String
 
   On Error GoTo cmdAddMedia_Click_Error

    cmdAddMedia.Enabled = False
    
    ReDim saRecordFields(egMediaEOD)
    sKeyValue = txtMediaIdentifier.Text
    saRecordFields(0) = txtMediaIdentifier.Text
    saRecordFields(egMediaTitle) = txtTitle.Text
    saRecordFields(egMediaPublisher) = txtPublisher.Text
    
    SetMediaTypeFromOptionBoxes saRecordFields(egMediaType)
     
    saRecordFields(egMediaEOD) = "-- unknown --"
    
    bOk = AddMediaToLibrary(sKeyValue, saRecordFields(), lblLoadingStatus)
    
    If bOk = False Then
        MsgBox "Media " & sKeyValue & " No Data Added!", vbCritical
        Exit Sub
    Else
        frMediaType.Enabled = False
        cmdMediaEdit.Enabled = True
        txtPublisher.Enabled = False
        txtTitle.Enabled = False
        MsgBox "Media " & sKeyValue & " Data Added!", vbInformation
        SetDirtyFlag
    End If

   On Error GoTo 0
   Exit Sub

cmdAddMedia_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAddMedia_Click of Form frmMainMediaLibrary"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SetDirtyFlag
' DateTime  : 05/05/2005 19:55
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub SetDirtyFlag()

    mnuSaveLibrary.Enabled = True
    bgDirtyFlag = True
    bgRefreshView = True
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : GetDirtyFlag
' DateTime  : 05/05/2005 19:56
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function GetDirtyFlag() As Boolean

    GetDirtyFlag = bgDirtyFlag

End Function

Private Sub cmdExit_Click()
    Unload Me
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdMediaDelete_Click
' DateTime  : 05/05/2005 19:56
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdMediaDelete_Click()
 Dim bOk As Boolean
 Dim sKeyValue As String
 Dim lDataIndex As Long
 
   On Error GoTo cmdMediaDelete_Click_Error

    cmdMediaDelete.Enabled = False
    
    sKeyValue = txtMediaIdentifier.Text
    
    bOk = DeleteMedia(sKeyValue, lDataIndex)
    
    If bOk = False Then
        MsgBox "Media " & sKeyValue & " Not Deleted!", vbExclamation
        Exit Sub
    End If
    
    MsgBox "Media Record @" & lDataIndex & " Key: " & sKeyValue & " Deleted!", vbInformation
    
    SetDirtyFlag

    frMediaType.Enabled = False
    
   On Error GoTo 0
   Exit Sub

cmdMediaDelete_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdMediaDelete_Click of Form frmMainMediaLibrary"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdMediaEdit_Click
' DateTime  : 05/05/2005 19:56
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdMediaEdit_Click()
    
   On Error GoTo cmdMediaEdit_Click_Error

    cmdMediaEdit.Enabled = False
    cmdAddMedia.Enabled = True
    txtMediaIdentifier.Enabled = True
    txtPublisher.Enabled = True
    txtTitle.Enabled = True
    frMediaType.Enabled = True
    cmdUpdateMedia.Enabled = True
    
   On Error GoTo 0
   Exit Sub

cmdMediaEdit_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdMediaEdit_Click of Form frmMainMediaLibrary"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdSearchMedia_Click
' DateTime  : 05/05/2005 19:56
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdSearchMedia_Click()
 Dim sKeyValue As String
 Dim lDataIndex As Long
 Dim saDataValues() As String
 Dim bOk As Boolean
 
   On Error GoTo cmdSearchMedia_Click_Error

    cmdAddMedia.Enabled = False
    cmdMediaDelete.Enabled = False
    cmdSearchMedia.Enabled = False
    cmdUpdateMedia.Enabled = False
    
    sKeyValue = txtMediaIdentifier.Text
    
    bOk = FindMedia(sKeyValue, lDataIndex)
    
    If bOk = False Then
        MsgBox "Media " & sKeyValue & " Not Found!", vbExclamation
        cmdSearchMedia.Enabled = True
        Exit Sub
    End If

    bOk = GetMediaRecordDataValues(lDataIndex, saDataValues())
    
    If bOk = False Then
        MsgBox "Media " & sKeyValue & " No Data Found!", vbExclamation
        cmdSearchMedia.Enabled = True
        Exit Sub
    End If
    
    txtMediaIdentifier.Text = saDataValues(egMediaIdentifier)  ' needs to be the key global value
    txtPublisher.Text = saDataValues(egMediaPublisher)
    txtTitle.Text = saDataValues(egMediaTitle)
    
    SetMediaTypeOptionBoxes saDataValues(egMediaType)
    
    
    cmdMediaDelete.Enabled = True
    cmdMediaEdit.Enabled = True
    cmdSearchMedia.Enabled = True

   On Error GoTo 0
   Exit Sub

cmdSearchMedia_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdSearchMedia_Click of Form frmMainMediaLibrary"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SetMediaTypeOptionBoxes
' DateTime  : 05/05/2005 19:56
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub SetMediaTypeOptionBoxes(ByVal spMediaType As String)
  
   On Error GoTo SetMediaTypeOptionBoxes_Error

  ' centralised
    OptMediaCD.Value = False
    OptMediaPS1Game.Value = False
    OptMediaPS2Game.Value = False
    OptMediaAudioTape.Value = False
    OptMediaDVD.Value = False
    OptMediaPS2Game.Value = False
    OptMediaBook.Value = False
    OptMediaPC.Value = False
    OptMediaVHS.Value = False
    
    Select Case spMediaType
        Case "CD Music"
            OptMediaCD = True
        Case "DVD"
            OptMediaDVD = True
        Case "Book"
            OptMediaBook = True
        Case "PS1 Game"
            OptMediaPS1Game = True
        Case "PS2 Game"
            OptMediaPS2Game = True
        Case "PC Game"
            OptMediaPC = True
        Case "Audio Tape"
            OptMediaAudioTape = True
        Case "VHS Video"
            OptMediaVHS.Value = True
    End Select

   On Error GoTo 0
   Exit Sub

SetMediaTypeOptionBoxes_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SetMediaTypeOptionBoxes of Form frmMainMediaLibrary"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdShowAll_Click
' DateTime  : 05/05/2005 19:56
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdShowAll_Click()
        
   On Error GoTo cmdShowAll_Click_Error

        Me.cmdAddMedia.Enabled = False
        Me.cmdMediaDelete.Enabled = False
        Me.cmdUpdateMedia.Enabled = False
        frmShowAllMedia.Show

   On Error GoTo 0
   Exit Sub

cmdShowAll_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdShowAll_Click of Form frmMainMediaLibrary"

End Sub

Private Sub cmdUpdateMedia_Click()
 Dim saRecordFields() As String
 Dim bOk As Boolean
 Dim sKeyValue As String
 
   On Error GoTo cmdUpdateMedia_Click_Error

    cmdUpdateMedia = False
    
    ReDim saRecordFields(egMediaEOD)
    sKeyValue = txtMediaIdentifier.Text
    saRecordFields(egMediaIdentifier) = txtMediaIdentifier.Text
    saRecordFields(egMediaPublisher) = txtPublisher.Text
    saRecordFields(egMediaTitle) = txtTitle.Text
    
    SetMediaTypeFromOptionBoxes saRecordFields(egMediaType)
    
    saRecordFields(egMediaEOD) = "-- unknown --"
    
    bOk = UpdateMediaToLibrary(sKeyValue, saRecordFields(), lblLoadingStatus)
    
    cmdUpdateMedia.Enabled = True

    If bOk = False Then
        MsgBox "Media " & sKeyValue & " No Data Updated!", vbCritical
        Exit Sub
    Else
        cmdMediaEdit.Enabled = True
        txtPublisher.Enabled = False
        txtTitle.Enabled = False
        cmdMediaEdit.Enabled = False
        cmdUpdateMedia.Enabled = False
        frMediaType.Enabled = False
        MsgBox "Media " & sKeyValue & " Data Updated!", vbInformation
        SetDirtyFlag
    End If

   On Error GoTo 0
   Exit Sub

cmdUpdateMedia_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdUpdateMedia_Click of Form frmMainMediaLibrary"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : SetMediaTypeFromOptionBoxes
' DateTime  : 05/05/2005 19:56
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub SetMediaTypeFromOptionBoxes(ByRef spRecordField As String)

   On Error GoTo SetMediaTypeFromOptionBoxes_Error

    If OptMediaPS1Game.Value = True Then
        spRecordField = "PS1 Game"
    End If
    If OptMediaAudioTape.Value = True Then
        spRecordField = "Audio Tape"
    End If
    If OptMediaBook.Value = True Then
        spRecordField = "Book"
    End If
    If OptMediaCD.Value = True Then
        spRecordField = "CD Music"
    End If
    If OptMediaDVD.Value = True Then
        spRecordField = "DVD"
    End If
    If OptMediaPC.Value = True Then
        spRecordField = "PC Game"
    End If
    If OptMediaPS2Game.Value = True Then
        spRecordField = "PS2 Game"
    End If
    If OptMediaVHS.Value = True Then
        spRecordField = "VHS Tape"
    End If

   On Error GoTo 0
   Exit Sub

SetMediaTypeFromOptionBoxes_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SetMediaTypeFromOptionBoxes of Form frmMainMediaLibrary"
    
End Sub
    
'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' DateTime  : 05/05/2005 19:56
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()
    
   On Error GoTo Form_Load_Error

    Call InitialiseLibrary
     
    mnuLoadLibrary.Enabled = False
    mnuSaveLibrary.Enabled = False
     
    mnuLoadLibrary_Click
    
    cmdAddMedia.Enabled = False

    bgRefreshView = True

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form frmMainMediaLibrary"
    
End Sub

 Private Sub Form_Unload(Cancel As Integer)
  Dim sYN As String
  Dim bOk As Boolean
  Dim bQuit As Boolean

   On Error GoTo Form_Unload_Error

     bOk = True
     bQuit = False
     If GetDirtyFlag = True Then
         sYN = UCase(InputBox("Save Library (Y/N/Q)?", "Save", "Y"))
         If sYN = "Q" Then
             bQuit = True
         End If
         If sYN = "Y" Then
             bOk = SaveLibrary(lblLoadingStatus)
         End If
     End If
     If (bOk = True) Or (bQuit = True) Then
         MsgBox "Exit", vbInformation
         Unload frmShowAllMedia
         Set frmMainMediaLibrary = Nothing
     Else
         Cancel = 1
         MsgBox "Problem with Library.", vbCritical
     End If

   On Error GoTo 0
   Exit Sub

Form_Unload_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form frmMainMediaLibrary"
     
 End Sub

'---------------------------------------------------------------------------------------
' Procedure : mnuExit_Click
' DateTime  : 05/05/2005 19:56
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
 Private Sub mnuExit_Click()
     
     Unload Me
     
 End Sub

Private Sub mnuAbout_Click()
 Dim sAbout As String
 
    sAbout = "Media Library Sample 1.0 (BETA)" & vbCrLf
    sAbout = sAbout & "Written by David K Richmond April 2005" & vbCrLf
    sAbout = sAbout & "[ Uses Binary Search Tree Node Class ]" & vbCrLf
    sAbout = sAbout & "[ Demonstration of small media library CDs,DVDs etc ]" & vbCrLf
    sAbout = sAbout & " { APP Details: " & App.FileDescription & " " & App.Major & "." & App.Minor & " }"
    MsgBox sAbout, vbInformation
     
End Sub


'---------------------------------------------------------------------------------------
' Procedure : mnuLoadLibrary_Click
' DateTime  : 05/05/2005 19:56
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuLoadLibrary_Click()
 Dim bOk As Boolean
 
   On Error GoTo mnuLoadLibrary_Click_Error

  'MsgBox "Load", vbInformation
  
'  CommonDialog1.DefaultExt = "dat"
'  CommonDialog1.FileName = sgcLibraryFileName
'  CommonDialog1.InitDir = App.Path & "\DATA"
'  CommonDialog1.ShowOpen

    ' quick testing option
    sgLibraryFileName = sgcLibraryFileName
    sgLibraryFilePathName = App.Path & "\DATA\" & sgLibraryFileName
    
    bOk = LoadLibrary(lblLoadingStatus)
    
    If bOk = True Then
'        MsgBox "Library Loaded Ok!", vbInformation
    Else
        MsgBox "Failed to Load Library!", vbInformation
    End If

   On Error GoTo 0
   Exit Sub

mnuLoadLibrary_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuLoadLibrary_Click of Form frmMainMediaLibrary"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : mnuSaveLibrary_Click
' DateTime  : 05/05/2005 19:56
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuSaveLibrary_Click()
 Dim bOk As Boolean
 
   On Error GoTo mnuSaveLibrary_Click_Error

    bOk = SaveLibrary(lblLoadingStatus)
    
    If bOk = True Then
        MsgBox "Library Saved Ok!", vbInformation
        ' menu and globals
        mnuSaveLibrary.Enabled = False
        bgDirtyFlag = False
    Else
        MsgBox "Failed to SaveLibrary!", vbInformation
    End If
    

   On Error GoTo 0
   Exit Sub

mnuSaveLibrary_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuSaveLibrary_Click of Form frmMainMediaLibrary"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : mnuShowAll_Click
' DateTime  : 05/05/2005 19:56
' Author    : D K Richmond
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub mnuShowAll_Click()

    frmShowAllMedia.Show
    
End Sub


' eom
