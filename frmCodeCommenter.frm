VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCodeCommenter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Little Helper: Auto Insert Error Trapping and Commenting"
   ClientHeight    =   9405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCodeCommenter.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraProgress 
      Caption         =   "Progress"
      Height          =   1530
      Left            =   210
      TabIndex        =   32
      Top             =   7515
      Width           =   7590
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   360
         Left            =   150
         TabIndex        =   34
         Top             =   630
         Width           =   6270
         _ExtentX        =   11060
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   330
         Left            =   6585
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   630
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Valid Variables:"
      Height          =   2430
      Left            =   5655
      TabIndex        =   31
      Top             =   600
      Width           =   2025
      Begin VB.TextBox txtVariableTypes 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Left            =   60
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   26
         Text            =   "frmCodeCommenter.frx":0442
         Top             =   210
         Width           =   1890
      End
   End
   Begin VB.CheckBox chkComments 
      Caption         =   "Add comme&nts to file"
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   630
      Width           =   1815
   End
   Begin VB.CheckBox chkErrorHandling 
      Caption         =   "Add error &handling to file"
      Height          =   195
      Left            =   90
      TabIndex        =   12
      Top             =   4605
      Width           =   2145
   End
   Begin VB.Frame fraAction 
      Caption         =   "Action"
      Height          =   1545
      Left            =   90
      TabIndex        =   30
      Top             =   7770
      Width           =   7590
      Begin VB.ListBox lstFileTypes 
         Enabled         =   0   'False
         Height          =   960
         Left            =   5340
         Style           =   1  'Checkbox
         TabIndex        =   22
         Top             =   435
         Width           =   1170
      End
      Begin VB.CheckBox chkBackup 
         Caption         =   "Make bac&kup"
         Height          =   375
         Left            =   6615
         TabIndex        =   23
         Top             =   195
         Value           =   1  'Checked
         Width           =   840
      End
      Begin VB.CommandButton cmdQuit 
         Caption         =   "E&xit"
         Height          =   330
         Left            =   6615
         TabIndex        =   25
         Top             =   1065
         Width           =   855
      End
      Begin VB.CommandButton cmdGo 
         Caption         =   "&Go"
         Height          =   330
         Left            =   6615
         TabIndex        =   24
         Top             =   675
         Width           =   855
      End
      Begin VB.TextBox txtApplyTo 
         Height          =   945
         Left            =   1515
         TabIndex        =   20
         Top             =   435
         Width           =   3750
      End
      Begin VB.CommandButton cmdApplyToFile 
         Caption         =   "Apply To F&ile"
         Height          =   330
         Left            =   105
         TabIndex        =   18
         Top             =   930
         Width           =   1290
      End
      Begin VB.CommandButton cmdApplyToFolder 
         Caption         =   "Apply To F&older"
         Height          =   330
         Left            =   105
         TabIndex        =   17
         Top             =   540
         Width           =   1290
      End
      Begin VB.Label lblToFileTypes 
         Caption         =   "To File &Types:"
         Height          =   240
         Left            =   5325
         TabIndex        =   21
         Top             =   195
         Width           =   1335
      End
      Begin VB.Label lblActionCaption 
         Caption         =   "A&pply the selected items to the following:"
         Height          =   210
         Left            =   1515
         TabIndex        =   19
         Top             =   195
         Width           =   3840
      End
   End
   Begin MSComDlg.CommonDialog cdlgFile 
      Left            =   7365
      Top             =   -60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FraInstructions 
      Caption         =   "Note:"
      Height          =   480
      Left            =   90
      TabIndex        =   28
      Top             =   90
      Width           =   7590
      Begin VB.Label Label1 
         Caption         =   "You may modify the templates below (use Ctl+Enter instead of Enter to insert a Carriage Return)"
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   180
         TabIndex        =   29
         Top             =   210
         Width           =   7050
      End
   End
   Begin VB.Frame fraDescription 
      Caption         =   "Description of proceed&ure:"
      Height          =   1485
      Left            =   90
      TabIndex        =   8
      Top             =   3060
      Width           =   7590
      Begin VB.CommandButton cmdSaveDescriptionTemplateToFile 
         Cancel          =   -1  'True
         Caption         =   "Sa&ve As Default"
         Height          =   330
         Left            =   1350
         TabIndex        =   11
         Top             =   1020
         Width           =   1410
      End
      Begin VB.CommandButton cmdLoadDescriptionTemplateFromFile 
         Caption         =   "Load &From File"
         Height          =   330
         Left            =   105
         TabIndex        =   10
         Top             =   1020
         Width           =   1200
      End
      Begin VB.TextBox txtDescription 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   105
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   9
         Text            =   "frmCodeCommenter.frx":04E9
         Top             =   270
         Width           =   7365
      End
   End
   Begin VB.Frame fraErrorHandling 
      Height          =   3180
      Left            =   90
      TabIndex        =   27
      Top             =   4560
      Width           =   7590
      Begin VB.CommandButton cmdSaveErrorTemplateToFile 
         Caption         =   "Save As &Default"
         Height          =   330
         Left            =   1380
         TabIndex        =   16
         Top             =   2745
         Width           =   1410
      End
      Begin VB.CommandButton cmdLoadErrorTemplateFromFile 
         Caption         =   "Load Fro&m File"
         Height          =   330
         Left            =   120
         TabIndex        =   15
         Top             =   2745
         Width           =   1200
      End
      Begin VB.TextBox txtErrorHandling 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Left            =   105
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   14
         Text            =   "frmCodeCommenter.frx":051A
         Top             =   945
         Width           =   7365
      End
      Begin VB.TextBox txtErrorHandlingTop 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   105
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Text            =   "frmCodeCommenter.frx":0821
         Top             =   300
         Width           =   7365
      End
   End
   Begin VB.Frame fraComments 
      Height          =   2430
      Left            =   90
      TabIndex        =   0
      Top             =   600
      Width           =   5490
      Begin VB.OptionButton optPlaceCommentsWhere 
         Caption         =   "&After Declaration"
         Height          =   195
         Index           =   1
         Left            =   3765
         TabIndex        =   3
         Top             =   30
         Width           =   1635
      End
      Begin VB.OptionButton optPlaceCommentsWhere 
         Caption         =   "&Before Declaration"
         Height          =   195
         Index           =   0
         Left            =   2025
         TabIndex        =   2
         Top             =   30
         Value           =   -1  'True
         Width           =   1650
      End
      Begin VB.CommandButton cmdSaveCommentTemplateToFile 
         Caption         =   "&Save As Default"
         Height          =   330
         Left            =   1335
         TabIndex        =   6
         Top             =   1980
         Width           =   1410
      End
      Begin VB.CommandButton cmdLoadCommentTemplateFromFile 
         Caption         =   "&Load From File"
         Height          =   330
         Left            =   90
         TabIndex        =   5
         Top             =   1980
         Width           =   1200
      End
      Begin VB.ComboBox cboLang 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         ItemData        =   "frmCodeCommenter.frx":0865
         Left            =   3570
         List            =   "frmCodeCommenter.frx":0875
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2010
         Width           =   1830
      End
      Begin VB.TextBox txtComments 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1650
         Left            =   105
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   4
         Text            =   "frmCodeCommenter.frx":089E
         Top             =   270
         Width           =   5280
      End
   End
End
Attribute VB_Name = "frmCodeCommenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

 
'*************************************************************************
' Author    : Charlie Kirkwood
' File      : frmCodeCommenter.frm
' NOTE:     : This program is based upon an seriously scaled-down
'               program by a developer named Jean-Philippe Leconte.
'               He created his program on 11 october 2000 03:48.
'
'*************************************************************************
' History   : 20001102 - added file system object to load info from text files
'                      - added a LOT of functionality and seriously modified the interface
'                      - allowed user to create stored error handling types
'                      - allowed user to select what files to apply the error trapping/commenting to
'                      - added progress meter
'                      - aw F-it... it's just a different program!!!  *hehe*
'             20001210 - added functionality for moving deinstantiating of objects
'                           to the error handler if they are deinstantiated at the end
'                           of hte procedure (note: any deinstantiation calls that are
'                           not at the end of the procedure, will not be moved into the
'                           error handler
'                      - modified the program to use an array and write to a new file
'                           instead of the existing file.  This was necessary to handle
'                           moving hte deinstantiation code.
'
'*************************************************************************

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Const BACKUP_EXT = ".backup"
Private Const PROCMODULENAME = "%m"
Private Const PROCSCOPE = "%s"
Private Const PROCTYPE = "%t"
Private Const PROCNAME = "%n"
Private Const PROCPARAM = "%p"
Private Const PROCRETURN = "%r"
Private Const PROCDESC = "%d"
Private Const CurrentDateTime = "%x"
Private Const PROCDEINSTANTIATION = "%c"


Private Const LANG_FRA = "Français"
Private Const LANG_ENG = "English"
Private Const LANG_ESP = "Español"
Private Const LANG_DEU = "Deutsch"


Private vArrFileTypes As Variant
Private vArrScopes As Variant
Private vArrMids As Variant
Private vArrProcedures As Variant
Private vArrProcedureEnds As Variant
Private vArrEnds As Variant
Private vArrOnError As Variant

Private lModProcedures As Long
Private lProcedures As Long

Private Enum eTemplate
    eTemplateComment
    eTemplateDescription
    eTemplateErrorTrapping
End Enum

Private Enum eAction
    eActionApplyToFolder = 0
    eActionApplyToFile = 1
End Enum

Private meAction As Long
Private Const mcsActionCaption As String = "Apply the selected items to the following"
Private Const mcsFile As String = " File:"
Private Const mcsFolder As String = " Folder:"
Private Const mcsNoFileSelected As String = "No file was selected, do you want to use the default template?"

Private Const mcsDefaultTemplateDescription As String = "_DescriptionTemplate.lht"
Private Const mcsDefaultTemplateComment As String = "_CommentTemplate.lht"
Private Const mcsDefaultTemplateErrorHandling As String = "_ErrorHandlingTemplate.lht"

Private Const mcsFilterForAllVBFiles As String = "All files (*.*)|*.*|Forms (*.frm)|*.frm|Modules (*.bas)|*.bas|Classes (*.cls)|*.cls|User controls (*.ctl)|*.ctl"
Private Const mcsFilterForTemplateFiles As String = "Little Helper Template Files (*.lht)|*.lht"

Private Const mcsProcTypePropLet As String = "Property Let"
Private Const mcsProcTypePropGet As String = "Property Get"
Private Const mcsProcTypePropSet As String = "Property Set"
Private Const mcsProperty As String = "Property"

Private Const mcsLineContinuationChar As String = "_"

Private Const mcsModuleNameIdentifier As String = "Attribute VB_Name = """

Private Enum eCommonDialogConst
    ecdlShowOpen = 1
    ecdlShowSave = 2
    ecdlShowColor = 3
    ecdlShowFont = 4
    ecdlShowPrinter = 5
    ecdlShowWinHelp32 = 6
End Enum

Dim moHourglass As clsHourglass
Dim fCancelProcessing As Boolean

Private Const mclOffScreen As Long = -20000
Private Const mcsExitStatement As String = "Exit "

Private Const mcsDeinstantiateStatement As String = "Set * = Nothing"
Private Const mclCommentsBeforeDeclaration As Long = 0
Private Const mclCommentsAfterDeclaration As Long = 1

Private Const mcsDefaultTab As String = "    "


'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : chkComments_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for chkComments_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Sub chkComments_Click()
    On Error GoTo Proc_Err
    Const csProcName As String = "chkComments_Click"


    Me.fraComments.Enabled = Me.chkComments


Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
        

    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmCodeCommenter->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Sub


'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : chkErrorHandling_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for chkErrorHandling_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Sub chkErrorHandling_Click()
    On Error GoTo Proc_Err
    Const csProcName As String = "chkErrorHandling_Click"


    Dim ctl As Control
    
    Me.fraErrorHandling.Enabled = Me.chkErrorHandling



Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
        

    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmCodeCommenter->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Sub


'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : cmdCancel_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for cmdCancel_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Sub cmdCancel_Click()
    On Error GoTo Proc_Err
    Const csProcName As String = "cmdCancel_Click"


    fCancelProcessing = True


Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
    
    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmCodeCommenter->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Sub


'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : cmdGo_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for cmdGo_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Sub cmdGo_Click()
    On Error GoTo Proc_Err
    Const csProcName As String = "cmdGo_Click"


    Dim lCount As Long
    Dim sFolder As String
    Dim sFileName  As String
    Dim sFiles As String
    Dim sFileTypes As String
    
    Dim oFileOps As clsFileOps
    Dim lFilesToProcess As Long
    Dim lFilesProcessed As Long
    Dim vFileArray() As String
    

    If (Me.chkErrorHandling <> vbChecked And Me.chkComments <> vbChecked) _
        Or Me.txtApplyTo & "" = "" Then
        MsgBox "You have not asked me to do anything.  Either check off Comments, Error Trapping or both, and select a file or folder to apply the code to.", vbExclamation + vbOKOnly
    Else
        
        Set moHourglass = New clsHourglass
        
        Select Case meAction
            Case eAction.eActionApplyToFile
                'ModifyFile Me.txtApplyTo, CBool(chkComments.Value), CBool(chkErrorHandling.Value), CBool(chkBackup.Value)
                ModifyFileUsingFileArray Me.txtApplyTo, CBool(chkComments.Value), CBool(chkErrorHandling.Value), CBool(chkBackup.Value)
                
            Case eAction.eActionApplyToFolder
                
                'build the list of file types to compare against
                For lCount = 0 To Me.lstFileTypes.ListCount - 1
                    If Me.lstFileTypes.Selected(lCount) Then
                        sFileTypes = sFileTypes & Me.lstFileTypes.List(lCount)
                    End If
                Next
                
                
                sFolder = Me.txtApplyTo
                If Not Right(sFolder, 1) = "\" Then
                    sFolder = sFolder + "\"
                End If
                
                Set oFileOps = New clsFileOps
                lFilesToProcess = oFileOps.FilesToArray(sFolder, False, False, vFileArray)
                Set oFileOps = Nothing
                
                Me.prgProgress.Min = LBound(vFileArray)
                Me.prgProgress.Max = UBound(vFileArray)
                
                'swap the frames to show progress bar
                Call Me.fraProgress.Move(Me.fraAction.Left, Me.fraAction.Top)
                Call Me.fraAction.Move(mclOffScreen, mclOffScreen)
                
                DoEvents
                Me.Refresh
                
                lCount = 0
                fCancelProcessing = False
                
                For lCount = LBound(vFileArray) To UBound(vFileArray)
                    
                    
                    sFileName = vFileArray(lCount)
                    
                    If InStr(1, sFileTypes, Right(sFileName, 4)) <> 0 Then
                        sFiles = sFiles + ModifyFileUsingFileArray(sFolder + sFileName, CBool(chkComments.Value), CBool(chkErrorHandling.Value), CBool(chkBackup.Value), True) + vbCrLf
                    End If
                    
                    'update the progress
                    Me.prgProgress.Value = lCount
                    
                    'allow program to halt to processor events - lets user click cancel
                    DoEvents
                    
                    'break processing appropriately
                    If fCancelProcessing Then
                        lCount = UBound(vFileArray) + 1
                    End If
                    
                    
                Next
                
                MsgBox "finished added selected information to VB code in folder: " + Me.txtApplyTo, vbInformation + vbInformation
                
                'swap the frames back
                Call Me.fraAction.Move(Me.fraProgress.Left, Me.fraProgress.Top)
                Call Me.fraProgress.Move(mclOffScreen, mclOffScreen)
                
            
            Case Else
                MsgBox "this action is not supported", vbExclamation + vbOKOnly
        End Select
    End If
    
    
    On Error Resume Next







Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
        Set moHourglass = Nothing
    Set oFileOps = Nothing
    
        
    

    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmCodeCommenter->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Sub



'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : cmdApplyToFolder_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for cmdApplyToFolder_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Sub cmdApplyToFolder_Click()
    On Error GoTo ErrorHandle_cmdApplyToFolder_Click
    
    Dim lCount
    
'    Dim sFolder As String
'    Dim sFilename As String
'    Dim sMsgBox As String
'    Dim sFiles As String
    
    'MsgBox "As a Safety precaution, this option will always make backups of files modified", vbInformation, "Commentor"
    Me.chkBackup.Value = vbChecked
    
    Dim oFileOps As clsFileOps
    Set oFileOps = New clsFileOps
    
    Me.txtApplyTo = oFileOps.BrowseForFolderPf(Me.hWnd, "Select Folder For Commenting and Error Trapping")
    
    If Me.txtApplyTo & "" <> "" Then
        meAction = eAction.eActionApplyToFolder
        Me.lblActionCaption = mcsActionCaption + mcsFolder
        Me.lstFileTypes.Enabled = True
        
        Me.lstFileTypes.Clear
        
        If GetAllFileExtensions(Me.txtApplyTo, vArrFileTypes) > 0 Then
        
            For lCount = LBound(vArrFileTypes) To UBound(vArrFileTypes)
                Me.lstFileTypes.AddItem vArrFileTypes(lCount)
            Next
            
        End If
        
        Me.lstFileTypes.SetFocus
        
    End If
    
    
    Set oFileOps = Nothing
    
        
    

ErrorHandle_cmdApplyToFolder_Click:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Sub"
    sErrorName = "cmdApplyToFolder_Click"
    sErrorReturns = "Nothing"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical, App.Title
          Err.Clear
          Resume Next
          Resume

    End Select
End Sub

'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : cmdApplyToFile_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for cmdApplyToFile_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Sub cmdApplyToFile_Click()
    On Error GoTo ErrorHandle_cmdApplyToFile_Click
    
    
    meAction = eAction.eActionApplyToFile
    Me.lblActionCaption = mcsActionCaption + mcsFile
    
    
    cdlgFile.DialogTitle = "Visual Basic file to load"
    cdlgFile.DefaultExt = ".*"
    cdlgFile.Filter = mcsFilterForAllVBFiles
    cdlgFile.CancelError = True
    cdlgFile.Flags = cdlOFNHideReadOnly
    cdlgFile.ShowOpen
    If Len(cdlgFile.FileName) > 0 And Len(Dir(cdlgFile.FileName)) > 0 Then
    
        Me.txtApplyTo = cdlgFile.FileName
        'ModifyFile  cdlgFile.FileName, CBool(chkComments.Value), CBool(chkErrorHandling.Value), CBool(chkBackup.Value)
    Else
        MsgBox "Cannot find file", vbCritical, "Commentor"
    End If


    Me.lstFileTypes.Enabled = False
    

ErrorHandle_cmdApplyToFile_Click:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Sub"
    sErrorName = "cmdApplyToFile_Click"
    sErrorReturns = "Nothing"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case 32755
          'Cancel was pressed
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical Or vbApplicationModal, App.Title
          Err.Clear
          Resume Next
    End Select
End Sub


'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : cmdLoadCommentTemplateFromFile_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for cmdLoadCommentTemplateFromFile_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Sub cmdLoadCommentTemplateFromFile_Click()
    On Error GoTo Proc_Err
    Const csProcName As String = "cmdLoadCommentTemplateFromFile_Click"


    Dim sTemplate As String
    sTemplate = GetTemplateFileName(ecdlShowOpen, App.Title & mcsDefaultTemplateComment)
        
    'if no file selected, use default
    If sTemplate & "" = "" Then
        If MsgBox(mcsNoFileSelected, vbQuestion + vbYesNoCancel) = vbYes Then
            sTemplate = App.Path + "\" + App.Title + mcsDefaultTemplateComment
        End If
    End If
    
    Call LoadTemplateFromFile(sTemplate, Me.txtComments)


Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
    
    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmCodeCommenter->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Sub




'__________________________________________________
' Scope  : Private
' Type   : Function
' Name   : GetTemplateFileName
' Params :
'          Optional eMode As eCommonDialogConst = ecdlShowOpen
'          Optional sDefaultTemplate As String = ""
' Returns: String
' Desc   : The Function uses parameters Optional eMode As eCommonDialogConst = ecdlShowOpen and Optional sDefaultTemplate As String = "" for GetTemplateFileName and returns String.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Function GetTemplateFileName(Optional eMode As eCommonDialogConst = ecdlShowOpen, Optional sDefaultTemplate As String = "") As String

    On Error GoTo Proc_Exit
    Dim sFile As String
    Dim fFileExists As Boolean
    Dim lResponse As Long
    
    cdlgFile.DialogTitle = "Code Template"
    cdlgFile.InitDir = App.Path
    cdlgFile.DefaultExt = ".lht"
    cdlgFile.FileName = sDefaultTemplate
    cdlgFile.Filter = mcsFilterForTemplateFiles
    cdlgFile.CancelError = True
    cdlgFile.Flags = cdlOFNHideReadOnly
    cdlgFile.Action = eMode
    
    
    
    Select Case eMode
        Case eCommonDialogConst.ecdlShowSave
            If Len(cdlgFile.FileName) = 0 Then
                'no file typed in
                Err.Raise Number:=4001, Description:="no file selected"
            
            Else
                sFile = cdlgFile.FileName
                fFileExists = CBool(Len(Dir(sFile)))
                If fFileExists Then
                    lResponse = MsgBox("File already exists, overwrite existing file?", vbQuestion + vbYesNoCancel, App.Title)
                    If lResponse <> vbYes Then
                        sFile = ""
                    End If
                End If
            End If
        Case eCommonDialogConst.ecdlShowOpen
            If Len(cdlgFile.FileName) > 0 And Len(Dir(cdlgFile.FileName)) > 0 Then
                sFile = cdlgFile.FileName
            Else
                Err.Raise Number:=4001, Description:="Cannot find file"
            End If
            
    End Select
    
Proc_Exit:

    GetTemplateFileName = sFile

End Function



'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : cmdLoadDescriptionTemplateFromFile_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for cmdLoadDescriptionTemplateFromFile_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Sub cmdLoadDescriptionTemplateFromFile_Click()
    On Error GoTo Proc_Err
    Const csProcName As String = "cmdLoadDescriptionTemplateFromFile_Click"


    Dim sTemplate As String
    sTemplate = GetTemplateFileName(ecdlShowOpen, App.Title & mcsDefaultTemplateDescription)
    
        
    'if no file selected, use default
    If sTemplate & "" = "" Then
        If MsgBox(mcsNoFileSelected, vbQuestion + vbYesNoCancel) = vbYes Then
            sTemplate = App.Path + "\" + App.Title + mcsDefaultTemplateDescription
        End If
    End If
    
    Call LoadTemplateFromFile(sTemplate, Me.txtDescription)


Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
    
    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmCodeCommenter->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Sub


'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : cmdLoadErrorTemplateFromFile_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for cmdLoadErrorTemplateFromFile_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Sub cmdLoadErrorTemplateFromFile_Click()
    On Error GoTo Proc_Err
    Const csProcName As String = "cmdLoadErrorTemplateFromFile_Click"


    Dim sTemplate As String
    sTemplate = GetTemplateFileName(ecdlShowOpen, App.Title & mcsDefaultTemplateErrorHandling)
        
    'if no file selected, use default
    If sTemplate & "" = "" Then
        If MsgBox(mcsNoFileSelected, vbQuestion + vbYesNoCancel) = vbYes Then
            sTemplate = App.Path + "\" + App.Title + mcsDefaultTemplateErrorHandling
        End If
    End If
    
    Call LoadTemplateFromFile(sTemplate, Me.txtErrorHandling)


Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
    
    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmCodeCommenter->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Sub

'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : cmdQuit_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for cmdQuit_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Sub cmdQuit_Click()
    On Error GoTo ErrorHandle_cmdQuit_Click
    Unload Me

ErrorHandle_cmdQuit_Click:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Sub"
    sErrorName = "cmdQuit_Click"
    sErrorReturns = "Nothing"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical, App.Title
          Err.Clear
          Resume Next
    End Select
End Sub




'__________________________________________________
' Scope  : Private
' Type   : Function
' Name   : ModifyFileUsingFileArray
' Params :
'          ByVal sFileName As String
'          ByVal bComments As Boolean
'          ByVal bErrorHandling As Boolean
'          ByVal bMakeBackup As Boolean
'          Optional bDoNotDisplay As Boolean = False
' Returns: String
' Desc   : The Function uses parameters ByVal sFileName As String, ByVal bComments As Boolean, ByVal bErrorHandling As Boolean, ByVal bMakeBackup As Boolean and Optional bDoNotDisplay As Boolean = False for ModifyFileUsingFileArray and returns String.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Function ModifyFileUsingFileArray(ByVal sFileName As String, ByVal bComments As Boolean, ByVal bErrorHandling As Boolean, ByVal bMakeBackup As Boolean, Optional bDoNotDisplay As Boolean = False) As String
    

    On Error GoTo ErrorHandle_ModifyFile
    
    Dim ofs As clsFs
    Dim tsOldFile As TextStream
    Dim tsNewFile As TextStream
    
    
    Dim sLastScope As String

    Dim sModuleName As String
    Dim sLastMid As String
    Dim sLastType As String
    Dim sLastName As String
    Dim sLastReturn As String
    Dim sScope As String
    Dim sMid As String
    Dim sType As String
    Dim sName As String
    Dim vArrParameters As Variant
    Dim sReturn As String
    Dim sDescription As String
    Dim sEnd As String
    Dim sSpace As String
    
    Dim bStartErrorHandling As Boolean
    Dim sFile As String
    Dim vArrFileContents As Variant
    Dim iOpen As Integer
    Dim lCount As Long
    Dim lLbound As Long
    Dim lUbound As Long
    Dim lChar As Long
    Dim lPos As Long
    

    Dim sMsgBox As String
    Dim bUp As Long
    Dim sTemplate As String
    Dim lTabLen As Long


    vArrScopes = Array("Private", "Public", "Global", "Friend", "Protected")
    vArrMids = Array("Static")
    vArrProcedures = Array("Function", "Sub", mcsProcTypePropLet, mcsProcTypePropGet, mcsProcTypePropSet)
    vArrEnds = Array("End")
    vArrOnError = Array("On Error")
    vArrProcedureEnds = Array("Function", "Sub", "Property")

'    Set ofs = New clsFs
'
'    If bMakeBackup Then
'        Call ofs.CopyFile(sFileName, sFileName + BACKUP_EXT, True)
'    End If
'
'
'    'open the existing file, dump it into an array for modification
'    Set tsOldFile = ofs.OpenTextFile(sFileName, ForReading)
'    sFile = tsOldFile.ReadAll
'    Set tsOldFile = Nothing
'
    
    If bMakeBackup Then FileCopy sFileName, sFileName + BACKUP_EXT

    iOpen = FreeFile(1)
    Open sFileName For Input As iOpen
        sFile = Input(LOF(iOpen), iOpen)
    Close iOpen

    vArrFileContents = Split(sFile, vbCrLf)
    lChar = 1

    
    
    'then roll through the array placing code where necessary.
    
    
    'get the mod name
    sModuleName = GetModuleName(vArrFileContents)
    
    'get the ubound each time since it'll change due to inserting stuff into the array
    'lLbound = LBound(vArrFileContents)

    'For lCount = lLbound To UBound(vArrFileContents)
    
    lCount = LBound(vArrFileContents)
    Do While lCount <= UBound(vArrFileContents)
        lPos = 1
        sScope = ""
        sMid = ""
        sType = ""
        sName = ""
        sReturn = ""
        vArrParameters = Null
        sDescription = ""
        sEnd = ""


        bUp = False

        sScope = CheckScope(vArrFileContents(lCount), lPos)
        lPos = lPos + IIf(Len(sScope) = 0, 0, Len(sScope) + 1)
        sMid = CheckMid(vArrFileContents(lCount), lPos)
        lPos = lPos + IIf(Len(sMid) = 0, 0, Len(sMid) + 1)

        If (Len(sScope) > 0 And Len(sMid) > 0) Then
            sSpace = " "
        Else
            sSpace = ""
        End If

        'check to see if the array element is a function/sub/property let/get etc
        sType = CheckProcedure(vArrFileContents(lCount), lPos)

        If Len(sType) > 0 Then
            lProcedures = lProcedures + 1



            lPos = lPos + Len(sType) + 1
            sName = GetName(vArrFileContents(lCount), lPos)

            'if this is a new code block, start the commenter / error handler insertion
            If Len(sName) > 0 Then
            
                lPos = lPos + Len(sName) + 1
                sReturn = GetReturn(vArrFileContents, lCount)
                If bComments Then
                    'vArrParameters = GetParams(vArrFileContents (lCount), lPos)
                    vArrParameters = GetParams(vArrFileContents, lCount)
                    
                    
                    'todo: must handle when parameters are on multiple lines!
                    
                    sDescription = MakeDescription(txtDescription.Text, sScope + sSpace + sMid, PROCSCOPE, _
                                                    sType, PROCTYPE, sName, PROCNAME, _
                                                    vArrParameters, PROCPARAM, _
                                                    sReturn, PROCRETURN, _
                                                    sModuleName, PROCMODULENAME, CurrentDateTime)
                    
                    
                    If Me.optPlaceCommentsWhere(mclCommentsAfterDeclaration).Value = True Then
                        
                        
                        'handle situation where there is a line continuation character
                        '   this places us on the last line of the declaration
                        Do While Right(vArrFileContents(lCount), 1) = mcsLineContinuationChar
                            lChar = lChar + Len(vArrFileContents(lCount)) + Len(vbCrLf)
                            lCount = lCount + 1
                        Loop
                        
                        'move to the next line PAST the declaration.
                        lCount = lCount + 1

                        'make sure the comments are tabbed the same as the tab of the next lind of code
                        lTabLen = getTabLenFromNextLineMl(vArrFileContents, lCount)

                        sTemplate = Replace(Me.txtComments.Text, "'", String(lTabLen, " ") & "'")
                    
                    Else
                    
                        sTemplate = Me.txtComments
                    
                    End If
                    
                    
                    lChar = lChar + AddCommentsToFileArray(vArrFileContents, lCount, sTemplate, sScope + sSpace + sMid, PROCSCOPE, _
                                                    sType, PROCTYPE, sName, PROCNAME, vArrParameters, _
                                                    PROCPARAM, sReturn, PROCRETURN, sDescription, PROCDESC, _
                                                    sModuleName, PROCMODULENAME, CurrentDateTime)
                    
                    
                    lModProcedures = lModProcedures + 1
                    bUp = True
                End If
                
                If bErrorHandling Then

                    'check here to see if there is already error handling in the code
                    '   look ahead until 'end' is encountered if we find an on error
                    '   then skip this code block
                    If Len(LookAhead(vArrFileContents, lCount, vArrOnError, vArrEnds)) > 0 Then
                        bStartErrorHandling = False
                    Else
                        bStartErrorHandling = True
                        sLastScope = sScope
                        sLastMid = sMid
                        sLastType = sType
                        sLastName = sName
                        sLastReturn = sReturn
                                                
                        If Me.optPlaceCommentsWhere(mclCommentsBeforeDeclaration).Value = True Then
                                                                        
                            'we are on the proc declaration, is it continued to the next line? if so,
                            '   account for the lenght of that line, then move to the next
                            Do While Right(vArrFileContents(lCount), 1) = mcsLineContinuationChar
                                lChar = lChar + Len(vArrFileContents(lCount)) + Len(vbCrLf)
                                lCount = lCount + 1
                            Loop
                        
                            'since we're on the procedure declaration line OR comments line if user
                            '   chose to place comments after declaration , move one more down
                            lCount = lCount + 1
                            
                        End If
                        
                        lChar = lChar + AddErrorHandlingTopToFileArray(vArrFileContents, lCount, _
                                                            txtErrorHandlingTop.Text, _
                                                             sScope + sSpace + sMid, _
                                                            PROCSCOPE, sType, PROCTYPE, sName, PROCNAME, sReturn, _
                                                            PROCRETURN, sModuleName, PROCMODULENAME, CurrentDateTime)
                                                            
                                                            
                        If Not bUp Then
                            lModProcedures = lModProcedures + 1
                        End If
                    End If
                End If
            End If
        End If
        If bErrorHandling And bStartErrorHandling Then
            lPos = 1
            sEnd = CheckEnd(vArrFileContents(lCount))
            lPos = lPos + IIf(Len(sEnd) = 0, 0, Len(sEnd) + 1)
            If Len(sEnd) > 0 And Len(CheckProcedureEnd(vArrFileContents(lCount), lPos)) > 0 Then
                lChar = lChar + AddErrorHandlingToFileArray(vArrFileContents, lCount, txtErrorHandling.Text, _
                                                sLastScope + IIf(Len(sLastScope) > 0 And Len(sLastMid) > 0, " ", "") + sLastMid, _
                                                PROCSCOPE, sLastType, PROCTYPE, sLastName, PROCNAME, sLastReturn, _
                                                PROCRETURN, sModuleName, PROCMODULENAME, CurrentDateTime, PROCDEINSTANTIATION)
                bStartErrorHandling = False
                sLastScope = ""
                sLastMid = ""
                sLastType = ""
                sLastName = ""
                sLastReturn = ""
            End If
        End If

        lChar = lChar + Len(vArrFileContents(lCount)) + Len(vbCrLf)

        'increment lcount - originally used a for-next, but the ubound of the array changed.  I tried to
        '   loop from lbound to ubound(array)  but even though the ubound of the array changed, the loop's end
        '   never did. so try "do while"
        lCount = lCount + 1
    Loop



'    'write the new file from the array
'    Set tsNewFile = ofs.CreateTextFile(sFileName & "_New", True, False)
    
    lLbound = LBound(vArrFileContents)
    lUbound = UBound(vArrFileContents)
    sFile = ""
    For lCount = lLbound To lUbound
        sFile = sFile & vArrFileContents(lCount) & vbCrLf
    Next
        
    'remove the last vbcrlf
    sFile = Left(sFile, Len(sFile) - Len(vbCrLf))
        
    iOpen = FreeFile(1)
    Open sFileName For Output As iOpen
        Print #iOpen, sFile
        
    Close iOpen



    If Not bDoNotDisplay Then
        sMsgBox = IIf(bComments, "comments", "")
        sMsgBox = sMsgBox + IIf(Len(sMsgBox) > 0 And bErrorHandling, " and ", "") + IIf(bErrorHandling, "error handling", "")
        MsgBox "Added " + IIf(Len(sMsgBox) > 0, sMsgBox, "nothing") + " to " + sFileName + vbCrLf + vbCrLf + "Modified " + CStr(lModProcedures) + " procedures of " + CStr(lProcedures), vbApplicationModal, "Commentor"
        lProcedures = 0
    End If

    'ModifyFile = sFileName

ErrorHandle_ModifyFile:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Function"
    sErrorName = "ModifyFile"
    sErrorReturns = "String"

    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical, App.Title
          Err.Clear
          Resume Next
          Resume

    End Select
End Function


'
'__________________________________________________
' Scope  : Private
' Type   : Function
' Name   : CheckScope
' Params :
'          ByVal sLine As String
'          ByVal lPos As Long
' Returns: String
' Desc   : The Function uses parameters ByVal sLine As String and ByVal lPos As Long for CheckScope and returns String.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Function CheckScope(ByVal sLine As String, ByVal lPos As Long) As String
    On Error GoTo ErrorHandle_CheckScope
    Dim sFound As String
    Dim lCount As Long
    
    While Not lCount > UBound(vArrScopes)
        If InStr(lPos, sLine, vArrScopes(lCount)) = lPos Then sFound = vArrScopes(lCount)
        lCount = lCount + 1
    Wend
    
    CheckScope = sFound

ErrorHandle_CheckScope:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Function"
    sErrorName = "CheckScope"
    sErrorReturns = "String"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical, App.Title
          Err.Clear
          Resume Next
    End Select
End Function

'__________________________________________________
' Scope  : Private
' Type   : Function
' Name   : CheckMid
' Params :
'          ByVal sLine As String
'          ByVal lPos As Long
' Returns: String
' Desc   : The Function uses parameters ByVal sLine As String and ByVal lPos As Long for CheckMid and returns String.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Function CheckMid(ByVal sLine As String, ByVal lPos As Long) As String
    On Error GoTo ErrorHandle_CheckMid
    Dim sFound As String
    Dim lCount As Long
    
    While Not lCount > UBound(vArrMids)
        If InStr(lPos, sLine, vArrMids(lCount)) = lPos Then sFound = vArrMids(lCount)
        lCount = lCount + 1
    Wend
    
    CheckMid = sFound

ErrorHandle_CheckMid:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Function"
    sErrorName = "CheckMid"
    sErrorReturns = "String"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical, App.Title
          Err.Clear
          Resume Next
    End Select
End Function

'__________________________________________________
' Scope  : Private
' Type   : Function
' Name   : CheckProcedure
' Params :
'          ByVal sLine As String
'          ByVal lPos As Long
' Returns: String
' Desc   : The Function uses parameters ByVal sLine As String and ByVal lPos As Long for CheckProcedure and returns String.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Function CheckProcedure(ByVal sLine As String, ByVal lPos As Long) As String
    On Error GoTo ErrorHandle_CheckProcedure
    Dim sFound As String
    Dim lCount As Long
    
    While Not lCount > UBound(vArrProcedures)
        If InStr(lPos, sLine, vArrProcedures(lCount)) = lPos Then sFound = vArrProcedures(lCount)
        lCount = lCount + 1
    Wend
    
    CheckProcedure = sFound

ErrorHandle_CheckProcedure:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Function"
    sErrorName = "CheckProcedure"
    sErrorReturns = "String"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical, App.Title
          Err.Clear
          Resume Next
    End Select
End Function

'__________________________________________________
' Scope  : Private
' Type   : Function
' Name   : LookAhead
' Params :
'          ByVal varr As Variant
'          lCurrentElement As Long
'          vLookFor As Variant
'          vLookUntil As Variant
' Returns: String
' Desc   : The Function uses parameters ByVal varr As Variant, lCurrentElement As Long, vLookFor As Variant and vLookUntil As Variant for LookAhead and returns String.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Function LookAhead(ByVal varr As Variant, lCurrentElement As Long, vLookFor As Variant, vLookUntil As Variant) As String
    On Error GoTo ErrorHandle_LookAhead
    Dim vCloneArr As Variant
    Dim sFound As String
    Dim lCountLookFor As Long
    Dim lCountLookUntil As Long
    Dim lCountCloneArrElement As Long
    Dim sTrimmedCloneArrVal As String
    
    'starting at lcurrentline, look ahead for 'vLookFor' until 'vLookUntil' if it exists, return the string
    vCloneArr = varr
    
    
    lCountCloneArrElement = lCurrentElement
    'look for each item in vLookFor in every line of vCloneArr until we hit vLookUntil or the end of hte array
    While Not lCountCloneArrElement > UBound(vCloneArr)
        
        lCountLookUntil = 0
        
        sTrimmedCloneArrVal = Trim(vCloneArr(lCountCloneArrElement))
        
        'check each line for vLookUntil if we hit it jump out, else check for vLookFor
        While Not lCountLookUntil > UBound(vLookUntil)
            If Left(sTrimmedCloneArrVal, Len(vLookUntil(lCountLookUntil))) = vLookUntil(lCountLookUntil) Then
                
                'if we found lookuntil in the array, jump out of the search
                lCountCloneArrElement = UBound(vCloneArr) + 1
                
            Else
    
                lCountLookFor = 0
                While Not lCountLookFor > UBound(vLookFor)
                                        
                    If Left(sTrimmedCloneArrVal, Len(vLookFor(lCountLookFor))) = vLookFor(lCountLookFor) Then
                        sFound = vLookFor(lCountLookFor)
                        lCountLookFor = UBound(vLookFor) + 1
                        lCountLookUntil = UBound(vLookUntil) + 1
                        lCountCloneArrElement = UBound(vCloneArr) + 1
                    End If
                    
                    lCountLookFor = lCountLookFor + 1
                    
                Wend
            End If
            
            lCountLookUntil = lCountLookUntil + 1
            
        Wend
        
        lCountCloneArrElement = lCountCloneArrElement + 1

    Wend
    
    LookAhead = sFound

ErrorHandle_LookAhead:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Function"
    sErrorName = "LookAhead"
    sErrorReturns = "String"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical, App.Title
          Err.Clear
          Resume Next
          Resume
    End Select
End Function



'__________________________________________________
' Scope  : Private
' Type   : Function
' Name   : CheckEnd
' Params :
'          ByVal sLine As String
' Returns: String
' Desc   : The Function uses parameters ByVal sLine As String for CheckEnd and returns String.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Function CheckEnd(ByVal sLine As String) As String
    On Error GoTo ErrorHandle_CheckEnd
    Dim sFound As String
    Dim lCount As Long
    
    While Not lCount > UBound(vArrEnds)
        If Left(Trim(sLine), Len(vArrEnds(lCount))) = vArrEnds(lCount) Then sFound = vArrEnds(lCount)
        lCount = lCount + 1
    Wend
    
    CheckEnd = sFound

ErrorHandle_CheckEnd:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Function"
    sErrorName = "CheckEnd"
    sErrorReturns = "String"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical, App.Title
          Err.Clear
          Resume Next
    End Select
End Function

'__________________________________________________
' Scope  : Private
' Type   : Function
' Name   : CheckOnError
' Params :
'          ByVal sLine As String
' Returns: String
' Desc   : The Function uses parameters ByVal sLine As String for CheckOnError and returns String.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Function CheckOnError(ByVal sLine As String) As String
    On Error GoTo ErrorHandle_CheckOnError
    Dim sFound As String
    Dim lCount As Long
    
    While Not lCount > UBound(vArrOnError)
        If Left(Trim(sLine), Len(vArrOnError(lCount))) = vArrOnError(lCount) Then sFound = vArrOnError(lCount)
        lCount = lCount + 1
    Wend
    
    CheckOnError = sFound

ErrorHandle_CheckOnError:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Function"
    sErrorName = "CheckOnError"
    sErrorReturns = "String"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical, App.Title
          Err.Clear
          Resume Next
    End Select
End Function

'__________________________________________________
' Scope  : Private
' Type   : Function
' Name   : CheckProcedureEnd
' Params :
'          ByVal sLine As String
'          ByVal lPos As Long
' Returns: String
' Desc   : The Function uses parameters ByVal sLine As String and ByVal lPos As Long for CheckProcedureEnd and returns String.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Function CheckProcedureEnd(ByVal sLine As String, ByVal lPos As Long) As String
    On Error GoTo ErrorHandle_CheckProcedureEnd
    Dim sFound As String
    Dim lCount As Long
    
    While Not lCount > UBound(vArrProcedureEnds)
        If InStr(lPos, sLine, vArrProcedureEnds(lCount)) = lPos Then sFound = vArrProcedureEnds(lCount)
        lCount = lCount + 1
    Wend
    
    CheckProcedureEnd = sFound

ErrorHandle_CheckProcedureEnd:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Function"
    sErrorName = "CheckProcedureEnd"
    sErrorReturns = "String"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical, App.Title
          Err.Clear
          Resume Next
    End Select
End Function

'__________________________________________________
' Scope  : Private
' Type   : Function
' Name   : GetName
' Params :
'          ByVal sLine As String
'          ByVal lPos As Long
' Returns: String
' Desc   : The Function uses parameters ByVal sLine As String and ByVal lPos As Long for GetName and returns String.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Function GetName(ByVal sLine As String, ByVal lPos As Long) As String
    On Error GoTo ErrorHandle_GetName
    Dim sFound As String
    Dim lEnd As Long
    
    lEnd = InStr(lPos, sLine, "(")
    If lEnd > 0 Then sFound = Mid(sLine, lPos, lEnd - lPos)
    
    GetName = sFound

ErrorHandle_GetName:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Function"
    sErrorName = "GetName"
    sErrorReturns = "String"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical, App.Title
          Err.Clear
          Resume Next
    End Select
End Function

'__________________________________________________
' Scope  : Private
' Type   : Function
' Name   : GetParams
' Params :
'          ByVal sLine As String
'          ByVal lPos As Long
' Returns: Variant
' Desc   : The Function uses parameters ByVal sLine As String and ByVal lPos As Long for GetParams and returns Variant.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________

'Private Function GetParams(ByVal sLine As String, ByVal lPos As Long) As Variant
'    On Error GoTo ErrorHandle_GetParams
'    Dim sFound As String
'    Dim lEnd As Long
'
'    lEnd = InStrRev(sLine, ")")
'    If lEnd > lPos Then sFound = Mid(sLine, lPos, lEnd - lPos)
'
'    GetParams = Split(sFound, ", ")

Private Function GetParams(ByVal varr As Variant, ByVal lCount As Long) As Variant
    On Error GoTo ErrorHandle_GetParams
    
    Dim sLine As String
    Dim lEnd As Long
    Dim lBegin As Long
    Dim vClone As Variant
    Dim lLooper As Long
    
    vClone = varr
    
    sLine = vClone(lCount) & " "
    
    Do While InStr(1, vClone(lCount), ")") = 0
        lCount = lCount + 1
        sLine = sLine & vClone(lCount) & " "
        
    Loop
    
    'sLine is now the complete declaration in one line
    
    'take only the params part
    lBegin = InStr(1, sLine, "(")
    lEnd = InStrRev(sLine, ")")
    sLine = Mid(sLine, lBegin + 1, lEnd - lBegin - 1)
    
    'take out any line continuation chars
    sLine = Replace(sLine, " _ ", " ")
    
    'split up the params
    vClone = Split(sLine, ", ")
    lBegin = LBound(vClone)
    lEnd = UBound(vClone)
    
    'kill all extra spaces
    For lLooper = lBegin To lEnd
        vClone(lLooper) = Trim(vClone(lLooper))
    Next
     
    GetParams = vClone

ErrorHandle_GetParams:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Function"
    sErrorName = "GetParams"
    sErrorReturns = "Variant"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical, App.Title
          Err.Clear
          Resume Next
    End Select
End Function

'__________________________________________________
' Scope  : Private
' Type   : Function
' Name   : GetReturn
' Params :
'          ByVal sLine As String
' Returns: String
' Desc   : The Function uses parameters ByVal sLine As String for GetReturn and returns String.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
'Private Function GetReturn(ByVal sLine As String) As String
Private Function GetReturn(ByVal varr As Variant, ByVal lCount As Long) As String
    On Error GoTo ErrorHandle_GetReturn
    
    Dim vClone As Variant
    
    vClone = varr
    
    'check to see if the curentline has a return char...if so, go to end of the continued line
    If Right(vClone(lCount), 1) = "_" Then
        Do While Right(vClone(lCount), 1) = mcsLineContinuationChar
            lCount = lCount + 1
        Loop
    End If
    
    If Right(RTrim(vClone(lCount)), 1) = ")" Then
        'If Right(Trim(Mid(sLine, InStrRev(sLine, " ") + 1)), 1) = ")" Then
        GetReturn = "Nothing"
    Else
        GetReturn = Trim(Mid(vClone(lCount), InStrRev(vClone(lCount), " ") + 1))
    End If
    

ErrorHandle_GetReturn:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Function"
    sErrorName = "GetReturn"
    sErrorReturns = "String"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical, App.Title
          Err.Clear
          Resume Next
    End Select
End Function


'__________________________________________________
' Scope  : Private
' Type   : Function
' Name   : MakeDescription
' Params :
'            ByVal sTemplate As String
'            ByVal sScope As String
'            ByVal sScopeParam As String
'            ByVal sType As String
'            ByVal sTypeParam As String
'            ByVal sName As String
'            ByVal sNameParam As String
'            ByVal vArrParameters As Variant
'            ByVal sParametersParam As String
'            ByVal sReturn As String
'            ByVal sReturnParam As String
'            ByVal sModuleName As String
'            ByVal sModuleNameParam As String
'            ByVal sDateParam As String
' Returns: String
' Desc   : The Function uses parameters  for MakeDescription and returns _.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Function MakeDescription(ByVal sTemplate As String, _
                                ByVal sScope As String, _
                                ByVal sScopeParam As String, _
                                ByVal sType As String, _
                                ByVal sTypeParam As String, _
                                ByVal sName As String, _
                                ByVal sNameParam As String, _
                                ByVal vArrParameters As Variant, _
                                ByVal sParametersParam As String, _
                                ByVal sReturn As String, _
                                ByVal sReturnParam As String, _
                                ByVal sModuleName As String, _
                                ByVal sModuleNameParam As String, _
                                ByVal sDateParam As String) As String
                                
                                
    On Error GoTo ErrorHandle_MakeDescription
    Dim sResult As String
    Dim sParams As String
    Dim sAnd As String
    Dim lCount As Long
    
    sResult = sTemplate
    
    sResult = Replace(sResult, sScopeParam, sScope)
    sResult = Replace(sResult, sTypeParam, sType)
    sResult = Replace(sResult, sNameParam, sName)
    sResult = Replace(sResult, sModuleNameParam, sModuleName)
    sResult = Replace(sResult, sDateParam, Format(Now, "YYYYMMDD"))
    
    For lCount = LBound(vArrParameters) To UBound(vArrParameters) - 1
        sParams = sParams + vArrParameters(lCount) + ", "
    Next lCount
    sAnd = " and "
    If cboLang.Text = LANG_ENG Then sAnd = " and "
    If cboLang.Text = LANG_FRA Then sAnd = " et "
    If cboLang.Text = LANG_ESP Then sAnd = " e "
    If cboLang.Text = LANG_DEU Then sAnd = " und "
    If Not UBound(vArrParameters) = -1 Then If UBound(vArrParameters) > LBound(vArrParameters) Then sParams = Left(sParams, Len(sParams) - 2) + sAnd + vArrParameters(UBound(vArrParameters)) Else sParams = vArrParameters(LBound(vArrParameters))
    sResult = Replace(sResult, sParametersParam, sParams)
    sResult = Replace(sResult, sReturnParam, sReturn)
    
    MakeDescription = sResult

ErrorHandle_MakeDescription:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Function"
    sErrorName = "MakeDescription"
    sErrorReturns = "String"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical, App.Title
          Err.Clear
          Resume Next
    End Select
End Function


'*************************************************************************
' Scope : Private
' Type : Function
' Name : AddComments
' Parameters :
'         ByRef sFile As String
'         ByVal sTemplate As String
'         ByVal lPos As Long
'         ByVal sScope As String
'         ByVal sScopeParam As String
'         ByVal sType As String
'         ByVal sTypeParam As String
'         ByVal sName As String
'         ByVal sNameParam As String
'         ByVal vArrParameters As Variant
'         ByVal sParametersParam As String
'         ByVal sReturn As String
'         ByVal sReturnParam As String
'         ByVal sDescription As String
'         ByVal sDescriptionParam As String
' Returns : Long
' Description : The Function uses parameters ByRef sFile As String, ByVal sTemplate As String, ByVal lPos As Long, ByVal sScope As String, ByVal sScopeParam As String, ByVal sType As String, ByVal sTypeParam As String, ByVal sName As String, ByVal sNameParam As String, ByVal vArrParameters As Variant, ByVal sParametersParam As String, ByVal sReturn As String, ByVal sReturnParam As String, ByVal sDescription As String and ByVal sDescriptionParam As String for AddComments and returns Long.
'*************************************************************************
Private Function AddCommentsToFileArray(ByRef vFileArray As Variant, _
                            ByRef lCurrentElement As Long, _
                            ByVal sTemplate As String, _
                            ByVal sScope As String, _
                            ByVal sScopeParam As String, _
                            ByVal sType As String, _
                            ByVal sTypeParam As String, _
                            ByVal sName As String, _
                            ByVal sNameParam As String, _
                            ByVal vArrParameters As Variant, _
                            ByVal sParametersParam As String, _
                            ByVal sReturn As String, _
                            ByVal sReturnParam As String, _
                            ByVal sDescription As String, _
                            ByVal sDescriptionParam As String, _
                            ByVal sModuleName As String, _
                            ByVal sModuleNameParam As String, _
                            ByVal sDateParam As String) As Long

    


    On Error GoTo ErrorHandle_AddComments
    Dim lCount As Long
    Dim sResult As String
    Dim sParamLine As String
    Dim sParams As String
    Dim lParamPos As Long
    Dim lLastCrLf As Long
    Dim lNextCrLf As Long
    Dim sAnd As String
    
    sResult = sTemplate
    
    sResult = Replace(sResult, sScopeParam, sScope)
    sResult = Replace(sResult, sTypeParam, sType)
    sResult = Replace(sResult, sNameParam, sName)
    sResult = Replace(sResult, sModuleNameParam, sModuleName)
    sResult = Replace(sResult, sDateParam, Format(Now, "YYYYDDMM"))
    
    lParamPos = InStr(sResult, sParametersParam)
    
    If lParamPos > 0 Then
        lLastCrLf = InStrRev(sResult, vbCrLf, lParamPos)
        lNextCrLf = InStr(lParamPos, sResult, vbCrLf)
        If lNextCrLf - lLastCrLf > Len(sParametersParam) Then
            sParamLine = Mid(sResult, IIf(lLastCrLf > 0, lLastCrLf, 1), lNextCrLf - IIf(lLastCrLf > 0, lLastCrLf, 1))
            For lCount = LBound(vArrParameters) To UBound(vArrParameters)
                sParams = sParams + Replace(sParamLine, sParametersParam, vArrParameters(lCount))
            Next lCount
            sResult = Replace(sResult, sParamLine, sParams)
        Else
            For lCount = LBound(vArrParameters) To UBound(vArrParameters) - 1
                sParams = sParams + vArrParameters(lCount) + ", "
            Next lCount
            sAnd = " and "
            If cboLang.Text = LANG_ENG Then sAnd = " and "
            If cboLang.Text = LANG_FRA Then sAnd = " et "
            If cboLang.Text = LANG_ESP Then sAnd = " e "
            If cboLang.Text = LANG_DEU Then sAnd = " und "
            If Not UBound(vArrParameters) = -1 Then If UBound(vArrParameters) > LBound(vArrParameters) Then sParams = Left(sParams, Len(sParams) - 2) + sAnd + vArrParameters(UBound(vArrParameters)) Else sParams = vArrParameters(LBound(vArrParameters))
            sResult = Replace(sResult, sParametersParam, sParams)
        End If
    End If
    sResult = Replace(sResult, sReturnParam, sReturn)
    sResult = Replace(sResult, sDescriptionParam, sDescription)
    
    'bump the array down and add the comments to the current array
    '   element and move down one to the line we just worked on
    Call BumpArrayDown(vFileArray, lCurrentElement)
    vFileArray(lCurrentElement) = sResult
    lCurrentElement = lCurrentElement + 1
        
        
ErrorHandle_AddComments:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Function"
    sErrorName = "AddComments"
    sErrorReturns = "Long"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical, App.Title
          Err.Clear
          Resume Next
    End Select
End Function

Private Function AddErrorHandlingTopToFileArray(ByRef vFileArray As Variant, _
                                                ByRef lCurrentElement As Long, _
                                                ByVal sTemplate As String, _
                                                ByVal sScope As String, _
                                                ByVal sScopeParam As String, _
                                                ByVal sType As String, _
                                                ByVal sTypeParam As String, _
                                                ByVal sName As String, _
                                                ByVal sNameParam As String, _
                                                ByVal sReturn As String, _
                                                ByVal sReturnParam As String, _
                                                ByVal sModuleName As String, _
                                                ByVal sModuleNameParam As String, _
                                                ByVal sDateParam As String) As Long
    '*************************************************************************
    ' Scope : Private
    ' Type : Function
    ' Name : AddErrorHandlingTop
    ' Parameters :
    '         ByRef sFile As String
    '         ByVal sTemplate As String
    '         ByVal lPos As String
    '         ByVal sScope As String
    '         ByVal sScopeParam As String
    '         ByVal sType As String
    '         ByVal sTypeParam As String
    '         ByVal sName As String
    '         ByVal sNameParam As String
    '         ByVal sReturn As String
    '         ByVal sReturnParam As String
    ' Returns : Long
    ' Description : The Function uses parameters ByRef sFile As String, ByVal sTemplate As String, ByVal lPos As String, ByVal sScope As String, ByVal sScopeParam As String, ByVal sType As String, ByVal sTypeParam As String, ByVal sName As String, ByVal sNameParam As String, ByVal sReturn As String and ByVal sReturnParam As String for AddErrorHandlingTop and returns Long.
    '
    ' History:  CDK - must handle the case where the function declaration is on more than one line
    '*************************************************************************
                                    
    On Error GoTo ErrorHandle_AddErrorHandlingTop
    Dim sResult As String
    
    sResult = sTemplate
    
    'replace all tokens in the template
    sResult = Replace(sResult, sScopeParam, sScope)
    sResult = Replace(sResult, sTypeParam, sType)
    sResult = Replace(sResult, sNameParam, sName)
    sResult = Replace(sResult, sReturnParam, sReturn)
    sResult = Replace(sResult, sModuleNameParam, sModuleName)
    sResult = Replace(sResult, sDateParam, Format(Now, "YYYYDDMM"))
    
    'read ahead until there is no "_" at the end of the line
    
    
    Call BumpArrayDown(vFileArray, lCurrentElement)
    vFileArray(lCurrentElement) = sResult
    lCurrentElement = lCurrentElement + 1
    
    
    
ErrorHandle_AddErrorHandlingTop:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Function"
    sErrorName = "AddErrorHandlingTop"
    sErrorReturns = "Long"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical, App.Title
          Err.Clear
          Resume Next
    End Select
End Function


Private Function AddErrorHandlingToFileArray(ByRef vFileArray As Variant, _
                                            ByRef lCurrentElement As Long, _
                                            ByVal sTemplate As String, _
                                            ByVal sScope As String, _
                                            ByVal sScopeParam As String, _
                                            ByVal sType As String, _
                                            ByVal sTypeParam As String, _
                                            ByVal sName As String, _
                                            ByVal sNameParam As String, _
                                            ByVal sReturn As String, _
                                            ByVal sReturnParam As String, _
                                            ByVal sModuleName As String, _
                                            ByVal sModuleNameParam As String, _
                                            ByVal sDateParam As String, _
                                            ByVal sDeinstantiationParam As String) As Long
                                
    '*************************************************************************
    ' Scope : Private
    ' Type  : Function
    ' Name  : AddErrorHandling
    ' Param :
    '         ByRef sFile As String
    '         ByVal sTemplate As String
    '         ByVal lPos As String
    '         ByVal sScope As String
    '         ByVal sScopeParam As String
    '         ByVal sType As String
    '         ByVal sTypeParam As String
    '         ByVal sName As String
    '         ByVal sNameParam As String
    '         ByVal sReturn As String
    '         ByVal sReturnParam As String
    ' Returns : Long
    ' Description : The Function uses parameters ByRef sFile As String, ByVal sTemplate As String, ByVal lPos As String, ByVal sScope As String, ByVal sScopeParam As String, ByVal sType As String, ByVal sTypeParam As String, ByVal sName As String, ByVal sNameParam As String, ByVal sReturn As String and ByVal sReturnParam As String for AddErrorHandling and returns Long.
    '*************************************************************************
                                
    On Error GoTo ErrorHandle_AddErrorHandling
    Dim sResult As String
    Dim lTempCurrentElement As Long
    Dim lCount As Long
    Dim sDeinstantiation As String
    
    sResult = sTemplate
    
    'this fixes the 'exit' statement, if it's a prop let get or set, the exit statement got screwed.  but no more
    Select Case sType
        Case mcsProcTypePropLet, mcsProcTypePropGet, mcsProcTypePropSet
            sType = mcsProperty
    End Select
    
    sResult = Replace(sResult, sScopeParam, sScope)
    sResult = Replace(sResult, sTypeParam, sType)
    sResult = Replace(sResult, sNameParam, sName)
    sResult = Replace(sResult, sReturnParam, sReturn)
    sResult = Replace(sResult, sModuleNameParam, sModuleName)
    
    
    'check to see if the error handler has the deinstantiation substitution character in it, if so,
    '   backup and grab all de-instantiations at the bottom of the module and move them to the error handler
    If InStr(1, sResult, sDeinstantiationParam) Then
       
        lTempCurrentElement = lCurrentElement
        
        'include ".close" for any recordset vars taht are being closed
        Do While InStr(1, vFileArray(lTempCurrentElement - 1), "= Nothing") <> 0 _
            Or InStr(1, vFileArray(lTempCurrentElement - 1), ".close") <> 0 _
            Or Trim(vFileArray(lTempCurrentElement - 1)) = ""
            
            lTempCurrentElement = lTempCurrentElement - 1
            
        Loop
    
        'we are either on the first line of closing/deinstantiation of objects
        '   in which case we'll copy and move the elements to the error handler
        'or we're in the last line of the Procedure
        '   in which case the loop won't execute
        For lCount = lTempCurrentElement To lCurrentElement - 1
            'only move the actual deinstantiation calls
            If vFileArray(lCount) & "" <> "" Then
                sDeinstantiation = sDeinstantiation & vFileArray(lCount) & vbCrLf
                vFileArray(lCount) = ""
            End If
        Next
            
    End If
    
    'replace the deinstantiation line with the deinstantiation Substitution character
    '   if there is no deinstantiation lines, replace with ""
    sResult = Replace(sResult, sDeinstantiationParam, sDeinstantiation & "")
    
    Call BumpArrayDown(vFileArray, lCurrentElement)
    vFileArray(lCurrentElement) = sResult
    lCurrentElement = lCurrentElement + 1
    
    

ErrorHandle_AddErrorHandling:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Function"
    sErrorName = "AddErrorHandling"
    sErrorReturns = "Long"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical, App.Title
          Err.Clear
          Resume Next
    End Select
End Function





'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : cmdSaveCommentTemplateToFile_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for cmdSaveCommentTemplateToFile_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Sub cmdSaveCommentTemplateToFile_Click()
    On Error GoTo Proc_Err
    Const csProcName As String = "cmdSaveCommentTemplateToFile_Click"



    Dim sTemplate As String
    sTemplate = GetTemplateFileName(ecdlShowSave, App.Title & mcsDefaultTemplateComment)
    
    If sTemplate & "" <> "" Then
        Call SaveTemplateToFile(sTemplate, Me.txtComments)
    End If


Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
    
    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmCodeCommenter->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Sub


'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : cmdSaveDescriptionTemplateToFile_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for cmdSaveDescriptionTemplateToFile_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Sub cmdSaveDescriptionTemplateToFile_Click()
    On Error GoTo Proc_Err
    Const csProcName As String = "cmdSaveDescriptionTemplateToFile_Click"


    Dim sTemplate As String
    sTemplate = GetTemplateFileName(ecdlShowSave, App.Title & mcsDefaultTemplateDescription)
    
    If sTemplate & "" <> "" Then
        Call SaveTemplateToFile(sTemplate, Me.txtDescription)
    End If



Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
    
    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmCodeCommenter->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Sub


'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : cmdSaveErrorTemplateToFile_Click
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for cmdSaveErrorTemplateToFile_Click and returns Nothing.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Sub cmdSaveErrorTemplateToFile_Click()
    On Error GoTo Proc_Err
    Const csProcName As String = "cmdSaveErrorTemplateToFile_Click"

    
    Dim sTemplate As String
    sTemplate = GetTemplateFileName(ecdlShowSave, App.Title & mcsDefaultTemplateErrorHandling)
    
    If sTemplate & "" <> "" Then
        'Call SaveTemplateToFile(sTemplate  , Me.txtErrorHandlingTop)
        Call SaveTemplateToFile(sTemplate, Me.txtErrorHandling)
    End If



Proc_Exit:
    GoSub Proc_Cleanup
    Exit Sub

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
    
    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmCodeCommenter->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Sub

'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : Form_Load
' Params :
' Returns: Nothing
' Desc   : The Sub uses parameters  for Form_Load and returns Nothing.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Sub Form_Load()
    
    On Error GoTo ErrorHandle_Form_Load
    
    Call CenterFormP(Me)
    
    Dim lCount As Long
    
    cboLang.ListIndex = 0

    'default is unchecked - force user to select the option they want
    Call chkComments_Click
    Call chkErrorHandling_Click
    
    Call Me.fraProgress.Move(mclOffScreen, mclOffScreen)
    
    Call LoadTemplateFromFile(App.Path & "\" & App.Title & mcsDefaultTemplateComment, Me.txtComments)
    Call LoadTemplateFromFile(App.Path & "\" & App.Title & mcsDefaultTemplateDescription, Me.txtDescription)
    Call LoadTemplateFromFile(App.Path & "\" & App.Title & mcsDefaultTemplateErrorHandling, Me.txtErrorHandling)


ErrorHandle_Form_Load:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Sub"
    sErrorName = "Form_Load"
    sErrorReturns = "Nothing"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical, App.Title
          Err.Clear
          Resume Next
    End Select
End Sub

'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : Form_Unload
' Params :
'          Cancel As Integer
' Returns: Nothing
' Desc   : The Sub uses parameters Cancel As Integer for Form_Unload and returns Nothing.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    Set moHourglass = Nothing

End Sub

'__________________________________________________
' Scope  : Private
' Type   : Function
' Name   : BrowseFolder
' Params :
' Returns: String
' Desc   : The Function uses parameters  for BrowseFolder and returns String.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Function BrowseFolder() As String
    On Error GoTo ErrorHandle_BrowseFolder
    Dim lIDList As Long
    Dim sPath As String
    Dim uBrowse As BrowseInfo
    
    uBrowse.hWndOwner = Me.hWnd
    uBrowse.lpszTitle = StrPtr("Choose folder" + vbNullChar)
    uBrowse.ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    lIDList = SHBrowseForFolder(uBrowse)
    If lIDList Then
        sPath = Space(MAX_PATH)
        SHGetPathFromIDList lIDList, sPath
        sPath = Left(sPath, InStr(sPath, vbNullChar) - 1)
    End If
    
    BrowseFolder = sPath

ErrorHandle_BrowseFolder:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Function"
    sErrorName = "BrowseFolder"
    sErrorReturns = "String"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical, App.Title
          Err.Clear
          Resume Next
    End Select
End Function



'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : LoadTemplateFromFile
' Params :
'          sTemplateToLoad As String
'          oCtlToLoad As Control
' Returns: Nothing
' Desc   : The Sub uses parameters sTemplateToLoad As String and oCtlToLoad As Control for LoadTemplateFromFile and returns Nothing.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Sub LoadTemplateFromFile(sTemplateToLoad As String, oCtlToLoad As Control)

    On Error GoTo ErrorHandle_LoadTemplateFromFile

    Dim ofs As clsFs
    Dim oTs As TextStream
    
    Set ofs = New clsFs
    
            
    'see if the file exists for the default information
    If ofs.FileExists(sTemplateToLoad) Then
        Set oTs = ofs.GetFile(sTemplateToLoad).OpenAsTextStream
        oCtlToLoad = oTs.ReadAll
    End If
    
ErrorHandle_LoadTemplateFromFile:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Function"
    sErrorName = "LoadTemplateFromFile"
    sErrorReturns = "String"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical, App.Title
          Err.Clear
          Resume Next
    End Select
    
    On Error Resume Next
    Set ofs = Nothing
    Set oTs = Nothing
    
End Sub




'__________________________________________________
' Scope  : Private
' Type   : Sub
' Name   : SaveTemplateToFile
' Params :
'          sFileName As String
'          oCtl As Control
' Returns: Nothing
' Desc   : The Sub uses parameters sFileName As String and oCtl As Control for SaveTemplateToFile and returns Nothing.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Sub SaveTemplateToFile(sFileName As String, oCtl As Control)

    On Error GoTo ErrorHandle_SaveTemplateToFile

    Dim ofs As clsFs
    Dim oTs As TextStream
    
    Set ofs = New clsFs
    Set oTs = ofs.CreateTextFile(sFileName, True, False)
    oTs.Write oCtl.Text
    
    Set oTs = Nothing
    
    
ErrorHandle_SaveTemplateToFile:
    Dim sErrorScope As String
    Dim sErrorType As String
    Dim sErrorName As String
    Dim sErrorReturns As String

    Dim sErrorMsgBox As String

    sErrorScope = "Private"
    sErrorType = "Function"
    sErrorName = "SaveTemplateToFile"
    sErrorReturns = "String"
  
    Select Case Err.Number
        Case vbEmpty
          'Nothing
        Case Else
          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
          MsgBox sErrorMsgBox, vbCritical, App.Title
          Err.Clear
          Resume Next
          Resume
    End Select
    
    On Error Resume Next
    Set ofs = Nothing
    Set oTs = Nothing
    
End Sub



'__________________________________________________
' Scope  : Private
' Type   : Function
' Name   : GetAllFileExtensions
' Params :
'          sDir As String
'          ByRef vArrToFill As Variant
'          Optional fIncludeHidden As Boolean = True
'          Optional fIncludeSystem As Boolean = True
' Returns: Long
' Desc   : The Function uses parameters sDir As String, ByRef vArrToFill As Variant, Optional fIncludeHidden As Boolean = True and Optional fIncludeSystem As Boolean = True for GetAllFileExtensions and returns Long.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Function GetAllFileExtensions(sDir As String, ByRef vArrToFill As Variant, Optional fIncludeHidden As Boolean = True, Optional fIncludeSystem As Boolean = True) As Long
    On Error GoTo Proc_Err
    Const csProcName As String = "GetAllFileExtensions"


    'returns all file extensions in an array for hte given folder

    Dim ofs As clsFileOps
    Dim lCount As Long
    Dim lPosition As Long
    Dim sTypesAdded As String
    Dim sFile As String
    Dim sExtension As String
    Dim vArrFileTypes() As String
    Set ofs = New clsFileOps
    Dim lTotalTypes As Long
    
    
    lCount = ofs.FilesToArray(sDir, fIncludeHidden, fIncludeSystem, vArrFileTypes)
    
    If lCount > 0 Then
        
        For lCount = LBound(vArrFileTypes) To UBound(vArrFileTypes)
        
            'check to see if the value is already in the array
             lPosition = InStrRev(vArrFileTypes(lCount), ".")
             sFile = vArrFileTypes(lCount)
             sExtension = Replace(sFile, Left(sFile, lPosition - 1), "")
             If InStr(1, sTypesAdded, sExtension, vbTextCompare) = 0 Then
                sTypesAdded = sTypesAdded & sExtension & "|"
                lTotalTypes = lTotalTypes + 1
             End If
             
        Next
        
        vArrToFill = Split(sTypesAdded, "|")
                
    End If
    
    GetAllFileExtensions = lTotalTypes


Proc_Exit:
    GoSub Proc_Cleanup
    Exit Function

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
        

    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmCodeCommenter->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Function


'__________________________________________________
' Scope  : Private
' Type   : Function
' Name   : GetModuleName
' Params :
'          ByVal varr As Variant
' Returns: String
' Desc   : The Function uses parameters ByVal varr As Variant for GetModuleName and returns String.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Private Function GetModuleName(ByVal varr As Variant) As String
    On Error GoTo Proc_Err
    Const csProcName As String = "GetModuleName"


    Dim lCountCloneArrElement As Long
    Dim vCloneArr As Variant
    Dim lPos As Long
    Dim sLine As String
    Dim lMaxIterations As Long
    
    vCloneArr = varr
    
    lMaxIterations = UBound(vCloneArr)
    lCountCloneArrElement = 0
    
    While Not lCountCloneArrElement > lMaxIterations
        
        sLine = vCloneArr(lCountCloneArrElement)
        lPos = InStr(1, sLine, mcsModuleNameIdentifier, vbTextCompare)
        If lPos > 0 Then
            GetModuleName = Replace(sLine, mcsModuleNameIdentifier, "")
            GetModuleName = Replace(GetModuleName, """", "")
            lCountCloneArrElement = lMaxIterations + 1
        End If
        
        lCountCloneArrElement = lCountCloneArrElement + 1
    Wend



Proc_Exit:
    GoSub Proc_Cleanup
    Exit Function

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
        

    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmCodeCommenter->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Function



'__________________________________________________
' Scope  : Public
' Type   : Function
' Name   : BumpArrayDown
' Params :
'          ByRef varr As Variant
'          lStartingElement As Long
' Returns: Nothing
' Desc   : The Function uses parameters ByRef varr As Variant and lStartingElement As Long for BumpArrayDown and returns Nothing.
'__________________________________________________
' History
' CDK: 20010102: Added Error Trapping & Comments
'__________________________________________________
Public Function BumpArrayDown(ByRef varr As Variant, ByVal lStartingElement As Long)
    On Error GoTo Proc_Err
    Const csProcName As String = "BumpArrayDown"


    Dim lLbound As Long
    Dim lUbound As Long
    Dim lCount As Long
    Dim lMax As Long
    
    lLbound = LBound(varr)
    lUbound = UBound(varr)
    lMax = lUbound - 1
    
    ReDim Preserve varr(lUbound + 1)
    For lCount = lMax To lStartingElement Step -1
    
        varr(lCount + 1) = varr(lCount)

    Next
    
    varr(lStartingElement) = ""


Proc_Exit:
    GoSub Proc_Cleanup
    Exit Function

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
        

    On Error GoTo 0
    Return

Proc_Err:
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmCodeCommenter->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Function


Private Function getTabLenFromNextLineMl(ByVal varr As Variant, ByVal lElementNum As Long) As Long

    Const csProcName As String = "getTabLenFromNextLineMl"

    Dim vClone As Variant
    Dim lMax As Long
    Dim lMin As Long
    Dim lCount As Long
    Dim lCount2 As Long
    Dim lLen As Long
    Dim lTabLen As Long
    
    vClone = varr
    
    lMin = lElementNum
    lMax = UBound(vClone)
    
    lTabLen = 0
    
    'check the next line of the array, if there is data there, take the leading tab
    For lCount = lMin To lMax
        'ignore empty lines
        If Trim(vClone(lCount)) <> "" Then
        
            'ignore lines without tabs
            If Left(vClone(lCount), 1) = " " Then
                    
                lLen = Len(vClone(lCount))
                
                For lCount2 = 1 To lLen
                
                    If InStr(1, vClone(lCount), String(lCount2, " ")) = 0 Then
                        'set it and get out
                        lTabLen = lCount2 - 1
                        lCount2 = lLen + 1
                        lCount = lMax + 1
                    End If
                    
                Next
                
            End If
            
        End If
            
    Next
    

    getTabLenFromNextLineMl = lTabLen

Proc_Exit:
    GoSub Proc_Cleanup
    Exit Function

Proc_Cleanup:
    On Error Resume Next
    'Place any cleanup of instantiated objects here
        

    On Error GoTo 0
    Return

Proc_Err:

    getTabLenFromNextLineMl = 0
    
    Dim lErrNum As String, sErrSource As String, sErrDesc As String
    lErrNum = VBA.Err.Number
    sErrSource = VBA.Err.Source & vbCrLf & "frmCodeCommenter->" & csProcName
    sErrDesc = VBA.Err.Description
    Resume Proc_Err_Continue
    Resume
    
Proc_Err_Continue:
    GoSub Proc_Cleanup
    MsgBox "The following Error Occurred" & vbCrLf & vbCrLf & _
                 "Source:" & vbCrLf & sErrSource & vbCrLf & vbCrLf & _
                 "Number:" & vbCrLf & lErrNum & vbCrLf & vbCrLf & _
                 "Description:" & vbCrLf & sErrDesc, vbExclamation + vbOKOnly

End Function




'Private Function AddErrorHandlingTop(ByRef sFile As String, _
'                                    ByVal sTemplate As String, _
'                                    ByVal lPos As String, _
'                                    ByVal sScope As String, _
'                                    ByVal sScopeParam As String, _
'                                    ByVal sType As String, _
'                                    ByVal sTypeParam As String, _
'                                    ByVal sName As String, _
'                                    ByVal sNameParam As String, _
'                                    ByVal sReturn As String, _
'                                    ByVal sReturnParam As String, _
'                                    ByVal sModuleName As String, _
'                                    ByVal sModuleNameParam As String, _
'                                    ByVal sDateParam As String) As Long
'    '*************************************************************************
'    ' Scope : Private
'    ' Type : Function
'    ' Name : AddErrorHandlingTop
'    ' Parameters :
'    '         ByRef sFile As String
'    '         ByVal sTemplate As String
'    '         ByVal lPos As String
'    '         ByVal sScope As String
'    '         ByVal sScopeParam As String
'    '         ByVal sType As String
'    '         ByVal sTypeParam As String
'    '         ByVal sName As String
'    '         ByVal sNameParam As String
'    '         ByVal sReturn As String
'    '         ByVal sReturnParam As String
'    ' Returns : Long
'    ' Description : The Function uses parameters ByRef sFile As String, ByVal sTemplate As String, ByVal lPos As String, ByVal sScope As String, ByVal sScopeParam As String, ByVal sType As String, ByVal sTypeParam As String, ByVal sName As String, ByVal sNameParam As String, ByVal sReturn As String and ByVal sReturnParam As String for AddErrorHandlingTop and returns Long.
'    '
'    ' History:  CDK - must handle the case where the function declaration is on more than one line
'    '*************************************************************************
'
'    On Error GoTo ErrorHandle_AddErrorHandlingTop
'    Dim sResult As String
'
'    sResult = sTemplate
'
'    'replace all tokens in the template
'    sResult = Replace(sResult, sScopeParam, sScope)
'    sResult = Replace(sResult, sTypeParam, sType)
'    sResult = Replace(sResult, sNameParam, sName)
'    sResult = Replace(sResult, sReturnParam, sReturn)
'    sResult = Replace(sResult, sModuleNameParam, sModuleName)
'    sResult = Replace(sResult, sDateParam, Format(Now, "YYYYDDMM"))
'
'    'read ahead until there is no "_" at the end of the line
'
'
'
'    sFile = Left(sFile, lPos - 1) + sResult + Mid(sFile, lPos)
'
'    AddErrorHandlingTop = Len(sResult)
'
'ErrorHandle_AddErrorHandlingTop:
'    Dim sErrorScope As String
'    Dim sErrorType As String
'    Dim sErrorName As String
'    Dim sErrorReturns As String
'
'    Dim sErrorMsgBox As String
'
'    sErrorScope = "Private"
'    sErrorType = "Function"
'    sErrorName = "AddErrorHandlingTop"
'    sErrorReturns = "Long"
'
'    Select Case Err.Number
'        Case vbEmpty
'          'Nothing
'        Case Else
'          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
'          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
'          MsgBox sErrorMsgBox, vbCritical, App.Title
'          Err.Clear
'          Resume Next
'    End Select
'End Function




'Private Function AddComments(ByRef sFile As String, _
'                            ByVal sTemplate As String, _
'                            ByVal lPos As Long, _
'                            ByVal sScope As String, _
'                            ByVal sScopeParam As String, _
'                            ByVal sType As String, _
'                            ByVal sTypeParam As String, _
'                            ByVal sName As String, _
'                            ByVal sNameParam As String, _
'                            ByVal vArrParameters As Variant, _
'                            ByVal sParametersParam As String, _
'                            ByVal sReturn As String, _
'                            ByVal sReturnParam As String, _
'                            ByVal sDescription As String, _
'                            ByVal sDescriptionParam As String, _
'                            ByVal sModuleName As String, _
'                            ByVal sModuleNameParam As String, _
'                            ByVal sDateParam As String) As Long
'
'                            '*************************************************************************
'                            ' Scope : Private
'                            ' Type : Function
'                            ' Name : AddComments
'                            ' Parameters :
'                            '         ByRef sFile As String
'                            '         ByVal sTemplate As String
'                            '         ByVal lPos As Long
'                            '         ByVal sScope As String
'                            '         ByVal sScopeParam As String
'                            '         ByVal sType As String
'                            '         ByVal sTypeParam As String
'                            '         ByVal sName As String
'                            '         ByVal sNameParam As String
'                            '         ByVal vArrParameters As Variant
'                            '         ByVal sParametersParam As String
'                            '         ByVal sReturn As String
'                            '         ByVal sReturnParam As String
'                            '         ByVal sDescription As String
'                            '         ByVal sDescriptionParam As String
'                            ' Returns : Long
'                            ' Description : The Function uses parameters ByRef sFile As String, ByVal sTemplate As String, ByVal lPos As Long, ByVal sScope As String, ByVal sScopeParam As String, ByVal sType As String, ByVal sTypeParam As String, ByVal sName As String, ByVal sNameParam As String, ByVal vArrParameters As Variant, ByVal sParametersParam As String, ByVal sReturn As String, ByVal sReturnParam As String, ByVal sDescription As String and ByVal sDescriptionParam As String for AddComments and returns Long.
'                            '*************************************************************************
'
'
'    On Error GoTo ErrorHandle_AddComments
'    Dim lCount As Long
'    Dim sResult As String
'    Dim sParamLine As String
'    Dim sParams As String
'    Dim lParamPos As Long
'    Dim lLastCrLf As Long
'    Dim lNextCrLf As Long
'    Dim sAnd As String
'
'    sResult = sTemplate
'
'    sResult = Replace(sResult, sScopeParam, sScope)
'    sResult = Replace(sResult, sTypeParam, sType)
'    sResult = Replace(sResult, sNameParam, sName)
'    sResult = Replace(sResult, sModuleNameParam, sModuleName)
'    sResult = Replace(sResult, sDateParam, Format(Now, "YYYYDDMM"))
'
'    lParamPos = InStr(sResult, sParametersParam)
'
'    If lParamPos > 0 Then
'        lLastCrLf = InStrRev(sResult, vbCrLf, lParamPos)
'        lNextCrLf = InStr(lParamPos, sResult, vbCrLf)
'        If lNextCrLf - lLastCrLf > Len(sParametersParam) Then
'            sParamLine = Mid(sResult, IIf(lLastCrLf > 0, lLastCrLf, 1), lNextCrLf - IIf(lLastCrLf > 0, lLastCrLf, 1))
'            For lCount = LBound(vArrParameters) To UBound(vArrParameters)
'                sParams = sParams + Replace(sParamLine, sParametersParam, vArrParameters(lCount))
'            Next lCount
'            sResult = Replace(sResult, sParamLine, sParams)
'        Else
'            For lCount = LBound(vArrParameters) To UBound(vArrParameters) - 1
'                sParams = sParams + vArrParameters(lCount) + ", "
'            Next lCount
'            sAnd = " and "
'            If cboLang.Text = LANG_ENG Then sAnd = " and "
'            If cboLang.Text = LANG_FRA Then sAnd = " et "
'            If cboLang.Text = LANG_ESP Then sAnd = " e "
'            If cboLang.Text = LANG_DEU Then sAnd = " und "
'            If Not UBound(vArrParameters) = -1 Then If UBound(vArrParameters) > LBound(vArrParameters) Then sParams = Left(sParams, Len(sParams) - 2) + sAnd + vArrParameters(UBound(vArrParameters)) Else sParams = vArrParameters(LBound(vArrParameters))
'            sResult = Replace(sResult, sParametersParam, sParams)
'        End If
'    End If
'    sResult = Replace(sResult, sReturnParam, sReturn)
'    sResult = Replace(sResult, sDescriptionParam, sDescription)
'
'    sFile = Left(sFile, lPos - 1) + sResult + Mid(sFile, lPos)
'
'    AddComments = Len(sResult)
'
'ErrorHandle_AddComments:
'    Dim sErrorScope As String
'    Dim sErrorType As String
'    Dim sErrorName As String
'    Dim sErrorReturns As String
'
'    Dim sErrorMsgBox As String
'
'    sErrorScope = "Private"
'    sErrorType = "Function"
'    sErrorName = "AddComments"
'    sErrorReturns = "Long"
'
'    Select Case Err.Number
'        Case vbEmpty
'          'Nothing
'        Case Else
'          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
'          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
'          MsgBox sErrorMsgBox, vbCritical, App.Title
'          Err.Clear
'          Resume Next
'    End Select
'End Function
'
'
'


'Private Function AddErrorHandling(ByRef sFile As String, _
'                                ByVal sTemplate As String, _
'                                ByVal lPos As String, _
'                                ByVal sScope As String, _
'                                ByVal sScopeParam As String, _
'                                ByVal sType As String, _
'                                ByVal sTypeParam As String, _
'                                ByVal sName As String, _
'                                ByVal sNameParam As String, _
'                                ByVal sReturn As String, _
'                                ByVal sReturnParam As String, _
'                                ByVal sModuleName As String, _
'                                ByVal sModuleNameParam As String, _
'                                ByVal sDateParam As String) As Long
'
'                                '*************************************************************************
'                                ' Scope : Private
'                                ' Type : Function
'                                ' Name : AddErrorHandling
'                                ' Parameters :
'                                '         ByRef sFile As String
'                                '         ByVal sTemplate As String
'                                '         ByVal lPos As String
'                                '         ByVal sScope As String
'                                '         ByVal sScopeParam As String
'                                '         ByVal sType As String
'                                '         ByVal sTypeParam As String
'                                '         ByVal sName As String
'                                '         ByVal sNameParam As String
'                                '         ByVal sReturn As String
'                                '         ByVal sReturnParam As String
'                                ' Returns : Long
'                                ' Description : The Function uses parameters ByRef sFile As String, ByVal sTemplate As String, ByVal lPos As String, ByVal sScope As String, ByVal sScopeParam As String, ByVal sType As String, ByVal sTypeParam As String, ByVal sName As String, ByVal sNameParam As String, ByVal sReturn As String and ByVal sReturnParam As String for AddErrorHandling and returns Long.
'                                '*************************************************************************
'
'    On Error GoTo ErrorHandle_AddErrorHandling
'    Dim sResult As String
'
'    sResult = sTemplate
'
'    'this fixes the 'exit' statement, if it's a prop let get or set, the exit statement got screwed.  but no more
'    Select Case sType
'        Case mcsProcTypePropLet, mcsProcTypePropGet, mcsProcTypePropSet
'            sType = mcsProperty
'    End Select
'
'    sResult = Replace(sResult, sScopeParam, sScope)
'    sResult = Replace(sResult, sTypeParam, sType)
'    sResult = Replace(sResult, sNameParam, sName)
'    sResult = Replace(sResult, sReturnParam, sReturn)
'    sResult = Replace(sResult, sModuleNameParam, sModuleName)
'
'    sFile = Left(sFile, lPos - 1) + sResult + Mid(sFile, lPos)
'
'    AddErrorHandling = Len(sResult)
'
'ErrorHandle_AddErrorHandling:
'    Dim sErrorScope As String
'    Dim sErrorType As String
'    Dim sErrorName As String
'    Dim sErrorReturns As String
'
'    Dim sErrorMsgBox As String
'
'    sErrorScope = "Private"
'    sErrorType = "Function"
'    sErrorName = "AddErrorHandling"
'    sErrorReturns = "Long"
'
'    Select Case Err.Number
'        Case vbEmpty
'          'Nothing
'        Case Else
'          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
'          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
'          MsgBox sErrorMsgBox, vbCritical, App.Title
'          Err.Clear
'          Resume Next
'    End Select
'End Function
'


'Private Function ModifyFile(ByVal sFileName As String, ByVal bComments As Boolean, ByVal bErrorHandling As Boolean, ByVal bMakeBackup As Boolean, Optional bDoNotDisplay As Boolean = False) As String
'
'    On Error GoTo ErrorHandle_ModifyFile
'    Dim sLastScope As String
'
'    Dim sModuleName As String
'    Dim sLastMid As String
'    Dim sLastType As String
'    Dim sLastName As String
'    Dim sLastReturn As String
'    Dim sScope As String
'    Dim sMid As String
'    Dim sType As String
'    Dim sName As String
'    Dim vArrParameters As Variant
'    Dim sReturn As String
'    Dim sDescription As String
'    Dim sEnd As String
'    Dim sSpace As String
'
'    Dim bStartErrorHandling As Boolean
'    Dim sFile As String
'    Dim vArrFileContents As Variant
'    Dim iOpen As Integer
'    Dim lCount As Long
'    Dim lLbound As Long
'    Dim lUbound As Long
'    Dim lChar As Long
'    Dim lPos As Long
'
'
'    Dim sMsgBox As String
'    Dim bUp As Long
'
'
'    vArrScopes = Array("Private", "Public", "Global", "Friend", "Protected")
'    vArrMids = Array("Static")
'    vArrProcedures = Array("Function", "Sub", mcsProcTypePropLet, mcsProcTypePropGet, mcsProcTypePropSet)
'    vArrEnds = Array("End")
'    vArrOnError = Array("On Error")
'    vArrProcedureEnds = Array("Function", "Sub", "Property")
'
'    If bMakeBackup Then FileCopy sFileName, sFileName + BACKUP_EXT
'
'    iOpen = FreeFile(1)
'    Open sFileName For Input As iOpen
'        sFile = Input(LOF(iOpen), iOpen)
'    Close iOpen
'
'    vArrFileContents = Split(sFile, vbCrLf)
'    lChar = 1
'
'    lLbound = LBound(vArrFileContents)
'    lUbound = UBound(vArrFileContents)
'
'    'get the mod name
'    sModuleName = GetModuleName(vArrFileContents)
'
'    For lCount = lLbound To lUbound
'        lPos = 1
'        sScope = ""
'        sMid = ""
'        sType = ""
'        sName = ""
'        sReturn = ""
'        vArrParameters = Null
'        sDescription = ""
'        sEnd = ""
'
'
'        bUp = False
'
'        sScope = CheckScope(vArrFileContents(lCount), lPos)
'        lPos = lPos + IIf(Len(sScope) = 0, 0, Len(sScope) + 1)
'        sMid = CheckMid(vArrFileContents(lCount), lPos)
'        lPos = lPos + IIf(Len(sMid) = 0, 0, Len(sMid) + 1)
'
'        If (Len(sScope) > 0 And Len(sMid) > 0) Then
'            sSpace = " "
'        Else
'            sSpace = ""
'        End If
'
'        'check to see if the array element is a function/sub/property let/get etc
'        sType = CheckProcedure(vArrFileContents(lCount), lPos)
'
'        If Len(sType) > 0 Then
'            lProcedures = lProcedures + 1
'            lPos = lPos + Len(sType) + 1
'            sName = GetName(vArrFileContents(lCount), lPos)
'
'            'if this is a new code block, start the commenter / error handler insertion
'            If Len(sName) > 0 Then
'                lPos = lPos + Len(sName) + 1
'                sReturn = GetReturn(vArrFileContents(lCount))
'                If bComments Then
'                    vArrParameters = GetParams(vArrFileContents(lCount), lPos)
'
'
'                    sDescription = MakeDescription(txtDescription.Text, sScope + sSpace + sMid, PROCSCOPE, _
'                                                    sType, PROCTYPE, sName, PROCNAME, _
'                                                    vArrParameters, PROCPARAM, _
'                                                    sReturn, PROCRETURN, _
'                                                    sModuleName, PROCMODULENAME, CurrentDateTime)
'
'                    lChar = lChar + AddComments(sFile, txtComments.Text, lChar, sScope + sSpace + sMid, PROCSCOPE, _
'                                                    sType, PROCTYPE, sName, PROCNAME, vArrParameters, _
'                                                    PROCPARAM, sReturn, PROCRETURN, sDescription, PROCDESC, _
'                                                    sModuleName, PROCMODULENAME, CurrentDateTime)
'
'                    lModProcedures = lModProcedures + 1
'                    bUp = True
'                End If
'                If bErrorHandling Then
'
'                    'check here to see if there is already error handling in the code
'                    '   look ahead until 'end' is encountered if we find an on error
'                    '   then skip this code block
'                    If Len(LookAhead(vArrFileContents, lCount, vArrOnError, vArrEnds)) > 0 Then
'                        bStartErrorHandling = False
'                    Else
'                        bStartErrorHandling = True
'                        sLastScope = sScope
'                        sLastMid = sMid
'                        sLastType = sType
'                        sLastName = sName
'                        sLastReturn = sReturn
'
'                        'we must check to see if the proc declaration was continued to the next line, if so,
'                        '   account for the lenght of that line, then move to the next
'                        Do While Right(vArrFileContents(lCount), 1) = mcsLineContinuationChar
'                            lChar = lChar + Len(vArrFileContents(lCount)) + Len(vbCrLf)
'                            lCount = lCount + 1
'                        Loop
'
'                        lChar = lChar + AddErrorHandlingTop(sFile, _
'                                                            txtErrorHandlingTop.Text, _
'                                                            lChar + Len(vArrFileContents(lCount)) + Len(vbCrLf), _
'                                                            sScope + sSpace + sMid, _
'                                                            PROCSCOPE, sType, PROCTYPE, sName, PROCNAME, sReturn, _
'                                                            PROCRETURN, sModuleName, PROCMODULENAME, CurrentDateTime)
'                        If Not bUp Then
'                            lModProcedures = lModProcedures + 1
'                        End If
'                    End If
'                End If
'            End If
'        End If
'        If bErrorHandling And bStartErrorHandling Then
'            lPos = 1
'            sEnd = CheckEnd(vArrFileContents(lCount))
'            lPos = lPos + IIf(Len(sEnd) = 0, 0, Len(sEnd) + 1)
'            If Len(sEnd) > 0 And Len(CheckProcedureEnd(vArrFileContents(lCount), lPos)) > 0 Then
'                lChar = lChar + AddErrorHandling(sFile, txtErrorHandling.Text, lChar, _
'                                                sLastScope + IIf(Len(sLastScope) > 0 And Len(sLastMid) > 0, " ", "") + sLastMid, _
'                                                PROCSCOPE, sLastType, PROCTYPE, sLastName, PROCNAME, sLastReturn, _
'                                                PROCRETURN, sModuleName, PROCMODULENAME, CurrentDateTime)
'                bStartErrorHandling = False
'                sLastScope = ""
'                sLastMid = ""
'                sLastType = ""
'                sLastName = ""
'                sLastReturn = ""
'            End If
'        End If
'
'        lChar = lChar + Len(vArrFileContents(lCount)) + Len(vbCrLf)
'
'    Next lCount
'
'    iOpen = FreeFile(1)
'    Open sFileName For Output As iOpen
'        Print #iOpen, sFile
'    Close iOpen
'
'    If Not bDoNotDisplay Then
'        sMsgBox = IIf(bComments, "comments", "")
'        sMsgBox = sMsgBox + IIf(Len(sMsgBox) > 0 And bErrorHandling, " and ", "") + IIf(bErrorHandling, "error handling", "")
'        MsgBox "Finished adding " + IIf(Len(sMsgBox) > 0, sMsgBox, "nothing") + " to " + sFileName + vbCrLf + vbCrLf + "With a total of " + CStr(lModProcedures) + " procedures modified on a great total of " + CStr(lProcedures), vbApplicationModal, "Commentor"
'        lProcedures = 0
'    End If
'
'    ModifyFile = sFileName
'
'ErrorHandle_ModifyFile:
'    Dim sErrorScope As String
'    Dim sErrorType As String
'    Dim sErrorName As String
'    Dim sErrorReturns As String
'
'    Dim sErrorMsgBox As String
'
'    sErrorScope = "Private"
'    sErrorType = "Function"
'    sErrorName = "ModifyFile"
'    sErrorReturns = "String"
'
'    Select Case Err.Number
'        Case vbEmpty
'          'Nothing
'        Case Else
'          sErrorMsgBox = "Error " + CStr(Err.Number) + " has occured."
'          sErrorMsgBox = sErrorMsgBox + vbCrLf + Err.Description
'          MsgBox sErrorMsgBox, vbCritical, App.Title
'          Err.Clear
'          Resume Next
'          Resume
'
'    End Select
'End Function
'



