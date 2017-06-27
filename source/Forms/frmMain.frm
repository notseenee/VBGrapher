VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VBGrapher"
   ClientHeight    =   7665
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   6240
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   6240
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialogFont 
      Left            =   4725
      Top             =   1470
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtSelectedFunc 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   225
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "Add a function below"
      ToolTipText     =   "Select the current function to copy and paste"
      Top             =   150
      Width           =   5625
   End
   Begin MSComDlg.CommonDialog CommonDialogSave 
      Left            =   3330
      Top             =   945
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   ".bmp"
      DialogTitle     =   "Save Graph"
      FileName        =   "graph"
      Filter          =   "Bitmap (*.bmp)|*.bmp|All files (*.*)|*.*"
      InitDir         =   "%userprofile%\Desktop"
   End
   Begin MSComDlg.CommonDialog CommonDialogColor 
      Left            =   2430
      Top             =   870
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClearGraph 
      Caption         =   "Clear"
      Height          =   720
      Left            =   4215
      MaskColor       =   &H00F0F0F0&
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Remove all functions from the graph"
      Top             =   660
      UseMaskColor    =   -1  'True
      Width           =   780
   End
   Begin VB.CommandButton cmdEditColour 
      Caption         =   "Colour"
      Enabled         =   0   'False
      Height          =   720
      Left            =   2175
      MaskColor       =   &H00F0F0F0&
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Edit the colour of the selected function"
      Top             =   660
      UseMaskColor    =   -1  'True
      Width           =   780
   End
   Begin VB.CommandButton cmdSaveGraph 
      Caption         =   "Save"
      Height          =   720
      Left            =   3315
      MaskColor       =   &H00F0F0F0&
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Save the graph to a file"
      Top             =   660
      UseMaskColor    =   -1  'True
      Width           =   780
   End
   Begin VB.CommandButton cmdRemoveFunc 
      Caption         =   "Remove"
      Enabled         =   0   'False
      Height          =   720
      Left            =   1035
      MaskColor       =   &H00F0F0F0&
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Remove the selected function"
      Top             =   660
      UseMaskColor    =   -1  'True
      Width           =   780
   End
   Begin VB.CommandButton cmdAddFunc 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   720
      Left            =   120
      MaskColor       =   &H00F0F0F0&
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Add a new function"
      Top             =   660
      UseMaskColor    =   -1  'True
      Width           =   780
   End
   Begin VB.PictureBox pctCrosshairs 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6000
      Left            =   120
      MousePointer    =   2  'Cross
      ScaleHeight     =   200
      ScaleLeft       =   -100
      ScaleMode       =   0  'User
      ScaleTop        =   -100
      ScaleWidth      =   200
      TabIndex        =   3
      Top             =   1545
      Width           =   6000
   End
   Begin VB.PictureBox pctGraph 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6000
      Left            =   300
      MousePointer    =   2  'Cross
      ScaleHeight     =   200
      ScaleLeft       =   -100
      ScaleMode       =   0  'User
      ScaleTop        =   -100
      ScaleWidth      =   200
      TabIndex        =   0
      Top             =   1725
      Visible         =   0   'False
      Width           =   6000
   End
   Begin VB.ComboBox cmbFuncs 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      ItemData        =   "frmMain.frx":000C
      Left            =   120
      List            =   "frmMain.frx":000E
      TabIndex        =   11
      Text            =   "cmbFuncs"
      ToolTipText     =   "Choose a different function from the graph"
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label lblXVal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "-10.00"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.0000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   0
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5190
      TabIndex        =   14
      Top             =   660
      Width           =   600
   End
   Begin VB.Label lblYVal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "-10.00"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.0000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   0
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5190
      TabIndex        =   13
      Top             =   915
      Width           =   600
   End
   Begin VB.Label lblFXVal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "-10.00"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.0000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   0
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5190
      TabIndex        =   12
      Top             =   1170
      Width           =   600
   End
   Begin VB.Label lblFX 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "f(x)"
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   5850
      TabIndex        =   9
      Top             =   1170
      Width           =   270
   End
   Begin VB.Label lblY 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "y"
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   5850
      TabIndex        =   2
      Top             =   915
      Width           =   270
   End
   Begin VB.Label lblX 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "x"
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   5850
      TabIndex        =   1
      Top             =   660
      Width           =   270
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuAddFunc 
         Caption         =   "&Add Function…"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemoveFunc 
         Caption         =   "&Remove Function"
         Enabled         =   0   'False
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuEditColour 
         Caption         =   "&Edit Colour…"
         Enabled         =   0   'False
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveGraph 
         Caption         =   "&Save Graph…"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuClearGraph 
         Caption         =   "&Clear Graph"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptShowPoints 
         Caption         =   "Show &Points on Hover"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOptShowCrosshairs 
         Caption         =   "Show &Crosshairs on Hover"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOptSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptFont 
         Caption         =   "Change Graph &Font…"
      End
      Begin VB.Menu mnuOptSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptShowDebugMenu 
         Caption         =   "Show &Debug Menu"
      End
   End
   Begin VB.Menu mnuDebug 
      Caption         =   "&Debug"
      Visible         =   0   'False
      Begin VB.Menu mnuFuncList 
         Caption         =   "Function &List"
      End
      Begin VB.Menu mnuTestPostfix 
         Caption         =   "Test &Postfix Converter"
      End
      Begin VB.Menu mnuTestEvaluator 
         Caption         =   "Test &Evaluator"
      End
   End
   Begin VB.Menu mnuHelpMenu 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help…"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuTroubleshooting 
         Caption         =   "&Troubleshooting…"
      End
      Begin VB.Menu mnuHelpSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About…"
      End
      Begin VB.Menu mnuLicence 
         Caption         =   "&Licence…"
      End
      Begin VB.Menu mnuAcknowldegements 
         Caption         =   "Ac&knowledgements…"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intActivations As Integer

Private Sub Form_Load()
    ' Initialise the funcList array
    Call funcListInitialise
    
    ' Number of times form has been activated
    intActivations = 0
    
    Randomize
    
    ' Add window icon
    Me.Icon = LoadResPicture(101, vbResIcon)
    
    ' Add icons to buttons
    cmdAddFunc.Picture = LoadResPicture(101, vbResBitmap)
    cmdRemoveFunc.Picture = LoadResPicture(106, vbResBitmap)
    cmdClearGraph.Picture = LoadResPicture(102, vbResBitmap)
    cmdSaveGraph.Picture = LoadResPicture(107, vbResBitmap)
    cmdEditColour.Picture = LoadResPicture(103, vbResBitmap)
End Sub

Private Sub Form_Activate()
    ' If first time run, draw axes
    If intActivations = 0 Then
        clearGraph True
    End If
    
    intActivations = intActivations + 1
    
    ' Sync crosshairs picturebox and normal pbox
    Call resetCrosshairs
    
    ' Update functions dropdown
    Call updateDropdown
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' FILE MENU
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuAddFunc_Click()
    frmAddFunc.Show , Me
End Sub
Private Sub cmdAddFunc_Click()
    frmAddFunc.Show , Me
End Sub

Private Sub mnuRemoveFunc_Click()
    Call removeFunc
End Sub
Private Sub cmdRemoveFunc_Click()
    Call removeFunc
End Sub

Private Sub mnuEditColour_Click()
    Call editColour
End Sub
Private Sub cmdEditColour_Click()
    Call editColour
End Sub

Private Sub mnuSaveGraph_Click()
    Call saveGraph
End Sub
Private Sub cmdSaveGraph_Click()
    Call saveGraph
End Sub

Private Sub mnuClearGraph_Click()
    clearGraph True
End Sub
Private Sub cmdClearGraph_click()
    clearGraph True
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' HELP MENU
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub mnuHelp_Click()
    frmHelp.Show , Me
End Sub

Private Sub mnuTroubleshooting_click()
    frmHelp.Show , Me
    frmHelp.lstHelpTopics.ListIndex = 2
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuLicence_Click()
    frmLicence.Show vbModal, Me
End Sub

Private Sub mnuAcknowldegements_Click()
    frmAcknowledgements.Show vbModal, Me
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' OPTIONS MENU
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub mnuOptShowPoints_Click()
    Call optShowPoints
End Sub

Private Sub mnuOptShowCrosshairs_Click()
    Call optShowCrosshairs
End Sub

Private Sub mnuOptFont_Click()
    Call optFont
End Sub

Private Sub mnuOptShowDebugMenu_Click()
    Call optShowDebugMenu
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DROPDOWN
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub cmbFuncs_Click()
    Call changeDropdown
End Sub

Private Sub txtSelectedFunc_GotFocus()
    selectAll txtSelectedFunc
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CROSSHAIRS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Draw crosshairs
Private Sub pctCrosshairs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call drawCrosshairs(X, Y)
End Sub

' Remove crosshairs when mouse over form
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call removeCrosshairs
End Sub

' Remove crosshairs when lost focus
Private Sub pctCrosshairs_LostFocus()
    Call removeCrosshairs
End Sub

' Save crosshairs
Private Sub pctCrosshairs_Click()
    Call saveCrosshairs
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DEBUG
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub mnuFuncList_Click()
    frmFuncList.Show
End Sub

Private Sub mnuTestEvaluator_Click()
    Call testEvaluator
End Sub

Private Sub mnuTestPostfix_Click()
    Call testPostfix
End Sub
