VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help"
   ClientHeight    =   18060
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16905
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1204
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1127
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pctHelpTopic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5955
      Index           =   5
      Left            =   9540
      ScaleHeight     =   397
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   11
      Top             =   12075
      Visible         =   0   'False
      Width           =   7200
      Begin VB.CommandButton cmdSaveGraph 
         Caption         =   "Save"
         Height          =   720
         Left            =   0
         MaskColor       =   &H00F0F0F0&
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Save the graph to a file"
         Top             =   870
         UseMaskColor    =   -1  'True
         Width           =   780
      End
      Begin VB.TextBox txt5 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4035
         Left            =   0
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         TabIndex        =   43
         Text            =   "frmHelp.frx":0000
         Top             =   525
         Width           =   7140
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Saving the Graph"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   6960
      End
   End
   Begin VB.PictureBox pctHelpTopic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5955
      Index           =   4
      Left            =   2265
      ScaleHeight     =   397
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   9
      Top             =   12060
      Visible         =   0   'False
      Width           =   7200
      Begin VB.CommandButton cmdClearGraph 
         Caption         =   "Clear"
         Height          =   720
         Left            =   0
         MaskColor       =   &H00F0F0F0&
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Remove all functions from the graph"
         Top             =   870
         UseMaskColor    =   -1  'True
         Width           =   780
      End
      Begin VB.TextBox txt4 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4035
         Left            =   0
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         TabIndex        =   41
         Text            =   "frmHelp.frx":012B
         Top             =   525
         Width           =   7140
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Clearing the Graph"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   6960
      End
   End
   Begin VB.PictureBox pctHelpTopic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5955
      Index           =   3
      Left            =   9480
      ScaleHeight     =   397
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   6045
      Visible         =   0   'False
      Width           =   7200
      Begin VB.TextBox txt3Link1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   0
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   40
         Text            =   "clear the graph"
         Top             =   4860
         Width           =   1350
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         ItemData        =   "frmHelp.frx":01D3
         Left            =   0
         List            =   "frmHelp.frx":01E0
         Style           =   2  'Dropdown List
         TabIndex        =   39
         ToolTipText     =   "Choose a different function from the graph"
         Top             =   3975
         Width           =   6015
      End
      Begin VB.CommandButton cmdRemoveFunc 
         Caption         =   "Remove"
         Height          =   720
         Left            =   0
         MaskColor       =   &H00F0F0F0&
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Remove the selected function"
         Top             =   1950
         UseMaskColor    =   -1  'True
         Width           =   780
      End
      Begin VB.ComboBox Combo1 
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
         ItemData        =   "frmHelp.frx":01FD
         Left            =   0
         List            =   "frmHelp.frx":020A
         Style           =   2  'Dropdown List
         TabIndex        =   37
         ToolTipText     =   "Choose a different function from the graph"
         Top             =   885
         Width           =   6015
      End
      Begin VB.TextBox txt3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5385
         Left            =   0
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         TabIndex        =   36
         Text            =   "frmHelp.frx":021E
         Top             =   525
         Width           =   7140
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Removing a Function"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   6960
      End
   End
   Begin VB.PictureBox pctHelpTopic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5955
      Index           =   2
      Left            =   2265
      ScaleHeight     =   397
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   6045
      Visible         =   0   'False
      Width           =   7200
      Begin VB.TextBox txt2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5385
         Left            =   0
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   35
         Text            =   "frmHelp.frx":03DF
         Top             =   525
         Width           =   7200
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Supported Terms"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   6195
      End
   End
   Begin VB.PictureBox pctHelpTopic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5955
      Index           =   1
      Left            =   9480
      ScaleHeight     =   397
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   45
      Visible         =   0   'False
      Width           =   7200
      Begin VB.TextBox txt1Link2 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   0
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   34
         Text            =   "Supported Terms"
         Top             =   5220
         Width           =   1515
      End
      Begin VB.TextBox txt1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1395
         Index           =   2
         Left            =   0
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         TabIndex        =   33
         Text            =   "frmHelp.frx":08A5
         Top             =   4455
         Width           =   7140
      End
      Begin VB.TextBox txtDomainMin 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   "-10"
         Top             =   3840
         Width           =   600
      End
      Begin VB.TextBox txtDomainMax 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "10"
         Top             =   3840
         Width           =   600
      End
      Begin VB.TextBox txt1Link1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   3255
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   30
         Text            =   "Supported Terms"
         Top             =   1290
         Width           =   1515
      End
      Begin VB.TextBox txtColour 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   450
         Left            =   0
         TabIndex        =   29
         Top             =   2385
         Width           =   450
      End
      Begin VB.CommandButton cmdColour 
         Appearance      =   0  'Flat
         Caption         =   "Change &colour"
         Height          =   450
         Left            =   555
         MaskColor       =   &H8000000F&
         TabIndex        =   28
         Top             =   2385
         Width           =   1920
      End
      Begin VB.TextBox txt1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   0
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   27
         Text            =   $"frmHelp.frx":091E
         Top             =   525
         Width           =   7140
      End
      Begin VB.TextBox txtFuncInput 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   540
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   825
         Width           =   3375
      End
      Begin VB.TextBox txt1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2565
         Index           =   1
         Left            =   0
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         TabIndex        =   26
         Text            =   "frmHelp.frx":0952
         Top             =   1290
         Width           =   7140
      End
      Begin VB.Image imgDomain 
         Height          =   240
         Left            =   675
         Top             =   3900
         Width           =   555
      End
      Begin VB.Image imgFX 
         Height          =   240
         Left            =   0
         Top             =   900
         Width           =   480
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Adding a Function"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   6960
      End
   End
   Begin VB.PictureBox pctHelpTopic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5955
      Index           =   0
      Left            =   2250
      ScaleHeight     =   397
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   75
      Width           =   7200
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
         ItemData        =   "frmHelp.frx":0A66
         Left            =   0
         List            =   "frmHelp.frx":0A68
         Style           =   2  'Dropdown List
         TabIndex        =   18
         ToolTipText     =   "Choose a different function from the graph"
         Top             =   3435
         Width           =   6015
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
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "Add a function below"
         ToolTipText     =   "Select the current function to copy and paste"
         Top             =   3465
         Width           =   5625
      End
      Begin VB.TextBox txt0Link1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   5175
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   16
         Text            =   "Adding a Function"
         Top             =   1545
         Width           =   1605
      End
      Begin VB.CommandButton cmdAddFunc 
         Caption         =   "Add"
         Height          =   720
         Left            =   0
         MaskColor       =   &H00F0F0F0&
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Add a new function"
         Top             =   1905
         UseMaskColor    =   -1  'True
         Width           =   780
      End
      Begin VB.TextBox txt0 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4425
         Left            =   0
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         TabIndex        =   14
         Text            =   "frmHelp.frx":0A6A
         Top             =   525
         Width           =   7140
      End
      Begin VB.Label lblX 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "x"
         ForeColor       =   &H00808080&
         Height          =   240
         Left            =   660
         TabIndex        =   24
         Top             =   4920
         Width           =   270
      End
      Begin VB.Label lblY 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "y"
         ForeColor       =   &H00808080&
         Height          =   240
         Left            =   660
         TabIndex        =   23
         Top             =   5175
         Width           =   270
      End
      Begin VB.Label lblFX 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "f(x)"
         ForeColor       =   &H00808080&
         Height          =   240
         Left            =   660
         TabIndex        =   22
         Top             =   5430
         Width           =   270
      End
      Begin VB.Label lblFXVal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "0.8085"
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
         Height          =   225
         Left            =   105
         TabIndex        =   21
         Top             =   5430
         Width           =   495
      End
      Begin VB.Label lblYVal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "-5.6"
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
         Height          =   225
         Left            =   300
         TabIndex        =   20
         Top             =   5175
         Width           =   300
      End
      Begin VB.Label lblXVal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "2.2"
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
         Height          =   225
         Left            =   375
         TabIndex        =   19
         Top             =   4920
         Width           =   225
      End
      Begin VB.Label lblTitle 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Getting Started"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   6960
      End
   End
   Begin VB.ListBox lstHelpTopics 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5385
      ItemData        =   "frmHelp.frx":0C60
      Left            =   120
      List            =   "frmHelp.frx":0C76
      TabIndex        =   0
      Top             =   570
      Width           =   1950
   End
   Begin VB.Label lblTopic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Topic"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   165
      TabIndex        =   13
      Top             =   240
      Width           =   510
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer

Private Sub Form_Load()

    ' Add window icon
    Me.Icon = LoadResPicture(101, vbResIcon)
    
    ' Change background colour
    Me.BackColor = vbWhite
    
    ' Resize
    Me.Width = 9525
    Me.Height = 6570
    
    ' Position all the frames
    For i = 0 To 5
        With pctHelpTopic(i)
            .Left = 150
            .Top = 4
        End With
    Next i
    
    
    ' Add images
    cmdAddFunc.Picture = LoadResPicture(101, vbResBitmap)
    
    imgFX.Picture = LoadResPicture(104, vbResBitmap)
    imgDomain.Picture = LoadResPicture(108, vbResBitmap)
    
    cmdRemoveFunc.Picture = LoadResPicture(106, vbResBitmap)
    Combo1.ListIndex = 1
    Combo2.ListIndex = 1
    
    cmdClearGraph.Picture = LoadResPicture(102, vbResBitmap)
    
    cmdSaveGraph.Picture = LoadResPicture(107, vbResBitmap)
    
    ' Select first topic
    lstHelpTopics.ListIndex = 0
    
End Sub

Private Sub lstHelpTopics_Click()
    ' Hide all
    For i = 0 To 5
        pctHelpTopic(i).Visible = False
    Next i

    ' Show selected
    pctHelpTopic(lstHelpTopics.ListIndex).Visible = True
End Sub

Private Sub txt0Link1_Click()
    lstHelpTopics.ListIndex = 1
End Sub
Private Sub txt1Link1_Click()
    lstHelpTopics.ListIndex = 2
End Sub
Private Sub txt1Link2_Click()
    lstHelpTopics.ListIndex = 2
End Sub
