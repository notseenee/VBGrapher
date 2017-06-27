VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmNewFunc 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Function"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4755
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
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   4755
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraDomain 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Domain"
      Height          =   915
      Left            =   225
      TabIndex        =   6
      Top             =   2790
      Width           =   3300
      Begin VB.TextBox txtDomainMax 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1485
         TabIndex        =   9
         Text            =   "10"
         Top             =   360
         Width           =   645
      End
      Begin VB.TextBox txtDomainMin 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   135
         TabIndex        =   8
         Text            =   "-10"
         Top             =   360
         Width           =   645
      End
      Begin VB.CommandButton cmdDomainDefault 
         Caption         =   "&Default"
         Height          =   375
         Left            =   2205
         TabIndex        =   7
         Top             =   315
         Width           =   915
      End
      Begin VB.Label lblDomainX 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "<  x  <"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   855
         TabIndex        =   10
         Top             =   370
         Width           =   600
      End
   End
   Begin VB.Frame fraColour 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Colour"
      Height          =   915
      Left            =   225
      TabIndex        =   5
      Top             =   1755
      Width           =   4335
      Begin MSComDlg.CommonDialog CommonDialogColor 
         Left            =   3735
         Top             =   270
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Orientation     =   2
      End
      Begin VB.CommandButton cmdColour 
         Appearance      =   0  'Flat
         Caption         =   "Change colour"
         Height          =   450
         Left            =   675
         MaskColor       =   &H8000000F&
         TabIndex        =   1
         Top             =   315
         Width           =   2625
      End
      Begin VB.Label lblColourLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Colour"
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
         Height          =   255
         Left            =   180
         TabIndex        =   14
         Top             =   45
         Width           =   585
      End
      Begin VB.Label lblColour 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   180
         TabIndex        =   13
         Top             =   360
         Width           =   390
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   450
      Left            =   2475
      TabIndex        =   3
      Top             =   3915
      Width           =   2130
   End
   Begin VB.CommandButton cmdNewFunc 
      Caption         =   "&New Function"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   450
      Left            =   360
      TabIndex        =   2
      Top             =   3915
      Width           =   2130
   End
   Begin VB.Frame fraEnterFunc 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Enter function"
      Height          =   915
      Left            =   225
      TabIndex        =   4
      Top             =   720
      Width           =   4335
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
         Height          =   405
         Left            =   900
         TabIndex        =   0
         Text            =   "test"
         Top             =   315
         Width           =   3225
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   180
         Picture         =   "frmNewFunc.frx":0000
         Top             =   420
         Width           =   480
      End
      Begin VB.Label lblEnterFunc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Enter Function"
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
         Height          =   255
         Left            =   130
         TabIndex        =   11
         Top             =   45
         Width           =   1230
      End
   End
   Begin VB.Label lblNewFunc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "New Function"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   345
      TabIndex        =   12
      Top             =   120
      Width           =   2370
   End
End
Attribute VB_Name = "frmNewFunc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim lngColourArray(8) As Long
    Dim i As Long
    
    ' Assign to colour array
    lngColourArray(0) = RGB(244, 67, 54)
    lngColourArray(1) = RGB(233, 30, 99)
    lngColourArray(2) = RGB(156, 39, 176)
    lngColourArray(3) = RGB(63, 81, 181)
    lngColourArray(4) = RGB(33, 150, 243)
    lngColourArray(5) = RGB(0, 150, 136)
    lngColourArray(6) = RGB(76, 175, 80)
    lngColourArray(7) = RGB(205, 220, 57)
    lngColourArray(8) = RGB(255, 152, 0)
    
    ' Choose random colour
    Randomize
    lblColour.BackColor = lngColourArray(Int(9 * Rnd))
    
End Sub

Private Sub cmdDomainDefault_Click()
    txtDomainMin.Text = "-10"
    txtDomainMax.Text = "10"
End Sub

Private Sub cmdNewFunc_Click()
    Dim lngColour As Long
    Dim decDomainMin As Variant
    Dim decDomainMax As Variant
    Dim strPostfix As String
    
    ' Change cursor
    Screen.MousePointer = vbHourglass
    
    ' Get values from form
    lngColour = CLng(lblColour.BackColor)
    decDomainMin = CDec(txtDomainMin.Text)
    decDomainMax = CDec(txtDomainMax.Text)

    ' Extend funcList array
    ReDim Preserve funcList(UBound(funcList) + 1)
    
    ' Write to funcList array
    With funcList(UBound(funcList) - 1)
        .strFuncInput = txtFuncInput.Text
        .lngColour = lngColour
        .decDomainMin = decDomainMin
        .decDomainMax = decDomainMax
    End With
    
    ' Parse equation into postfix
    strPostfix = toPostfix(txtFuncInput.Text)
    funcList(UBound(funcList) - 1).strPostfix = strPostfix
    
    ' Draw function
    drawFunc UBound(funcList) - 1, strPostfix, decDomainMin, decDomainMax, lngColour
    
    ' Reset cursor
    Screen.MousePointer = vbDefault
    
    ' Add to list
    frmMain.cmbFunctions.AddItem txtFuncInput.Text, UBound(funcList) - 1
    
    ' Close window
    Unload Me
    
    ' Show main window
    frmMain.Show
End Sub

Private Sub cmdCancel_Click()
    Dim cancelDialogue As Integer
    cancelDialogue = MsgBox("Are you sure you want to cancel adding a new function?", _
        vbYesNo + vbQuestion, _
        "Cancel New Function")
    
    If cancelDialogue = vbYes Then
        Unload Me
        frmMain.Show
    End If
    
End Sub

Private Sub cmdColour_Click()
    ' Set Cancel to True
    CommonDialogColor.CancelError = True
    On Error GoTo ErrHandler
    'Set the Flags property
    CommonDialogColor.Flags = cdlCCRGBInit
    ' Display the Color Dialog box
    CommonDialogColor.ShowColor
    ' Set the label's background color to selected color
    lblColour.BackColor = CommonDialogColor.Color
    Exit Sub

ErrHandler:
    ' User pressed the Cancel button
    Exit Sub
End Sub

Private Sub txtFuncInput_KeyPress(KeyAscii As Integer)
    cmdNewFunc.Enabled = True
End Sub


Private Sub txtFuncInput_Validate(Cancel As Boolean)
    If txtFuncInput.Text = "" Or txtFuncInput.Text = "This field must not be blank." Then
        txtFuncInput.Text = "This field must not be blank."
        txtFuncInput.BackColor = vbRed
        txtFuncInput.ForeColor = vbWhite
    Else
        cmdNewFunc.Enabled = True
    End If
End Sub

Private Sub txtDomainMin_Validate(Cancel As Boolean)
    If txtDomainMin.Text = "" Then
        txtDomainMin.Text = "-10"
    End If
End Sub

Private Sub txtDomainMax_Validate(Cancel As Boolean)
    If txtDomainMax.Text = "" Then
        txtDomainMax.Text = "10"
    End If
End Sub

' Selec tall in a textbox
Private Sub selectAll(textbox As textbox)
    textbox.SelStart = 0
    textbox.SelLength = Len(textbox)
End Sub

Private Sub txtFuncInput_GotFocus()
    selectAll txtFuncInput
    txtFuncInput.BackColor = vbWhite
    txtFuncInput.ForeColor = vbBlack
End Sub
Private Sub txtDomainMin_GotFocus()
    selectAll txtDomainMin
End Sub
Private Sub txtDomainMax_GotFocus()
    selectAll txtDomainMax
End Sub
