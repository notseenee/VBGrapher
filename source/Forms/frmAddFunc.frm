VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddFunc 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Function"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4590
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
   ScaleHeight     =   4590
   ScaleWidth      =   4590
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdHelp 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3810
      TabIndex        =   15
      Top             =   150
      Width           =   450
   End
   Begin VB.Frame fraDomain 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Domain"
      Height          =   810
      Left            =   225
      TabIndex        =   9
      Top             =   2595
      Width           =   4320
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
         Left            =   1440
         TabIndex        =   3
         Text            =   "10"
         Top             =   330
         Width           =   600
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
         Left            =   120
         TabIndex        =   2
         Text            =   "-10"
         Top             =   330
         Width           =   600
      End
      Begin VB.CommandButton cmdDomainDefault 
         Caption         =   "&Default domain"
         Height          =   450
         Left            =   2160
         TabIndex        =   4
         Top             =   270
         Width           =   1920
      End
      Begin VB.Label lblDomain 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Domain"
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
         Left            =   105
         TabIndex        =   14
         Top             =   0
         Width           =   735
      End
      Begin VB.Image imgDomain 
         Height          =   240
         Left            =   800
         Top             =   390
         Width           =   555
      End
   End
   Begin VB.Frame fraColour 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Colour"
      Height          =   915
      Left            =   225
      TabIndex        =   8
      Top             =   1605
      Width           =   4335
      Begin MSComDlg.CommonDialog CommonDialogColor 
         Left            =   3555
         Top             =   285
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Orientation     =   2
      End
      Begin VB.CommandButton cmdColour 
         Appearance      =   0  'Flat
         Caption         =   "Change &colour"
         Height          =   450
         Left            =   675
         MaskColor       =   &H8000000F&
         TabIndex        =   1
         Top             =   330
         Width           =   1920
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   105
         TabIndex        =   13
         Top             =   0
         Width           =   615
      End
      Begin VB.Label lblColour 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   120
         TabIndex        =   12
         Top             =   330
         Width           =   450
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   450
      Left            =   2385
      TabIndex        =   6
      Top             =   3810
      Width           =   1920
   End
   Begin VB.CommandButton cmdAddFunc 
      Caption         =   "&Add Function"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   450
      Left            =   345
      TabIndex        =   5
      Top             =   3810
      Width           =   1920
   End
   Begin VB.Frame fraEnterFunc 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Enter function"
      Height          =   915
      Left            =   225
      TabIndex        =   7
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
         Height          =   345
         Left            =   660
         TabIndex        =   0
         Top             =   330
         Width           =   3375
      End
      Begin VB.Image imgFX 
         Height          =   240
         Left            =   120
         Top             =   405
         Width           =   480
      End
      Begin VB.Label lblFunc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Function"
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
         Left            =   105
         TabIndex        =   10
         Top             =   0
         Width           =   810
      End
   End
   Begin VB.Label lblAddFunc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Add Function"
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
      TabIndex        =   11
      Top             =   120
      Width           =   2310
   End
End
Attribute VB_Name = "frmAddFunc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnWarnOnlyOnce() As Boolean

Private Sub Form_Load()
    Dim lngColourArray(8) As Long
    
    ReDim blnWarnOnlyOnce(5)
    
    ' Add window icon
    Me.Icon = LoadResPicture(101, vbResIcon)
    
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
    lblColour.BackColor = lngColourArray(Int(9 * Rnd))
    
    ' Add icons
    imgFX.Picture = LoadResPicture(104, vbResBitmap)
    imgDomain.Picture = LoadResPicture(108, vbResBitmap)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ADD FUNCTION
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAddFunc_Click()
    Dim lngColour As Long
    Dim decDomainMin As Variant
    Dim decDomainMax As Variant
    Dim strPostfix As String
    
    On Error GoTo ErrorHandler
    
    ' Validate input
    If validateInput = False Then
        Me.Enabled = False
        Exit Sub
    End If
    
    ' Change cursor
    Screen.MousePointer = vbHourglass
    
    ' Get values from form
    lngColour = CLng(lblColour.BackColor)
    decDomainMin = CDec(txtDomainMin.Text)
    decDomainMax = CDec(txtDomainMax.Text)
    
    ' Extend funcList array
    ReDim Preserve funcList(UBound(funcList) + 1)
    
    ' Write to funcList array
    With funcList(UBound(funcList))
        .strFuncInput = txtFuncInput.Text
        .lngColour = lngColour
        .decDomainMin = decDomainMin
        .decDomainMax = decDomainMax
    End With
    
    ' Parse equation into postfix
    strPostfix = toPostfix(txtFuncInput.Text)
    
    ' Write to funcList array
    funcList(UBound(funcList)).strPostfix = strPostfix
    
    ' Draw function
    drawFunc UBound(funcList), strPostfix, decDomainMin, decDomainMax, lngColour
    
    ' Reset cursor
    Screen.MousePointer = vbDefault
    
    ' Select function in list
    intSelectedFunc = UBound(funcList)
    
    ' Close window
    Unload Me
    
    ' Show main window
    frmMain.Show
    
    Exit Sub
    
ErrorHandler:
    MsgBox "There was an error with your function." & vbCrLf & vbCrLf & _
        "Error Number " & Err.Number & vbCrLf & _
        Err.Description, _
        vbOKOnly & vbCritical, _
        "Function Input Error"
        
    ' Mark as error
    With funcList(UBound(funcList))
        .strFuncInput = "Error: " & .strFuncInput
        .lngColour = vbRed
    End With
        
    ' Reset cursor
    Screen.MousePointer = vbDefault
        
    Exit Sub
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DISPLAY WARNING
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub displayWarning(strMessage As String, intWarningNumber As Integer)

    If blnWarnOnlyOnce(intWarningNumber) <> True Then
        blnWarnOnlyOnce(intWarningNumber) = True
        
        MsgBox "There was an issue with your function." & vbCrLf & _
            "A portion of it was not graphed." & vbCrLf & vbCrLf & _
            strMessage, _
            vbOKOnly & vbInformation, _
            "Function Input Warning"
    End If
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CANCEL
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
    Dim cancelDialogue As Integer
    cancelDialogue = MsgBox("Are you sure you want to cancel adding a new function?", _
        vbYesNo + vbQuestion, _
        "Cancel Add Function")
    
    If cancelDialogue = vbYes Then
        Unload Me
        frmMain.SetFocus
    End If
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' FUNCTION INPUT
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Validate input
Private Function validateInput() As Boolean

    If txtFuncInput.Text = "" Then
        lblFunc.Caption = "Function field must not be blank."
        lblFunc.ForeColor = vbRed
        
        Beep
        validateInput = False
        
    ElseIf InStr(1, txtFuncInput.Text, "Removed: ") = 1 Then
        lblFunc.Caption = "Function must not begin with ""Removed: """
        lblFunc.ForeColor = vbRed
        
        Beep
        validateInput = False
        
    ElseIf InStr(1, txtFuncInput.Text, "Error: ") = 1 Then
        lblFunc.Caption = "Function must not begin with ""Error: """
        lblFunc.ForeColor = vbRed
        
        Beep
        validateInput = False
        
    Else
        lblFunc.Caption = "Function"
        lblFunc.ForeColor = &H808080
        
        validateInput = True
    End If
        
End Function

Private Sub txtFuncInput_Validate(Cancel As Boolean)
    cmdAddFunc.Enabled = False
    
    If validateInput = True Then cmdAddFunc.Enabled = True
End Sub

' Select all and disable button when textbox is clicked to force validation
Private Sub txtFuncInput_GotFocus()
    selectAll txtFuncInput
End Sub

' Enable button when clicked for fast adding
Private Sub txtFuncInput_KeyPress(KeyAscii As Integer)
    cmdAddFunc.Enabled = True
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CHANGE COLOUR
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub changeColour()
    ' Set Cancel to True
    CommonDialogColor.CancelError = True
    On Error GoTo ErrHandler
    'Set the Flags property
    CommonDialogColor.Flags = cdlCCRGBInit
    ' Set colour to current
    CommonDialogColor.Color = lblColour.BackColor
    ' Display the Color Dialog box
    CommonDialogColor.ShowColor
    ' Set the label's background color to selected color
    lblColour.BackColor = CommonDialogColor.Color
    Exit Sub

ErrHandler:
    ' User pressed the Cancel button
    Exit Sub
End Sub

Private Sub cmdColour_Click()
    Call changeColour
End Sub

Private Sub lblColour_Click()
    Call changeColour
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DOMAIN
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' When blank, set to default values
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

' When clicked, select all
Private Sub txtDomainMin_GotFocus()
    selectAll txtDomainMin
End Sub
Private Sub txtDomainMax_GotFocus()
    selectAll txtDomainMax
End Sub

' Reset to default values
Private Sub cmdDomainDefault_Click()
    txtDomainMin.Text = "-10"
    txtDomainMax.Text = "10"
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' HELP
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdHelp_click()
    frmHelp.Show
    frmHelp.lstHelpTopics.ListIndex = 1
End Sub
