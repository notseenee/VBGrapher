VERSION 5.00
Begin VB.Form frmFuncList 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Debug Output: Function List"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10860
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFuncList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4305
   ScaleWidth      =   10860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   9180
      TabIndex        =   1
      Top             =   3735
      Width           =   1575
   End
   Begin VB.CommandButton cmdLoadFuncs 
      Caption         =   "&Load Functions"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   7560
      TabIndex        =   0
      Top             =   3735
      Width           =   1575
   End
End
Attribute VB_Name = "frmFuncList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub printFuncs(Optional blnWithYVals As Boolean)
    Cls
    frmFuncList.FontBold = True
    Print "INDEX", "FUNCINPUT", "DOMAINMIN", "DOMAINMAX", "POSTFIX", "YVALS"
    frmFuncList.FontBold = False
    Dim i As Integer
    Dim j As Integer
    For i = LBound(funcList) To UBound(funcList)
        With funcList(i)
            Print i, .strFuncInput, .decDomainMin, .decDomainMax, .strPostfix
            
            If blnWithYVals = True Then
            
'                If i > 0 And UBound(.decyValues) <> 0 Then
'                    For j = 0 To UBound(funcList(i).decyValues)
'                        Print j, funcList(i).decyValues(j)
'                    Next j
'                End If
                
            End If
            
        End With
    Next i
End Sub

Private Sub cmdLoadFuncs_Click()
    printFuncs True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Call printFuncs
End Sub
