Attribute VB_Name = "modDropdown"
Option Explicit
Public intSelectedFunc As Integer

' Update functions dropdown
Public Sub updateDropdown()

    Dim intFuncIterate As Integer
        
    ' clear listbox
    frmMain.cmbFuncs.Clear
    
    If UBound(funcList) > 0 Then
        frmMain.cmbFuncs.Enabled = True
        
        For intFuncIterate = LBound(funcList) To UBound(funcList)
            frmMain.cmbFuncs.AddItem funcList(intFuncIterate).strFuncInput, intFuncIterate
        Next intFuncIterate
        
        ' Retain selected function
        frmMain.cmbFuncs.ListIndex = intSelectedFunc
    Else
        ' Reset textbox
        frmMain.txtSelectedFunc.Text = "Add a function below"
        frmMain.txtSelectedFunc.ForeColor = vbBlack
        ' disable combobox
        frmMain.cmbFuncs.Enabled = False
        ' Disable buttons
        frmMain.cmdRemoveFunc.Enabled = False
        frmMain.mnuRemoveFunc.Enabled = False
        frmMain.cmdEditColour.Enabled = False
        frmMain.mnuEditColour.Enabled = False
    End If
    
End Sub

' On dropdown change
Public Sub changeDropdown()
    ' Reset textbox
    frmMain.txtSelectedFunc.Text = "Select a function"
    frmMain.txtSelectedFunc.ForeColor = vbBlack
    
    ' Disable buttons
    frmMain.cmdRemoveFunc.Enabled = False
    frmMain.mnuRemoveFunc.Enabled = False
    frmMain.cmdEditColour.Enabled = False
    frmMain.mnuEditColour.Enabled = False
    
    If frmMain.cmbFuncs.ListIndex > 0 Then
        ' Update intSelectedFunc
        intSelectedFunc = frmMain.cmbFuncs.ListIndex
        
        ' Write to textbox
        frmMain.txtSelectedFunc.Text = frmMain.cmbFuncs.List(frmMain.cmbFuncs.ListIndex)
        
        ' Enable buttons, if not a removed or errored function
        If InStr(1, frmMain.txtSelectedFunc.Text, "Removed: ") <> 1 And _
        InStr(1, frmMain.txtSelectedFunc.Text, "Error: ") <> 1 Then
            frmMain.cmdRemoveFunc.Enabled = True
            frmMain.mnuRemoveFunc.Enabled = True
            frmMain.cmdEditColour.Enabled = True
            frmMain.mnuEditColour.Enabled = True
        End If
        
        ' Show colour
        frmMain.txtSelectedFunc.ForeColor = funcList(frmMain.cmbFuncs.ListIndex).lngColour
    End If
    
End Sub

' Select all in a textbox
Public Sub selectAll(textbox As textbox)
    textbox.SelStart = 0
    textbox.SelLength = Len(textbox)
End Sub
