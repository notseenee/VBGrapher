Attribute VB_Name = "modGraph"
Option Explicit

Public Sub clearGraph(Optional blnClear As Boolean)

    Dim i As Integer
    Dim intFuncs As Integer
    Dim lngColour As Long
    
    If blnClear = True Then
        frmMain.pctGraph.BackColor = vbWhite
        frmMain.pctGraph.Cls
        ' Remove all functions from array
        ReDim funcList(0)
        ' Reset dropdown
        Call updateDropdown
    End If
    
    frmMain.pctGraph.DrawWidth = 1
    
    For i = -10 To 9
    
        ' Draw gridlines
        frmMain.pctGraph.Line (i * 10, -100)-(i * 10, 100), &HDDDDDD
        frmMain.pctGraph.Line (-100, i * 10)-(100, i * 10), &HDDDDDD
    
    Next i
    
    ' Draw last gridlines
    frmMain.pctGraph.Line (99.5, -100)-(99.5, 100), &HDDDDDD
    frmMain.pctGraph.Line (-100, 99.5)-(100, 99.5), &HDDDDDD
    
    ' Add axes
    frmMain.pctGraph.Line (0, -100)-(0, 100), vbBlack
    frmMain.pctGraph.Line (-100, 0)-(100, 0), vbBlack
    
    For i = -8 To 8 Step 2
        
        ' Add coordinates
        If i <> 0 Then
            ' Add x coordinates
            frmMain.pctGraph.CurrentY = 1
            frmMain.pctGraph.CurrentX = i * 10 - 2.5
            frmMain.pctGraph.Print i
            
            ' Add y coordinates
            ' Fix for negative numbers
            If i > 0 Then
                frmMain.pctGraph.CurrentX = -6.5
            Else
                frmMain.pctGraph.CurrentX = -5.5
            End If
            
            frmMain.pctGraph.CurrentY = i * 10 - 3
            frmMain.pctGraph.Print i * -1
        End If
        
    Next i
    
    ' Clear crosshairs pbox
    Call resetCrosshairs
    
End Sub

Public Sub removeFunc()
    ' Redraw as white
    Call drawFunc(intSelectedFunc, _
        funcList(intSelectedFunc).strPostfix, _
        funcList(intSelectedFunc).decDomainMin, _
        funcList(intSelectedFunc).decDomainMax, _
        vbWhite)
        
    ' sync crosshairs pbox
    Call resetCrosshairs
    
    ' mark as removed
    funcList(intSelectedFunc).strFuncInput = "Removed: " & funcList(intSelectedFunc).strFuncInput
    
    ' Remove yval array
    ReDim funcList(intSelectedFunc).decyValues(0)
    
    ' set colour to white
    funcList(intSelectedFunc).lngColour = vbRed
    
    ' update dropdown
    Call updateDropdown
End Sub

Public Sub saveGraph()
    ' Set Cancel to True
    frmMain.CommonDialogSave.CancelError = True
    On Error GoTo ErrHandler
    
    frmMain.CommonDialogSave.ShowSave
    
    SavePicture frmMain.pctCrosshairs.Image, frmMain.CommonDialogSave.FileName
    
    Exit Sub
    
ErrHandler:
    Exit Sub
    
End Sub

Public Sub editColour()
    ' Set Cancel to True
    frmMain.CommonDialogColor.CancelError = True
    On Error GoTo ErrHandler
    'Set the Flags property
    frmMain.CommonDialogColor.Flags = cdlCCRGBInit
    ' Set colour to current
    frmMain.CommonDialogColor.Color = frmMain.txtSelectedFunc.ForeColor
    ' Display the Color Dialog box
    frmMain.CommonDialogColor.ShowColor
    ' Show colour
    frmMain.txtSelectedFunc.ForeColor = frmMain.CommonDialogColor.Color
    ' Set colour in funclist
    funcList(intSelectedFunc).lngColour = frmMain.txtSelectedFunc.ForeColor
    
    ' Change cursor
    Screen.MousePointer = vbHourglass
    ' Redraw
    Call drawFunc(intSelectedFunc, _
        funcList(intSelectedFunc).strPostfix, _
        funcList(intSelectedFunc).decDomainMin, _
        funcList(intSelectedFunc).decDomainMax, _
        frmMain.txtSelectedFunc.ForeColor)
    ' Reset cursor
    Screen.MousePointer = vbDefault
    Exit Sub

ErrHandler:
    ' User pressed the Cancel button
    Exit Sub
End Sub
