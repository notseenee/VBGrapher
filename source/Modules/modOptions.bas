Attribute VB_Name = "modOptions"
Option Explicit

Public Sub optShowPoints()
    If frmMain.mnuOptShowPoints.Checked = False Then
        frmMain.mnuOptShowPoints.Checked = True
    Else
        frmMain.mnuOptShowPoints.Checked = False
    End If
End Sub

Public Sub optShowCrosshairs()
    If frmMain.mnuOptShowCrosshairs.Checked = False Then
        frmMain.mnuOptShowCrosshairs.Checked = True
    Else
        frmMain.mnuOptShowCrosshairs.Checked = False
    End If
End Sub

Public Sub optFont()
    ' Enable cancel
    frmMain.CommonDialogFont.CancelError = True
    On Error GoTo Cancel:

    ' Show both screen and printer fonts
    frmMain.CommonDialogFont.Flags = cdlCFBoth
    
    ' Set to current fonts
    frmMain.CommonDialogFont.FontName = frmMain.pctGraph.FontName
    frmMain.CommonDialogFont.FontSize = frmMain.pctGraph.FontSize
    
    ' Show dialog
    frmMain.CommonDialogFont.ShowFont
    
    ' Update pictureboxes fonts
    frmMain.pctGraph.FontName = frmMain.CommonDialogFont.FontName
    frmMain.pctGraph.FontSize = frmMain.CommonDialogFont.FontSize
    
    frmMain.pctCrosshairs.FontName = frmMain.CommonDialogFont.FontName
    frmMain.pctCrosshairs.FontSize = frmMain.CommonDialogFont.FontSize
    
    ' Reset graph
    clearGraph True
    
    Exit Sub
    
Cancel:
    Exit Sub
    
End Sub

Public Sub optShowDebugMenu()
    
    If frmMain.mnuOptShowDebugMenu.Checked = False Then
        ' Warn
        MsgBox "The debug menu is for development purposes only and may cause program crashes.", _
            vbOKOnly & vbExclamation, _
            "Show Debug Menu"
        
        frmMain.mnuOptShowDebugMenu.Checked = True
        frmMain.mnuDebug.Visible = True
    Else
        frmMain.mnuOptShowDebugMenu.Checked = False
        frmMain.mnuDebug.Visible = False
    End If
    
End Sub
