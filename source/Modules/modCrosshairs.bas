Attribute VB_Name = "modCrosshairs"
Option Explicit

Public Sub resetCrosshairs()
    ' Copy from main picturebox
    frmMain.pctCrosshairs.PaintPicture frmMain.pctGraph.Image, -100, -100
End Sub

Public Sub drawCrosshairs(X As Single, Y As Single)
    Call resetCrosshairs
    
    ' status bar coordinates
    frmMain.lblXVal.Caption = Round(Int(X) / 10, 1)
    frmMain.lblYVal.Caption = Round(Y / -10, 2)
    frmMain.lblFXVal.Caption = ""
    
    Dim intCurrX As Integer
    Dim intCurrElem As Integer
    Dim decCurrFX As Variant
    Dim decDomainMin As Variant
    Dim decDomainMax As Variant
    Dim strCoordinates As Variant
    
    ' Get index of current function selected
    intSelectedFunc = frmMain.cmbFuncs.ListIndex
    
    ' if not the blank function
    If intSelectedFunc > 0 Then
    
        ' get domains
        decDomainMin = funcList(intSelectedFunc).decDomainMin
        decDomainMax = funcList(intSelectedFunc).decDomainMax
        
        ' get current element of yval array
        intCurrX = Int(X + decDomainMin * -10)
        
        ' Only if in domain
        If intCurrX > -1 Then
        
            ' If there is no f(x) value
            On Error GoTo NoValue
        
            ' get current f(x) value
            decCurrFX = funcList(intSelectedFunc).decyValues(intCurrX)
            frmMain.lblFXVal.Caption = Round(decCurrFX, 4)
            
            ' Draw lines to axes
            If frmMain.mnuOptShowCrosshairs.Checked = True Then
                frmMain.pctCrosshairs.DrawStyle = vbDot
                frmMain.pctCrosshairs.DrawWidth = 1
                frmMain.pctCrosshairs.Line (Int(X), 0)-(Int(X), decCurrFX * -10), &H808080
                frmMain.pctCrosshairs.Line (0, decCurrFX * -10)-(X, decCurrFX * -10), &H808080
            End If
            
            If frmMain.mnuOptShowPoints.Checked = True Then
                ' Get coordinates string
                strCoordinates = "(" & Round(Int(X) / 10, 1) & ", " & Round(decCurrFX, 4) & ")"

                ' Add white box around text
                frmMain.pctCrosshairs.DrawStyle = vbFSSolid
                frmMain.pctCrosshairs.Line (Int(X) + 1, _
                    decCurrFX * -10 + 1)- _
                    (Int(X) + frmMain.pctCrosshairs.TextWidth(strCoordinates) + 1, _
                    decCurrFX * -10 + frmMain.pctCrosshairs.TextHeight(strCoordinates) + 1), _
                    vbWhite, BF

                ' Draw text
                frmMain.pctCrosshairs.CurrentX = Int(X) + 1
                frmMain.pctCrosshairs.CurrentY = decCurrFX * -10 + 1
                frmMain.pctCrosshairs.FontBold = True
                frmMain.pctCrosshairs.Print strCoordinates

                ' Draw point
                frmMain.pctCrosshairs.DrawStyle = vbSolid
                frmMain.pctCrosshairs.DrawWidth = 2
                frmMain.pctCrosshairs.FillStyle = vbFSSolid
                frmMain.pctCrosshairs.FillColor = funcList(intSelectedFunc).lngColour
                frmMain.pctCrosshairs.Circle (Int(X), decCurrFX * -10), 3, &HFFFFFF
            End If
            
        Else
            
            frmMain.lblFXVal.Caption = ""
            
        End If
            
    End If
    
    Exit Sub
    
NoValue:
    frmMain.lblFXVal.Caption = ""
    Exit Sub
    
End Sub

Public Sub removeCrosshairs()
    Call resetCrosshairs
    
    ' Reset labels
    frmMain.lblXVal.Caption = ""
    frmMain.lblYVal.Caption = ""
    frmMain.lblFXVal.Caption = ""
End Sub

Public Sub saveCrosshairs()
    frmMain.pctGraph.PaintPicture frmMain.pctCrosshairs.Image, -100, -100
End Sub
