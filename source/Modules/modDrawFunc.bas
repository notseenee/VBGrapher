Attribute VB_Name = "modDrawFunc"
Option Explicit

Public Sub drawFunc(intArrayIndex As Integer, _
    strPostfix As String, _
    decDomainMin As Variant, _
    decDomainMax As Variant, _
    lngColour As Long)
    
    On Error GoTo ErrorHandler
    
    ' For getting y-values
    Dim decCurrX As Variant
    Dim decyValues() As Variant
    
    ' For plotting
    Dim intCurrElem As Integer
    Dim intCurrX As Integer
    Dim decCurrY As Variant
    Dim decNextX As Variant
    Dim decNextY As Variant
    
    ' Initialise array
    ReDim decyValues(0)
    
    ' Make y value array
    For decCurrX = decDomainMin To decDomainMax Step 0.1
        
        ' Resize decyValues array
        ReDim Preserve decyValues(UBound(decyValues) + 1)
        
        ' Write to array
        decyValues(UBound(decyValues) - 1) = evaluate(strPostfix, decCurrX)
    Next decCurrX
    
    ' Sync arrays
    funcList(intArrayIndex).decyValues = decyValues
    
    ' Draw
    frmMain.pctGraph.DrawWidth = 2
    ' Start at minimum
    decCurrX = decDomainMin
    For intCurrElem = 0 To UBound(decyValues) - 2
        ' Get current Y
        decCurrY = decyValues(intCurrElem)
        
        ' get next X & Y
        decNextX = decCurrX + CDec(0.1)
        decNextY = decyValues(intCurrElem + 1)
        
        ' Draw
        If decCurrY <> "" And decNextY <> "" Then
            frmMain.pctGraph.Line (decCurrX * 10, decCurrY * -10)-(decNextX * 10, decNextY * -10), lngColour
        End If
        
        ' Increment
        decCurrX = decCurrX + 0.1
    Next intCurrElem
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
    
End Sub
