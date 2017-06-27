Attribute VB_Name = "modDebug"
Option Explicit

Public Sub testEvaluator()
    On Error GoTo ErrorHandler

    Dim strTestInput As String
    Dim strPostfix As String
    Dim strValues As String
    Dim X As Variant
    strTestInput = InputBox("infix function:", "Debug Input: Test Evaluator")
    strPostfix = toPostfix(strTestInput)
    
    For X = CDec(-1) To CDec(1) Step CDec(0.1)
        strValues = strValues & vbCrLf & X & " : " & evaluate(strPostfix, X)
    Next X
    
    MsgBox "infix: " & strTestInput & vbNewLine & "postfix: " & strPostfix & vbNewLine _
        & "result: " & strValues, _
        vbOKOnly, _
        "Debug Input: Test Evaluator"
        
ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox "There was an error with your function." & vbCrLf & vbCrLf & _
            "Error Number " & Err.Number & vbCrLf & _
            Err.Description, _
            vbOKOnly & vbCritical, _
            "Function Input Error"
    End If
End Sub

Public Sub testPostfix()
    On Error GoTo ErrorHandler

    Dim strTestInput As String
    strTestInput = InputBox("infix function:", "Debug Input: Test Postfix Converter")
    MsgBox "infix: " + strTestInput + vbNewLine + "postfix: " + toPostfix(strTestInput), _
        vbOKOnly, _
        "Debug Output: Test Postfix Converter"
        
ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox "There was an error with your function." & vbCrLf & vbCrLf & _
            "Error Number " & Err.Number & vbCrLf & _
            Err.Description, _
            vbOKOnly & vbCritical, _
            "Function Input Error"
    End If
End Sub
