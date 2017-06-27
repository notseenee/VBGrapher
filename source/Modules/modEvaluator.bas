Attribute VB_Name = "modEvaluator"
Option Explicit
Dim strStack() As String
Dim strOutput  As String
Const pi = 3.1415926535898
Const e = 2.718281828459

Public Function evaluate(strPostfix As String, decCurrX As Variant) As Variant
    Dim strTokens()  As String
    Dim intToken     As Integer
    Dim strCurrToken As String
    
    Dim decOperatorL As Variant
    Dim decOperatorR As Variant
    Dim decResult    As Variant
    
    On Error GoTo ErrorHandler
    
    ' Skip to next
    Dim blnSkip As Integer
    blnSkip = False
    
    ' Tokenise input string
    strTokens = Split(strPostfix)
    
    ' Initialise & clear stack array
    ReDim strStack(0)
    
    ' Clear output string
    strOutput = ""
    
    ' Loop through tokens
    For intToken = 0 To UBound(strTokens)
        strCurrToken = strTokens(intToken)
        
        ' If blank or a space, ignore
        If strCurrToken = " " Or strCurrToken = "" Then
        ' If token is an operand
        ElseIf isOperator(strCurrToken) = False And _
               isFunction(strCurrToken) = False Then
                ' If token is x
                If strCurrToken = "x" Then
                    push (decCurrX)
                ' If token is -x
                ElseIf strCurrToken = "-x" Then
                    push (decCurrX * -1)
                ' If token is e or pi, assign the value
                ElseIf strCurrToken = "pi" Then
                    push (pi)
                ElseIf strCurrToken = "e" Then
                    push (e)
                ' Check if number
                ElseIf IsNumeric(strCurrToken) = False Then
                    Err.Raise 1008, , strCurrToken & " is not valid."
                ' Else Push the token to the operand stack
                Else
                    push (strCurrToken)
                End If
                
        ' If operator or function
        Else
            ' Store current top operand in temp variable if not empty
            If peek <> "" Then
                decOperatorR = CDec(peek)
            Else
                blnSkip = True
            End If
            ' Pop strOperator2
            pop
            
            ' If operation requires two operators
            If strCurrToken = "+" _
            Or strCurrToken = "-" _
            Or strCurrToken = "*" _
            Or strCurrToken = "/" _
            Or strCurrToken = "^" Then
                ' Store next top if not empty
                If peek <> "" Then
                    decOperatorL = CDec(peek)
                Else
                    blnSkip = True
                End If
                ' Pop strOperator 1
                pop
            End If
            
            ' Do current operation
            If blnSkip = False Then
            
                Select Case strCurrToken
                    Case "+"
                        decResult = decOperatorL + decOperatorR
                    Case "-"
                        decResult = decOperatorL - decOperatorR
                    Case "*"
                        decResult = decOperatorL * decOperatorR
                    Case "/"
                        ' Prevent division by zero
                        If decOperatorR = 0 Then
                            decResult = Null
                            ' Warn
                            Call frmAddFunc.displayWarning("Division by zero.", 1)
                        Else
                            decResult = decOperatorL / decOperatorR
                        End If
                    Case "^"
                        ' Prevent roots of negative numbers
                        If decOperatorL < 0 And Abs(decOperatorR) < 1 Then
                            decResult = Null
                            ' Warn
                            Call frmAddFunc.displayWarning("Root of a negative.", 2)
                            
                        ' Prevent negative roots of 0
                        ElseIf decOperatorL = 0 And decOperatorR < 0 Then
                            decResult = Null
                            ' Warn
                            Call frmAddFunc.displayWarning("Negative root of zero leads to division by zero.", 3)
                        Else
                            decResult = decOperatorL ^ decOperatorR
                        End If
                    Case "sin"
                        decResult = Sin(decOperatorR)
                    Case "cos"
                        decResult = Cos(decOperatorR)
                    Case "tan"
                        decResult = Tan(decOperatorR)
                    Case "ln", "log"
                        ' Prevent log of numbers less than 0
                        If decOperatorR <= 0 Then
                            decResult = Null
                            ' Warn
                            Call frmAddFunc.displayWarning("Log of a number smaller than 0.", 4)
                        Else
                            decResult = Log(decOperatorR)
                        End If
                    Case "sqrt"
                        ' Prevent roots of negatives
                        If decOperatorR < 0 Then
                            decResult = Null
                            ' Warn
                            Call frmAddFunc.displayWarning("Square root of a negative.", 5)
                        Else
                            decResult = Sqr(decOperatorR)
                        End If
                    Case "abs"
                        decResult = Abs(decOperatorR)
                    Case Else
                        Err.Raise 1007, , strCurrToken & " is not a supported operation or function."
                End Select
                
            End If
            
            ' Push current result
            push (decResult)
            
        End If
        
    Next intToken
    
    ' At end, output
    evaluate = peek
    
    Exit Function
    
ErrorHandler:
    If Err.Description = "" Then
        Err.Description = "Error with term: " & decOperatorL & " " & strCurrToken & " " & decOperatorR
    End If
    
    Err.Raise Err.Number, , Err.Description
    
End Function

' Stack subroutines
Private Sub pop()
    ReDim Preserve strStack(UBound(strStack) - 1)
End Sub

Private Sub push(item As Variant)
    ReDim Preserve strStack(UBound(strStack) + 1)
    If Not (IsNull(item)) Then
        strStack(UBound(strStack)) = item
    End If
End Sub

Private Function peek() As Variant
    peek = strStack(UBound(strStack))
End Function
