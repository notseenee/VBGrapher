Attribute VB_Name = "modPostfix"
Option Explicit
' Moved here from toPostfix function to allow access from stack subroutines
Dim strStack() As String
Dim strOutput  As String

Public Function toPostfix(ByVal strInfix As String) As String
    Dim strFormatted As String
    Dim strTokens()  As String
    Dim intToken     As Integer
    Dim strCurrToken As String
    
    ' Format input
    strFormatted = formatter(strInfix)
    
    ' Tokenise input into array
    strTokens = Split(strFormatted)
    
    ' Initialise & clear stack array
    ReDim strStack(0)
    
    ' Clear output string
    strOutput = ""
    
    ' Loop through tokens
    For intToken = 0 To UBound(strTokens)
        strCurrToken = strTokens(intToken)
        
        ' If the token is a function token, then push it onto the stack.
        If isFunction(strCurrToken) = True Then
            push (strCurrToken)
            
        ' If the token is an operator o1:
        ElseIf isOperator(strCurrToken) = True Then
        
            ' While there is an operator token o2 at the top of the operator stack and
            ' o1 is left-associative and its precedence is less than or equal to that of o2
            Do While isOperator(peek) = True _
                And strCurrToken <> "^" And _
                getPrecedence(strCurrToken) <= getPrecedence(peek)
                
                    ' pop o2 off the operator stack, onto the output queue
                    output (peek)
                    pop
            Loop
            
            ' push o1 onto the operator stack
            push (strCurrToken)
        
        ' If the token is a left parenthesis, then push it onto the stack.
        ElseIf strCurrToken = "(" Then
            push (strCurrToken)
            
        ' If the token is a right parenthesis
        ElseIf strCurrToken = ")" Then
            ' Until the token at the top of the stack is a left parenthesis,
            Do While peek <> "("
                ' pop operators off the stack onto the output queue.
                output (peek)
                On Error Resume Next
                pop
                
                ' If the stack runs out without finding a left parenthesis,
                ' then there are mismatched parentheses.
                If UBound(strStack) = 0 Then
                    Err.Raise 1006, , "There are too many parentheses in your function."
                    Exit Do
                End If
            Loop
            ' Pop the left parenthesis from the stack, but not onto the output queue.
            If peek = "(" Then pop
            ' If the token at the top of the stack is a function token, pop it onto the output queue.
            If isFunction(peek) = True Then
                output (peek)
                pop
            End If
        
        ' If the token is an operand, append it to the postfix output.
        Else
            output (strCurrToken)
        
        End If
        
    Next intToken
    
    ' When there are no more tokens to read:
    ' While there are still operator tokens in the stack:
    Do While UBound(strStack) <> 0
        ' If the operator token on the top of the stack is a parenthesis,
        ' then there are mismatched parentheses.
        If peek = "(" Or peek = ")" Then
            Err.Raise 1005, , "There are too many parentheses in your function."
            Exit Do
        ' Pop the operator onto the output queue.
        Else
            output (peek)
            pop
        End If
    Loop
    
    ' Final output
    toPostfix = strOutput
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' HELPER FUNCTIONS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function formatter(ByVal strInfix As String) As String
    Dim strFormatted As String
    
    ' Convert to lowercase
    strFormatted = LCase(strInfix)
    
    ' Adds spaces before & after operators
    strFormatted = Replace$(strFormatted, "+", " + ")
    ' strFormatted = Replace$(strFormatted, "-", " - ")
    strFormatted = Replace$(strFormatted, "*", " * ")
    strFormatted = Replace$(strFormatted, "/", " / ")
    strFormatted = Replace$(strFormatted, "^", " ^ ")
    
    ' Adds spaces before & after parentheses
    strFormatted = Replace$(strFormatted, "(", " ( ")
    strFormatted = Replace$(strFormatted, ")", " ) ")
    
    ' Remove double spaces
    Do While InStr(strFormatted, "  ") <> 0
        strFormatted = Replace$(strFormatted, "  ", " ")
    Loop
    
    ' Trim leading & trailing whitespaces
    strFormatted = Trim$(strFormatted)
    
    ' Output
    formatter = strFormatted
End Function

' Check if operator
Public Function isOperator(ByVal token As String) As Boolean
    Select Case token
        Case "+", "-", "*", "/", "^"
            isOperator = True
        Case Else
            isOperator = False
    End Select
End Function

' Check if function
Public Function isFunction(ByVal token As String) As Boolean
    Select Case token
        Case "sin", "cos", "tan", "ln", "log", "sqrt", "abs"
            isFunction = True
        Case Else
            isFunction = False
    End Select
End Function

' Returns an integer value for operator precedence based on BODMAS
Private Function getPrecedence(ByVal strOperation As String) As Integer
    Select Case strOperation
        Case "^"
            getPrecedence = 3
        Case "/", "*"
            getPrecedence = 2
        Case "+", "-"
            getPrecedence = 1
        Case Else
            getPrecedence = 0
    End Select
End Function

' Stack subroutines
Private Sub pop()
    ReDim Preserve strStack(UBound(strStack) - 1)
End Sub

Private Sub push(item As String)
    ReDim Preserve strStack(UBound(strStack) + 1)
    strStack(UBound(strStack)) = item
End Sub

Private Function peek() As String
    peek = strStack(UBound(strStack))
End Function

Private Sub output(item As String)
    strOutput = strOutput & item & " "
End Sub
