Attribute VB_Name = "modFunctionRecord"
Option Explicit

Public Type functionRecord
    strFuncInput As String
    decDomainMin As Variant
    decDomainMax As Variant
    lngColour    As Long
    strPostfix   As String
    decyValues   As Variant
End Type

Public funcList() As functionRecord

Public Sub funcListInitialise()
    ReDim funcList(0)
End Sub
