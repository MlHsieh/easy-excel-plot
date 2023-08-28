Attribute VB_Name = "Module1"
Option Explicit
Public Const name1 As String = "data1"
Public Const name2 As String = "data2"

Sub OnClickSelect1()
    Call SelectRange(name1)
End Sub

Sub OnClickSelect2()
    Call SelectRange(name2)
End Sub

Sub SelectRange(rangeName As String)
    Dim selected As Range
    On Error Resume Next
    Set selected = Application.InputBox(Prompt:="Select data", Type:=8, _
        Default:=Selection.Address)
    If Not selected Is Nothing Then
        selected.Worksheet.Names.Add name:=rangeName, RefersTo:=selected
    End If
End Sub

Sub OnClickClear2()
    Call ClearRange(name2)
End Sub

Sub ClearRange(rangeName As String)
    On Error Resume Next
    ActiveSheet.Names("'" & ActiveSheet.name & "'!" & name2).Delete
End Sub

Sub Plot()
    Dim n As Variant
    Dim nameStr As String
       
    MLEvalString "clear variables"
    
    MLPutMatrix name1, Range("'" & ActiveSheet.name & "'!" & name1)
    
    On Error Resume Next
    MLPutMatrix name2, Range("'" & ActiveSheet.name & "'!" & name2)
    
    For Each n In GetNames()
        nameStr = "'" & ActiveSheet.name & "'!" & n
        MLPutMatrix n, Range(nameStr)
    Next n
    
    On Error GoTo 0
    MLEvalString "PlotColumns"
End Sub

Function GetNames()
    GetNames = Array("optionName1", "optionVal1", "optionName2", "optionVal2")
End Function
