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
    Dim lastRow As Integer
    
    On Error Resume Next
    Set selected = Application.InputBox(Prompt:="Select data", Type:=8, _
        Default:=Selection.Address)
    With selected
        If Not selected Is Nothing Then
            If .Address = .EntireColumn.Address Then
                lastRow = .Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
                Set selected = .Resize(lastRow)
            End If
            .Worksheet.Names.Add Name:=rangeName, RefersTo:=selected
            ThisWorkbook.Names("'" & .Worksheet.Name & "'!" & rangeName).Visible = False
        End If
    End With
End Sub

Sub OnClickClear2()
    Call ClearRange(name2)
End Sub

Sub ClearRange(rangeName As String)
    On Error Resume Next
    ActiveSheet.Names("'" & ActiveSheet.Name & "'!" & rangeName).Delete
End Sub

Sub Plot()
    Dim n As Variant
    Dim nameStr As String
       
    MLEvalString "clear variables"
    
    MLPutMatrix name1, Range("'" & ActiveSheet.Name & "'!" & name1)
    
    On Error Resume Next
    MLPutMatrix name2, Range("'" & ActiveSheet.Name & "'!" & name2)
    
    For Each n In GetNames()
        nameStr = "'" & ActiveSheet.Name & "'!" & n
        MLPutMatrix n, Range(nameStr)
    Next n
    
    On Error GoTo 0
    MLEvalString "PlotColumns"
End Sub

Function GetNames()
    GetNames = Array("optionName1", "optionVal1", "optionName2", "optionVal2")
End Function

Sub Test()
    Dim lastRow As Integer
    Dim selected As Range
    Set selected = Selection
    With Selection
        If .Address = .EntireColumn.Address Then
            lastRow = .Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
            selected.Resize(lastRow).Select
        End If
    End With
End Sub
