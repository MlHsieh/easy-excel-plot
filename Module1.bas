Attribute VB_Name = "Module1"
Option Explicit
Public Const data1 As String = "data1"
Public Const data2 As String = "data2"
Public Const displayRange1 As String = "D7"
Public Const displayRange2 As String = "D21"

Sub OnClickSelect1()
    Call SelectRange(displayRange1, data1)
End Sub

Sub OnClickSelect2()
    Call SelectRange(displayRange2, data2)
End Sub

Sub SelectRange(display As String, rangeName As String)
    Dim selected As Range
    Dim lastRow As Integer
    
    On Error Resume Next
    Set selected = Application.InputBox(Prompt:="Select data", Type:=8, _
        Default:=Selection.Address)
    With selected
        If Not selected Is Nothing Then
            .Worksheet.Names.Add Name:=rangeName, RefersTo:=selected
            ThisWorkbook.Names("'" & .Worksheet.Name & "'!" & rangeName).Visible = False
            Range(display).Value = selected.Address(False, False)
        End If
    End With
End Sub

Sub OnClickClear2()
    Call ClearRange(displayRange2, data2)
End Sub

Sub ClearRange(display As String, rangeName As String)
    On Error Resume Next
    ActiveSheet.Names("'" & ActiveSheet.Name & "'!" & rangeName).Delete
    Range(display).Value = "--"
End Sub

Sub Plot()
    Dim n As Variant
    Dim nameStr As String
       
    MLEvalString "clear variables"
    
    MLPutMatrix data1, SelectNonBlank(Range("'" & ActiveSheet.Name & "'!" & data1))

    On Error Resume Next
    MLPutMatrix data2, SelectNonBlank(Range("'" & ActiveSheet.Name & "'!" & data2))
    
    For Each n In GetNames()
        nameStr = "'" & ActiveSheet.Name & "'!" & n
        MLPutMatrix n, Range(nameStr)
    Next n
    
    On Error GoTo 0
    MLEvalString "PlotColumns"
End Sub

Function SelectNonBlank(dataRange As Range) As Range
    Dim lastRow As Integer
    With dataRange
        If .Address = .EntireColumn.Address Then
            lastRow = .Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
            Set SelectNonBlank = dataRange.Resize(lastRow)
        Else
            Set SelectNonBlank = dataRange
        End If
    End With
End Function

Function GetNames()
    GetNames = Array("optionName1", "optionVal1", "optionName2", "optionVal2")
End Function

Sub Test()
    Dim lastRow As Integer
    Dim selected As Range
    Set selected = Range("'" & ActiveSheet.Name & "'!" & data2)
    SelectNonBlank(selected).Select
End Sub
