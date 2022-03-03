Attribute VB_Name = "utils"
Option Explicit

'==============================================================================

' Function Statement Syntax Guide:
' --------------------------------

' Function FunctionName(variableName As Argument_DataType) As Return_DataType

'==============================================================================

Function ARRAY_XLOOKUP(arr As String, search_arr As Range, output_arr As Range) As String
    Dim item As Variant
    Dim output As Variant
    Dim res As String
    If arr <> "" Then
        For Each item In Split(arr, ";")
            output = Application.WorksheetFunction.XLookup(Trim(item), search_arr, output_arr, "-")
            If output <> "-" Then
                res = output + ";" + res
            End If
        Next
        ARRAY_XLOOKUP = Left(res, Len(res) - 1)
    Else
        ARRAY_XLOOKUP = "-"
    End If
End Function