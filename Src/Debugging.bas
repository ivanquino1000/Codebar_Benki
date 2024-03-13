Attribute VB_Name = "Debugging"
Option Explicit


Public Sub DebugArray(ByVal arr As Variant)
    Dim i As Long
    Dim result As String
    
    ' Check if the array is empty or missing
    If IsMissing(arr) Or IsEmpty(arr) Then
        Debug.Print "Input array is empty or missing."
        Exit Sub
    End If
    
    ' Check if the array is one-dimensional
    If Not IsArray(arr) Or UBound(arr, 1) = 0 Then
        Debug.Print "Input array is not valid."
        Exit Sub
    End If
    
    ' Loop through the array and concatenate elements to the result string
    For i = LBound(arr) To UBound(arr)
        Debug.Print i, arr(i), vbNewLine
    Next i
End Sub

    
