Attribute VB_Name = "PublicFunctions"
Public Sub ApplyFormat(ByVal rng As range, ByVal format As FormatSettings)

    With rng
        .Interior.Color = format.BgColor
        .Font.Name = format.FontName
        .Font.Size = format.FontSize
        .Font.Color = format.FontColor
        .HorizontalAlignment = format.HAlign
        .VerticalAlignment = format.VAlign
        .Borders.Weight = format.BorderWeight
        .Borders.LineStyle = format.BorderStyle
        .NumberFormat = format.NumberFormat
        .ShrinkToFit = format.Shrink
    End With
End Sub

Public Function getLastRow(ByVal column As Variant, ByRef ws As Worksheet) As Integer
    With ws
        getLastRow = .Cells(.Rows.Count, column).End(xlUp).row
    End With
End Function

Public Function GetParentPath(ByVal Path As String) As String
    Dim currentPath As String

    currentPath = Path

    ' Check if the current path ends with a backslash
    If Right(currentPath, 1) = "\" Then
        ' If it does, remove the trailing backslash
        currentPath = Left(currentPath, Len(currentPath) - 1)
    End If
    ' Use VBA's built-in functions to extract the parent path
    GetParentPath = Left(currentPath, InStrRev(currentPath, "\"))


End Function
Public Function ConvertToNumbers(ByVal arr As Variant) As Variant
    Dim reg         As New RegExp

    Dim Numbers     As Variant
    With reg
        .Global = True
        .IgnoreCase = True
        .Pattern = "\d+"
    End With
    Dim i           As Long
    ReDim Numbers(LBound(arr) To UBound(arr))
    For i = LBound(arr) To UBound(arr)
        Numbers(i) = reg.Execute(arr(i))(0)
    Next i
    ConvertToNumbers = Numbers
End Function

Public Function MissingNumbers(ByVal arr As Variant) As Variant
    'Assumptions:
    'Takes a Sorted  Numeric Array

    'Empty Array
    If IsEmpty(arr) Then
        Exit Function
    End If

    Dim Mia         As Variant: Mia = Array()
    Dim i, j        As Long

    'Save Missings until Lowest Array Value
    Dim LowestValue As Long: LowestValue = arr(LBound(arr))
    Dim HighestValue As Long: HighestValue = arr(UBound(arr))
    'Handle_Missing_Values
    If LowestValue > 1 Then

        ReDim Mia(0 To LowestValue - 2)
        For j = 0 To LowestValue - 2
            Mia(j) = j + 1
        Next j
    End If
    'Store Missing in array
    Dim k           As Long

    For i = arr(LBound(arr)) To arr(UBound(arr))
        If arr(k) = i Then
            k = k + 1
        Else
            ReDim Preserve Mia(0 To UBound(Mia) + 1)
            Mia(UBound(Mia)) = i
        End If
    Next i
'    For i = 0 To UBound(Mia)
'        Debug.Print Mia(i)
'    Next i
    MissingNumbers = Mia
End Function

Sub LabelTest()
    'TODO LIST
    'Empty Value on Fallback ISEMPTY(Array)
'SHRINK TO FIT ALL CELLS
'

    Dim product     As New item
    Dim Company     As New Company
    Dim Lab         As New Label
    With ThisWorkbook.Sheets("LabelSheet")
        .Cells.ClearContents
        .Cells.ClearFormats
         .Parent.Windows(1).Zoom = 50
    End With
    With Lab.product
        .Description = ""
        .Cost = 1
        .Supplier = ""
    End With
    Lab.ToRange


End Sub

Sub PushToVariant(ByRef arr As Variant, ByVal value As Variant)
    ' Redimension the array to add one more element
    ReDim Preserve arr(LBound(arr) To UBound(arr) + 1)
    
    ' Assign the value to the last element
    arr(UBound(arr)) = value
End Sub

Sub MergeRange()
    Dim cell        As range: Set cell = ThisWorkbook.Sheets("LabelSheet").Cells(5, 5)
    Dim Direction   As String
    Dim Places      As Integer

    Direction = "L"
    Places = 3

    Select Case Direction    'R, L, U, D
        Case "R"
            Set cell = cell.Resize(1, 1 + Places)
        Case "L"
            Set cell = cell.Offset(0, -Places).Resize(1, Places + 1)
        Case "U"
            Set cell = cell.Offset(-Places, 0).Resize
        Case "D"
            Set cell = cell.Resize(1 + Places, 1)
    End Select

    ' Merge the resulting range
    cell.Merge
    'Debug.Assert cell.Address
    Debug.Print cell.Address
End Sub

Function BubbleSort(arr As Variant) As Variant
    Dim i As Long, j As Long
    Dim Temp        As Double
    Dim n           As Long

    n = UBound(arr)

    For i = 1 To n - 1
        For j = 1 To n - i
            If arr(j) > arr(j + 1) Then
                ' Swap arr(j) and arr(j + 1)
                Temp = arr(j)
                arr(j) = arr(j + 1)
                arr(j + 1) = Temp
            End If
        Next j
    Next i

    BubbleSort = arr
End Function


Sub test()
    Dim arr1, Arr2  As Variant
    arr1 = Array(2, 4, 5, 7, 8, 12)
    Arr2 = Array(3, 5)
    Dim S, E        As Double
    S = Timer

    arr3 = MissingNumbers(arr1)

    E = Timer
    Debug.Print "Performance - FindMissing:", E - S, "sec"

End Sub


Public Function CodeBuilder(ByVal CodeLet As String, ByVal CodeId As Integer) As String
    Dim Code        As String

    Select Case CodeId
        Case CodeId < 10
            Code = CodeLet & "00" & CodeId
        Case CodeId < 100
            Code = CodeLet & "0" & CodeId
        Case Else
            Code = CodeLet & CodeId
    End Select

    CodeBuilder = Code

End Function

Public Function ExtractNumber(ByVal inputString As String) As Variant

    Dim NumReg      As New RegExp
    With NumReg
        .Global = True
        .IgnoreCase = True
        .Pattern = "\d+"
    End With

    If NumReg.test(inputString) Then
        Set ExtractNumber = NumReg.Execute(inputString)(0)
    Else
        ExtractNumber = "Not Matches Found in Input"
    End If

End Function

Public Function ExtractLetter(ByVal inputString As String) As Variant

    Dim LetReg      As New RegExp
    With LetReg
        .Global = True
        .IgnoreCase = True
        .Pattern = "^([a-zA-Z]+)"
    End With

    If LetReg.test(inputString) Then
        ExtractLetter = UCase(LetReg.Execute(inputString)(0))
    Else
        ExtractLetter = "Not Matches Found in Input"
    End If

End Function
'Deletes Duplicates
Public Function JoinArrays(ByVal MainNumArr As Variant, _
        ByVal OptionalNumArr As Variant) As Variant

    If IsEmpty(MainNumArr) Then
        FindMissCodeId = 0
        Exit Function
    End If

    Dim CombinedArray As Variant
    Dim i As Long, j As Long, k As Long
    Dim isDuplicate As Boolean

    ' Determine the size of the combined array
    ReDim CombinedArray(0 To UBound(MainNumArr) + UBound(OptionalNumArr) + 1)

    ' Merge both arrays into combinedArray


    'Copy InitialArray to CombinedArray
    For i = LBound(MainNumArr) To UBound(MainNumArr)
        CombinedArray(i) = MainNumArr(i)

    Next i
    'Next IdElem after MainArrayCopied / Unique Elem Counter
    k = UBound(MainNumArr) + 1

    'Duplicates Deletion
    For i = LBound(OptionalNumArr) To UBound(OptionalNumArr)
        isDuplicate = False
        For j = LBound(MainNumArr) To UBound(MainNumArr)
            If OptionalNumArr(i) = MainNumArr(j) Then
                isDuplicate = True
                Exit For
            End If
        Next j

        If Not isDuplicate Then
            CombinedArray(k) = OptionalNumArr(i)
            k = k + 1
        End If
    Next i

    ' Redimension the array to the actual size
    ReDim Preserve CombinedArray(1 To k)  '- 1)
    JoinArrays = CombinedArray
End Function


Function FindLatestXLSXFile(ByVal pathDir As String) As String

    Dim fileSystem  As New FileSystemObject
    Dim folderPath  As String
    Dim latestFile  As String
    Dim latestDate  As Date
    Dim file        As Object

    ' Replace "C:\Users\EQUIPO\Downloads" with the path to your data_path directory
    folderPath = pathDir
    ' Initialize variables to hold the latest file information
    latestFile = ""
    latestDate = DateSerial(1900, 1, 1)

    ' Loop through each file in the directory
    For Each file In fileSystem.GetFolder(folderPath).Files
        ' Check if the file is an XLSX file and compare its last modified date
        If LCase(Right(file.Name, 5)) = ".xlsx" And file.DateLastModified > latestDate Then
            latestFile = file.Path
            latestDate = file.DateLastModified
        End If
    Next file
    'Debug.Print latestFile
    FindLatestXLSXFile = latestFile
End Function





