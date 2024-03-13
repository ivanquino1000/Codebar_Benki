Attribute VB_Name = "PublicFunctions"


Public Sub ApplyFormat(ByVal rng As Range, ByVal format As FormatSettings)

    With rng
        .Interior.Color = format.BgColor
        .Font.Name = format.FontName
        .Font.Size = format.fontSize
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


Function RandomLetter() As String
    ' Seed the random number generator
    Randomize

    ' Generate a random number between 65 (ASCII for 'A') and 90 (ASCII for 'Z')
    Dim randomAsciiCode As Integer
    randomAsciiCode = Int((90 - 65 + 1) * Rnd + 65)

    ' Convert the ASCII code to a character and return it
    RandomLetter = Chr(randomAsciiCode)
End Function

Public Function ConvertToNumbers(ByVal arr As Variant) As Variant
    Dim reg         As New RegExp

    Dim Numbers     As Variant
    With reg
        .Global = True
        .IgnoreCase = True
        .pattern = "\d+"
    End With
    Dim i           As Long
    ReDim Numbers(LBound(arr) To UBound(arr))
    For i = LBound(arr) To UBound(arr)
        Numbers(i) = reg.Execute(arr(i))(0)
    Next i
    ConvertToNumbers = Numbers
End Function

Public Function MissingNumbers(ByVal arr As Variant) As Variant
    '@Assumptions:
    '       -Takes a Sorted  Numeric Array

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
    If UBound(Mia) = -1 Then
        ' Return an empty array
        MissingNumbers = Empty
    Else
        MissingNumbers = Mia
    End If
    'MissingNumbers = Mia
End Function

Public Function ClearSearchInput(ByVal inputString As Variant) As Variant

    If Len(inputString) > 0 Then
        Dim firstChar As String
        firstChar = Left(inputString, 1)
        Select Case firstChar
            Case "*"
                ClearSearchInput = Mid(inputString, 2, Len(inputString))

            Case "<"
                ClearSearchInput = Mid(inputString, 3, Len(inputString))
            Case Else
                ClearSearchInput = inputString
        End Select
    End If
End Function

Sub LabelTest()

    Dim product     As New item
    Dim company     As New company
    Dim Lab         As New Label
    With ThisWorkbook.Sheets("LabelSheet")
        .Cells.ClearContents
        .Cells.ClearFormats
        .Parent.Windows(1).Zoom = 50
    End With
    With Lab.product
        .Description = ""
        .Description = "@ITEM_DESCRIPTION_DEFAULT"
        .Supplier = ""
        '.BoxQty = 2
        '.Cost = 3.5
        '.WholeSalePrice = 10
        '.BoxPrice = 12
        .SellingPrice = 15
    End With
    Lab.ToRange


End Sub

Public Function ValidBookPath(ByVal Path As String) As Boolean
    On Error Resume Next

    ' Attempt to open the workbook
    Dim tempWorkbook As Workbook
    Set tempWorkbook = Workbooks.Open(Path)

    ' Check if there was an error
    If err.Number <> 0 Then
        ValidBookPath = False
    Else
        ValidBookPath = True
        ' Close the workbook if it was successfully opened
        tempWorkbook.Close SaveChanges:=False
    End If

    ' Reset error handling
    On Error GoTo 0
End Function

Sub PushToVariant(ByRef arr As Variant, ByVal value As Variant)
    ' Redimension the array to add one more element
    ReDim Preserve arr(LBound(arr) To UBound(arr) + 1)

    ' Assign the value to the last element
    arr(UBound(arr)) = value
End Sub

Sub MergeRange()
    Dim cell        As Range: Set cell = ThisWorkbook.Sheets("LabelSheet").Cells(5, 5)
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
    Dim temp        As Double
    Dim n           As Long

    n = UBound(arr)

    For i = 1 To n - 1
        For j = 1 To n - i
            If arr(j) > arr(j + 1) Then
                ' Swap arr(j) and arr(j + 1)
                temp = arr(j)
                arr(j) = arr(j + 1)
                arr(j + 1) = temp
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
    Dim code        As String

    Select Case True
        Case CodeId < 10
            code = CodeLet & "00" & CodeId
        Case CodeId < 100
            code = CodeLet & "0" & CodeId
        Case Else
            code = CodeLet & CodeId
    End Select

    CodeBuilder = code
End Function

Public Function ExtractNumber(ByVal inputString As String) As Variant

    Dim NumReg      As New RegExp
    With NumReg
        .Global = True
        .IgnoreCase = True
        .pattern = "\d+"
    End With

    If NumReg.test(inputString) Then
        ExtractNumber = CInt(NumReg.Execute(inputString)(0))
    Else
        ExtractNumber = "Not Matches Found in Input"
    End If

End Function

Function FindFirstNonIncludedElement(ByVal srcArr As Variant, ByVal compareArr As Variant) As Variant
    Dim i           As Long
    Dim j           As Long
    Dim isElementIncluded As Boolean

    ' Check if the arrays are empty or missing
    If IsMissing(srcArr) Or IsEmpty(srcArr) Then
        FindFirstNonIncludedElement = Null
        Exit Function
    End If

    If IsMissing(compareArr) Or IsEmpty(compareArr) Then
        FindFirstNonIncludedElement = srcArr(LBound(srcArr))
        Exit Function
    End If

    ' Check if the arrays are one-dimensional
    If Not IsArray(srcArr) Then
        Debug.Print "Input array is not valid."
        FindFirstNonIncludedElement = Null
        Exit Function
    End If

    If Not IsArray(compareArr) Then
        Debug.Print "compare array is not valid."
        FindFirstNonIncludedElement = Null
        Exit Function
    End If

    ' Loop through srcArr and compare each element with compareArr
    For i = LBound(srcArr) To UBound(srcArr)
        isElementIncluded = False

        For j = LBound(compareArr) To UBound(compareArr)
            If srcArr(i) = compareArr(j) Then
                ' Element is included, set flag and exit inner loop
                isElementIncluded = True
                Exit For
            End If
        Next j

        ' If the element is not included in compareArr, return it
        If Not isElementIncluded Then
            FindFirstNonIncludedElement = srcArr(i)
            Exit Function
        End If
    Next i

    ' All elements are included in compareArr
    FindFirstNonIncludedElement = ""
End Function

Public Function RemoveFromArr(ByVal arr As Variant, ByVal value As Variant) As Variant
    Dim i           As Long
    Dim foundIndex  As Long

    ' Find the index of the value to delete
    foundIndex = -1
    For i = LBound(arr) To UBound(arr)
        If arr(i) = value Then
            foundIndex = i
            Exit For
        End If
    Next i

    ' If the value is found, delete it
    If foundIndex >= 0 Then
        For i = foundIndex To UBound(arr) - 1
            arr(i) = arr(i + 1)
        Next i

        ' Resize the array to remove the last element
        If UBound(arr) = LBound(arr) Then
            ' If only one element is left, set the array to an empty array
            arr = Empty
        Else
            ReDim Preserve arr(LBound(arr) To UBound(arr) - 1)
        End If

    End If
    RemoveFromArr = arr
End Function


Public Function ExtractLetter(ByVal inputString As String) As Variant

    Dim LetReg      As New RegExp
    With LetReg
        .Global = True
        .IgnoreCase = True
        .pattern = "^([a-zA-Z]+)"
    End With

    If LetReg.test(inputString) Then
        ExtractLetter = UCase(LetReg.Execute(inputString)(0))
    Else
        ExtractLetter = "Not Matches Found in Input"
    End If

End Function
'Deletes Duplicates
Public Function JoinArrays(ByVal MainNumArr As Variant, _
                            ByVal OptionalNumArr As Variant, _
                            Optional ByVal DeleteDuplicates As Boolean = True) As Variant

    If IsEmpty(MainNumArr) Then
        FindMissCodeId = 0
        Exit Function
    End If

    If IsEmpty(OptionalNumArr) Then
        JoinArrays = MainNumArr
        Exit Function
    End If

    Dim CombinedArray As Variant
    Dim i As Long, j As Long, k As Long
    Dim isDuplicate As Boolean

    ' Determine the size of the combined array
    ReDim CombinedArray(0 To UBound(MainNumArr) + UBound(OptionalNumArr) + 1)

    ' Merge both arrays into CombinedArray

    ' Copy MainNumArr to CombinedArray
    For i = LBound(MainNumArr) To UBound(MainNumArr)
        CombinedArray(i) = MainNumArr(i)
    Next i

    ' Next IdElem after MainArrayCopied / Unique Elem Counter
    k = UBound(MainNumArr) + 1

    ' Duplicates Deletion if specified
    If DeleteDuplicates Then
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
    Else
        ' If duplicates are not deleted, just append all elements from OptionalNumArr
        For i = LBound(OptionalNumArr) To UBound(OptionalNumArr)
            CombinedArray(k) = OptionalNumArr(i)
            k = k + 1
        Next i
    End If

    ' Redimension the array to the actual size
    ReDim Preserve CombinedArray(1 To k)
    JoinArrays = CombinedArray
End Function

Public Sub convertCost(ByRef Cost As Variant)
    If Cost <> "" And IsNumeric(Cost) And Cost <> 0 Then
        Dim Text    As String
        Text = Trim(CStr(Application.WorksheetFunction.RoundUp(Cost, 2) * 100))
        Numbers = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9")
        letters = Array("L", "Y", "Q", "I", "M", "P", "O", "R", "T", "S")
        For Each a In Numbers
            Dim pos
            pos = Application.match(a, Numbers, False)
            If Not IsError(pos) Then
                Text = Replace(Text, a, letters(pos - 1))
            Else: Exit Sub
            End If
        Next a
        Select Case Cost
            Case Is < 0.1
                Cost = "L.L" & Text
            Case Is < 1
                Cost = "L." & Text
            Case Else
                Cost = Left(Text, Len(Text) - 2) & "." & Right(Text, 2)

        End Select
    End If
End Sub

Public Sub ConfigurePrintingPreferences(ByVal ws As Worksheet, ByVal labelRange As String)
    Application.PrintCommunication = False
    With ws.PageSetup
        .LeftMargin = Application.InchesToPoints(0)
        .RightMargin = Application.InchesToPoints(0)
        .TopMargin = Application.InchesToPoints(0)
        .BottomMargin = Application.InchesToPoints(0)
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0)

        .PrintArea = labelRange
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    Application.PrintCommunication = False
End Sub

Function FindHighestOrJump(ByVal arr As Variant) As Long
    Dim sortedArr() As Variant
    Dim i           As Long

    ' Remove duplicates
    sortedArr = RemoveDuplicates(arr)

    ' Sort the array
    QuickSort sortedArr, LBound(sortedArr), UBound(sortedArr)

    ' Iterate through the sorted array to find the first jump
    For i = LBound(sortedArr) To UBound(sortedArr) - 1
        If sortedArr(i + 1) <> sortedArr(i) + 1 Then
            ' If a jump is found, return the value before the jump
            FindHighestOrJump = sortedArr(i)
            Exit Function
        End If
    Next i

    ' If no jump is identified, return the last value in the array
    FindHighestOrJump = sortedArr(UBound(sortedArr))
End Function

Sub QuickSort(ByVal arr As Variant, ByVal low As Long, ByVal high As Long)
    Dim pivot       As Variant
    Dim temp        As Variant
    Dim i           As Long
    Dim j           As Long

    If low < high Then
        pivot = arr((low + high) \ 2)
        i = low
        j = high

        Do
            Do While arr(i) < pivot
                i = i + 1
            Loop

            Do While arr(j) > pivot
                j = j - 1
            Loop

            If i <= j Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
                i = i + 1
                j = j - 1
            End If
        Loop While i <= j

        QuickSort arr, low, j
        QuickSort arr, i, high
    End If
End Sub

Function RemoveDuplicates(ByVal arr As Variant) As Variant
    Dim uniqueDict  As Object
    Set uniqueDict = CreateObject("Scripting.Dictionary")

    Dim element     As Variant
    For Each element In arr
        If Not uniqueDict.Exists(element) Then
            uniqueDict.Add element, element
        End If
    Next element

    RemoveDuplicates = uniqueDict.Keys
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





