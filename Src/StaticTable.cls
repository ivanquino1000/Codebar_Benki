VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StaticTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const FileNotFoundMsg As String = "No Se Encontro un Ningun Archivo en la ruta: " & "/downloads" & vbNewLine & _
        "Se Usaran los Datos de la Ultima Sesion Termida"

Const NonRecentFileMsg As String = "No Se Encontro Ningun Archivo de Hoy Dia " & vbNewLine & _
        "Se Usaran los Datos de la Ultima Sesion" & vbNewLine & _
        "Nota: Descarga de la Web y reinicia este libro para usar datos Actuales"

Implements IDataTable

Public Event HeadersPlaced(ByVal Headers As Variant)
Public Event elementAdded(ByVal Element As item)
Public Event TableUpdated()
Public Event ElementDeleted(ByVal Element As item)
Public Event DataGotten(ByVal elem_num As Long)

Public WithEvents HeadersComp As Headers
Attribute HeadersComp.VB_VarHelpID = -1

Private Type SourceTab
    FirstRow        As Long
    CodeCol         As String
    NameCol         As String
    PriceCol        As String
End Type

Private SourceTableRef As SourceTab

Private Type TStaticTable
    row             As Integer
    column          As Integer
    range           As range
    ws              As Worksheet
    'Headers         As Headers
    ElementsRange   As range
    PreviousElemRange As range
    ElementsCounter As Integer
    ElementsFormat  As FormatSettings
    SrcBookPath     As String
    MissingElementsIds As Dictionary
End Type

Private this        As TStaticTable



Private Sub Class_Initialize()
    With this
        .row = 1
        .column = 1

        '.SrcBookPath = FindLatestXLSXFile(Environ("USERPROFILE") & "\Downloads\")
        Set .ws = ThisWorkbook.ActiveSheet
        Set .ElementsFormat = New FormatSettings
        Set .MissingElementsIds = New Dictionary
        With .ElementsFormat
            .BgColor = RGB(249, 244, 245)
            .FontColor = vbBlack
            .FontName = "Bahnschrift SemiLight"
            .FontSize = 16
        End With

    End With
    Set HeadersComp = New Headers
    With SourceTableRef
        .CodeCol = "B"
        .FirstRow = 4
        .NameCol = "E"
        .PriceCol = "K"
    End With

End Sub

Public Sub init(ByVal row As Integer, _
        ByVal column As Integer, _
        ByVal ws As Worksheet, _
        ByVal Headers As Variant _
        )
    With this
        .row = row
        .column = column
        Set .ws = ws
        With HeadersComp
            .values = Headers
        End With

        Update
    End With


End Sub

'TODO: IMPLEMENT LOGIC FOR LAST FILE UPDATED TO BE DATA SOURCE
'TODO: DOWNLOADS/ LASTSESSION CHECK LATEST MODIFIED
Public Sub GetData()

    On Error GoTo getDataError

    Dim InitTime, EndTime As Double
    InitTime = Timer

    Dim BackupBookPath As String
    BackupBookPath = GetParentPath(ThisWorkbook.Path) & "SessionBackup" & "\LastSessionItems.xlsx"
    Dim SourceBook  As Workbook

    On Error Resume Next
    With this
        .SrcBookPath = FindLatestXLSXFile(Environ("USERPROFILE") & "\Downloads\")
        Set SourceBook = Workbooks.Open(.SrcBookPath)

        If err.Number <> 0 Then
            On Error GoTo 0
            MsgBox FileNotFoundMsg
            GoTo UseBackup
        End If

        On Error GoTo 0
        If Int(FileDateTime(.SrcBookPath)) <> Int(Now) Or _
                Int(FileDateTime(.SrcBookPath)) < Int(FileDateTime(BackupBookPath)) Then
            MsgBox NonRecentFileMsg
            SourceBook.Close False
            GoTo UseBackup
        End If

        GoTo ProcessData
    End With

UseBackup:
    On Error Resume Next
    Set SourceBook = Workbooks.Open(BackupBookPath)

    If err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "Libro de Respaldo No Encontrado: FATAL ERROR "
        Exit Sub
    End If
    GoTo ProcessData

ProcessData:

    Dim SourceSheet As Worksheet: Set SourceSheet = SourceBook.Worksheets(1)

    'SET_TABLE_REFERENCES
    Dim ItemsCounter As Long
    Dim ItemsLastRow As Long
    Dim CodesRange  As range
    Dim NameRange   As range
    Dim PriceRange  As range
    Dim TableRange  As range
    Dim regex       As New RegExp


    With SourceTableRef
        ItemsLastRow = getLastRow(.CodeCol, SourceSheet)
        SetRangeRefences CodesRange, SourceSheet, .FirstRow, .CodeCol, ItemsLastRow
        SetRangeRefences NameRange, SourceSheet, .FirstRow, .NameCol, ItemsLastRow
        SetRangeRefences PriceRange, SourceSheet, .FirstRow, .PriceCol, ItemsLastRow
        SetRangeRefences TableRange, SourceSheet, .FirstRow, .CodeCol, ItemsLastRow, .PriceCol
    End With

    'CLEAR_REGEX_NOTMATCHING
    Dim codearr     As Variant
    
    With TableRange
        .RemoveDuplicates Columns:=Array(1), Header:=xlNo
    End With

    codearr = Application.Transpose(CodesRange)

    CleanCodes codearr, "^[A-Z,a-z]{1,2}\d{3,4}\b"
    CodesRange.value = Application.Transpose(codearr)

    'SORT_SELECTION
    With TableRange
        .Sort key1:=.Columns(1), order1:=xlAscending, Header:=xlNo
    End With

    'SET_TABLE_REFERENCES
    With SourceTableRef
        ItemsLastRow = getLastRow(.CodeCol, SourceSheet)
        SetRangeRefences CodesRange, SourceSheet, .FirstRow, .CodeCol, ItemsLastRow
        SetRangeRefences NameRange, SourceSheet, .FirstRow, .NameCol, ItemsLastRow
        SetRangeRefences PriceRange, SourceSheet, .FirstRow, .PriceCol, ItemsLastRow
        ItemsCounter = ItemsLastRow - .FirstRow + 1
    End With


    'SAVE_TO_TABLE_ELEMENTS_REFERENCE_ID

    Dim nameArr, priceArr, tableArr As Variant
    codearr = Application.Transpose(CodesRange)
    nameArr = Application.Transpose(NameRange)
    priceArr = Application.Transpose(PriceRange)

    With this
        .ElementsCounter = ItemsCounter
        UpdateRanges
        .ElementsRange.Value2 = CombineArr(codearr, nameArr, priceArr)
    End With

    StoreCodeIds codearr

    SourceBook.Close False
    EndTime = Timer
    Debug.Print "GetData - Performance:", EndTime - InitTime & " sec"

    Update
    Exit Sub

getDataError:
    Debug.Print "getDataError:  "; err.Description
    If Not SourceBook Is Nothing And SourceSheet.Visible Then
        SourceBook.Close False
    End If

End Sub

Private Function CombineArr(ByVal arr1 As Variant, ByVal Arr2 As Variant, ByVal arr3 As Variant)
    Dim i           As Integer
    Dim arr         As Variant
    ReDim arr(1 To UBound(arr1), 1 To 3)
    For i = 1 To UBound(arr1)
        arr(i, 1) = arr1(i)
        arr(i, 2) = Arr2(i)
        arr(i, 3) = arr3(i)
    Next i
    CombineArr = arr
End Function


Private Sub CleanCodes(ByRef arr As Variant, ByVal regexPattern As String)
    Dim reg         As New RegExp
    Dim i           As Integer

    With reg
        .Global = True
        .Pattern = regexPattern
    End With

    For i = LBound(arr) To UBound(arr)

        Dim Code    As String
        Code = arr(i)
        If Code = "" Then
            GoTo nextId
        End If
        If Not reg.test(Code) Then
            arr(i) = ""
            GoTo nextId
        End If
        arr(i) = UCase(Code)
nextId:
    Next i

End Sub

Private Sub StoreCodeIds(ByRef codearr As Variant)
    Dim Ids()       As Variant
    Dim ConsecutiveCounter As Long
    Dim startIndex  As Long
    Dim i           As Long

    ' Convert the codearr to an array of numbers
    Ids = ConvertToNumbers(codearr)
    startIndex = LBound(Ids)

    Do While startIndex <= UBound(Ids)
        Dim key     As String
        Dim Temp()  As Long
        ConsecutiveCounter = 0
        key = ExtractLetter(codearr(startIndex))

        ' COUNT_ALL_KEYCODE_RELATED_IDS
        i = startIndex
        While i < UBound(Ids) And Ids(i) < Ids(i + 1)
            ConsecutiveCounter = ConsecutiveCounter + 1
            i = i + 1
            If i = UBound(Ids) Then
                GoTo FinalIdSegment
            End If

        Wend
FinalIdSegment:
        'STORE_KEY_IDS_TEMP_ARRAY
        ReDim Temp(0 To ConsecutiveCounter)
        For i = 0 To ConsecutiveCounter
            Temp(i) = Ids(startIndex + i)
        Next i

        ' Add the Temp array to the dictionary
        this.MissingElementsIds.Add key, MissingNumbers(Temp)

        ' Move startIndex to the next non-consecutive number
        startIndex = startIndex + ConsecutiveCounter + 1
    Loop
End Sub

Public Sub SetRangeRefences(ByRef range As range, ByRef ws As Worksheet, _
        ByVal startRow As Variant, ByVal startColumn As Variant, _
        ByVal endRow As Variant, Optional endColumn As Variant)
    If IsMissing(endColumn) Then
        Set range = ws.range(startColumn & startRow & ":" & _
                startColumn & endRow)
        Exit Sub
    End If
    Set range = ws.range( _
            startColumn & startRow & ":" & _
            endColumn & endRow)

End Sub

Public Sub Update()
    UpdateRanges
    With this
        .ws.Select
        HeadersComp.Update .ws
        ElementsUpdate
        .ws.Columns.AutoFit
    End With

End Sub

Public Sub ClearPrevRange()
    If Not this.PreviousElemRange Is Nothing Then
        With this.PreviousElemRange
            .ClearContents
            .ClearFormats
            .Interior.Color = RGB(17, 21, 28)
        End With
    End If
End Sub

Public Sub ElementsUpdate()

    With this
        Call ApplyFormat(.ElementsRange, .ElementsFormat)

        '        .ElementsRange.Sort key1:=.ElementsRange.Columns(UBound(HeadersComp.values) + 1), _
                 '                order1:=xlDescending, Header:=xlNo

        Set .PreviousElemRange = .ElementsRange

        .range.EntireRow.RowHeight = 40
    End With

End Sub


'Dependencies: Row,Column, _
 Headers.values,ElementsCounter

Private Sub UpdateRanges()
    With HeadersComp
        Set .range = this.ws.range( _
                this.ws.Cells(this.row, this.column), _
                this.ws.Cells(this.row, this.column + UBound(.values)) _
                )
    End With

    With this
        Dim EmptyElements As Integer
        If .ElementsCounter = 0 Then
            EmptyElements = 1
        End If
        Set .ElementsRange = this.ws.range( _
                this.ws.Cells(this.row + 1, this.column), _
                this.ws.Cells(this.row + EmptyElements + .ElementsCounter, this.column + UBound(HeadersComp.values)) _
                )
        Set .range = Union( _
                HeadersComp.range, _
                .ElementsRange _
                )
    End With
End Sub



Private Sub HeadersComp_HeadersUpdated(ByVal Headers As Variant)
    UpdateRanges
End Sub

Public Property Let row(ByVal value As Integer)
    this.row = value
    UpdateRanges
End Property


Public Property Let column(ByVal value As Integer)
    this.column = value
    UpdateRanges
End Property

Public Property Let ElementsCounter(ByVal RHS As Integer)
    this.ElementsCounter = RHS
    Update
End Property


'Public Property Get Headers() As Headers
'    Set Headers = this.Headers
'End Property
'
'Public Property Set Headers(ByVal RHS As Headers)
'    Set this.Headers = RHS
'End Property

Public Property Get ElementsRange() As range
    Set ElementsRange = this.ElementsRange
End Property

Public Property Set ElementsRange(ByVal RHS As range)
    Set this.ElementsRange = RHS
    UpdateRanges
End Property

Public Property Get PreviousElemRange() As range
    Set PreviousElemRange = this.PreviousElemRange
End Property

Public Property Set PreviousElemRange(ByVal RHS As range)
    Set this.PreviousElemRange = RHS
End Property

Public Property Get ElementsCounter() As Integer
    ElementsCounter = this.ElementsCounter
End Property

Public Property Get range() As range
    Set range = this.range
End Property

Public Property Set range(ByVal value As range)
    Set this.range = value
End Property

Public Property Get row() As Integer
    row = this.row
End Property



Public Property Get ws() As Worksheet
    Set ws = this.ws
End Property

Public Property Set ws(ByVal value As Worksheet)
    Set this.ws = value
End Property

Public Property Get column() As Integer
    column = this.column
End Property


Public Property Get ElementsFormat() As FormatSettings
    Set ElementsFormat = this.ElementsFormat
End Property

Public Property Set ElementsFormat(ByVal RHS As FormatSettings)
    Set this.ElementsFormat = RHS
End Property

Private Sub IDataTable_Update()
    err.Raise 5                                  'TODO implement interface member
End Sub


Public Property Get SrcBookPath() As String
    SrcBookPath = this.SrcBookPath
End Property

Public Property Let SrcBookPath(ByVal RHS As String)
    this.SrcBookPath = RHS
End Property



Public Property Get MissingElementsIds() As Dictionary
    Set MissingElementsIds = this.MissingElementsIds
End Property

Public Property Set MissingElementsIds(ByVal RHS As Dictionary)
    Set this.MissingElementsIds = RHS
End Property


