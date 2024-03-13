VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DynamicTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Implements IDataTable

Public Event HeadersPlaced(ByVal Headers As Variant)
Public Event elementAdded(ByVal element As item)
Public Event TableUpdated()
Public Event ElementDeleted(ByVal element As item)
Public Event DataGotten(ByVal elem_num As Long)

Public WithEvents HeadersComp As Headers
Attribute HeadersComp.VB_VarHelpID = -1
Public WithEvents ElementsComp As Elements
Attribute ElementsComp.VB_VarHelpID = -1


Private Type SourceTab
    FirstRow        As Long
    CodeCol         As String
    NameCol         As String
    PriceCol        As String
End Type

Private SourceTableRef As SourceTab

Private Type TDynamicTable
    row             As Integer
    column          As Integer
    Range           As Range
    SrcBookPath     As String
    NewItemsBookPath As String
    TestBookPath    As String
    ws              As Worksheet
    SessionNewItemsPath As String
    SesionNewItems  As Collection
    OdooBookPath    As String
End Type



Private this        As TDynamicTable

Private Sub Class_Initialize()
    With this
        .row = 1
        .column = 1

        Set .ws = ThisWorkbook.ActiveSheet
        Set .SesionNewItems = New Collection
    End With
    Set HeadersComp = New Headers
    Set ElementsComp = New Elements
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
        .SrcBookPath = FindLatestXLSXFile(Environ("USERPROFILE") & "\Downloads\")
        .NewItemsBookPath = Environ("USERPROFILE") & "\Documents\IMPORT" & "\items.xlsx"
        .SessionNewItemsPath = GetParentPath(ThisWorkbook.Path) & "Config" & "\items.xlsx"
        .OdooBookPath = GetParentPath(ThisWorkbook.Path) & "Config" & "\Odoo_Items.xls"
        With HeadersComp
            .values = Headers
        End With
        Update
    End With


End Sub

Public Sub Update()
    UpdateRanges
    With this
        .ws.Select
        HeadersComp.Update .ws
        ElementsComp.Update .ws, HeadersComp.values
        .ws.Columns.AutoFit
    End With
End Sub

'Dependencies: Row,Column, _
 Headers.values,Elements.Values

Private Sub UpdateRanges()

    With HeadersComp
        Set .Range = this.ws.Range( _
                this.ws.Cells(this.row, this.column), _
                this.ws.Cells(this.row, this.column + UBound(.values)) _
                )
    End With

    With ElementsComp
        Dim EmptyElements As Integer
        If .values.Count = 0 Then
            EmptyElements = 1
        End If
        Set .Range = this.ws.Range( _
                this.ws.Cells(this.row + 1, this.column), _
                this.ws.Cells(this.row + EmptyElements + .values.Count, this.column + UBound(HeadersComp.values)) _
                )
    End With
    With this
        Set .Range = Union( _
                HeadersComp.Range, _
                ElementsComp.Range _
                )
    End With
End Sub

Public Sub ShowPrintPreview()
    On Error Resume Next
    If ElementsComp.values.Count = 0 Then
        Debug.Print "Preview Error: No Element To Print Found"
        Exit Sub
    End If

    Application.Dialogs(xlDialogPrinterSetup).Show

    Dim PreviewLabel As New Label
    With PreviewLabel
        Set .product = ElementsComp.values.item(1)
        .Model.InitializeFormat light
        .ToRange

        '  !  @@@ Just First Time Production Run Time'  @@@

        'ConfigurePrintingPreferences .ws, .PrintArea.Address
        .ws.Activate
    End With
    On Error GoTo 0
End Sub


Public Sub PrintTest()
    On Error Resume Next
    If ElementsComp.values.Count = 0 Then
        Debug.Print "Print Test Error: No Element To Print Found"
        Exit Sub
    End If

    Dim ListLabel   As New Label
    Dim i           As Long

    Application.Dialogs(xlDialogPrinterSetup).Show

    With ListLabel
        Set .product = ElementsComp.values.item(1)
        .Model.InitializeFormat dark
        .ToRange


        .PrintArea.PrintOut Copies:=1
    End With

    On Error GoTo 0
End Sub


'.PaperSize: Label prop EXCEL = 4.65 * 0.79
Public Sub PrintList()

    If ElementsComp.values.Count = 0 Then
        Debug.Print "Print Error: No Element To Print Found"
        Exit Sub
    End If


    Dim PrintQty    As Integer: PrintQty = 1

    Dim i           As Long

    Dim ListLabel   As Label

    Application.Dialogs(xlDialogPrinterSetup).Show

    Dim selectedPrintingOption As String

    ' Retrieve the selected option
    selectedPrintingOption = MsgBox("USAR CORTE MANUAL?", vbYesNo, "PrintingMode")
    

    ' Check the selected option and proceed accordingly
    If selectedPrintingOption = "" Then
        Exit Sub
    End If


    For i = 1 To ElementsComp.values.Count
        Set ListLabel = New Label

        With ListLabel

            With ElementsComp
                If .values(i).Qty > 0 Then
                    PrintQty = WorksheetFunction.Ceiling(ElementsComp.values(i).Qty / 3, 1)
                End If
            End With

            Set .product = ElementsComp.values.item(i)

            .Model.InitializeFormat IIf(i Mod 2 = 0, dark, light)

            .ToRange

            On Error Resume Next
            .PrintArea.PrintOut Copies:=PrintQty
            On Error GoTo 0

            'Corte Manual Seleccionado
            If selectedPrintingOption = vbYes Then
                MsgBox "Corte y Presione Enter", vbOKOnly, "PAUSA DE CORTE"
            End If
            
        End With

        Set ListLabel = Nothing
    Next i


End Sub


Private Sub ElementsComp_RangeUpdated()
    UpdateRanges
End Sub

Private Sub HeadersComp_HeadersUpdated(ByVal Headers As Variant)
    UpdateRanges
End Sub

Public Property Get row() As Integer
    row = this.row
End Property

Public Property Let row(ByVal value As Integer)
    this.row = value
    UpdateRanges
End Property

Public Property Get column() As Integer
    column = this.column
End Property

Public Property Let column(ByVal value As Integer)
    this.column = value
    UpdateRanges
End Property

Public Property Get Range() As Range
    Set Range = this.Range
End Property

Public Property Set Range(ByVal value As Range)
    Set this.Range = value
End Property

Public Property Get SrcBookPath() As String
    SrcBookPath = this.SrcBookPath
End Property

Public Property Let SrcBookPath(ByVal value As String)
    this.SrcBookPath = value
End Property


'Latest Inplementation
Public Property Get OdooBookPath() As String
    OdooBookPath = this.OdooBookPath
End Property

Public Property Let OdooBookPath(ByVal value As String)
    this.OdooBookPath = value
End Property
'Latest Inplementation

Public Property Get NewItemsBookPath() As String
    NewItemsBookPath = this.NewItemsBookPath
End Property

Public Property Let NewItemsBookPath(ByVal value As String)
    this.NewItemsBookPath = value
End Property

Public Property Get TestBookPath() As String
    TestBookPath = this.TestBookPath
End Property

Public Property Let TestBookPath(ByVal value As String)
    this.TestBookPath = value
End Property

Public Property Get ws() As Worksheet
    Set ws = this.ws
End Property

Public Property Set ws(ByVal value As Worksheet)
    Set this.ws = value
End Property

Public Property Get SessionNewItemsPath() As String
    SessionNewItemsPath = this.SessionNewItemsPath
End Property

Public Property Let SessionNewItemsPath(ByVal RHS As String)
    this.SessionNewItemsPath = RHS
End Property

Private Sub IDataTable_Update()
    err.Raise 5                                  'TODO implement interface member
End Sub



Public Property Get SesionNewItems() As Collection
    Set SesionNewItems = this.SesionNewItems
End Property

Public Property Set SesionNewItems(ByVal RHS As Collection)
    Set this.SesionNewItems = RHS
End Property

