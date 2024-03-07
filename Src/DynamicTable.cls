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
Public Event elementAdded(ByVal Element As item)
Public Event TableUpdated()
Public Event ElementDeleted(ByVal Element As item)
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
    range           As range
    SrcBookPath     As String
    NewItemsBookPath As String
    TestBookPath    As String
    ws              As Worksheet
    TempBookPath    As String

End Type


Private this        As TDynamicTable

Private Sub Class_Initialize()
    With this
        .row = 1
        .column = 1

        Set .ws = ThisWorkbook.ActiveSheet
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
        .NewItemsBookPath = Environ("USERPROFILE") & "\Documents\IMPORT" & "Items.xlsx"
        .TempBookPath = ThisWorkbook.Path & "\Config" & "\TempData.xmls\"
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
        Set .range = this.ws.range( _
                    this.ws.Cells(this.row, this.column), _
                    this.ws.Cells(this.row, this.column + UBound(.values)) _
                    )
    End With

    With ElementsComp
        Dim EmptyElements As Integer
        If .values.Count = 0 Then
            EmptyElements = 1
        End If
        Set .range = this.ws.range( _
                    this.ws.Cells(this.row + 1, this.column), _
                    this.ws.Cells(this.row + EmptyElements + .values.Count, this.column + UBound(HeadersComp.values)) _
                    )
    End With
    With this
        Set .range = Union( _
                    HeadersComp.range, _
                    ElementsComp.range _
                    )
    End With
End Sub

Private Sub PrintTest()
    Dim printingRange As range
    Dim Label As New CodebarLabel
    Dim row, column As Integer
    'Dim LabelRange As Range
    Set printingRange = Label.range




    printingRange.PrintOut Copies:=1, Preview:=True, ActivePrinter:=PrinterName
End Sub
Private Sub PrintList()

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

Public Property Get range() As range
    Set range = this.range
End Property

Public Property Set range(ByVal value As range)
    Set this.range = value
End Property

Public Property Get SrcBookPath() As String
    SrcBookPath = this.SrcBookPath
End Property

Public Property Let SrcBookPath(ByVal value As String)
    this.SrcBookPath = value
End Property

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

Public Property Get TempBookPath() As String
    TempBookPath = this.TempBookPath
End Property

Public Property Let TempBookPath(ByVal RHS As String)
    this.TempBookPath = RHS
End Property

Private Sub IDataTable_Update()
    err.Raise 5                                  'TODO implement interface member
End Sub




