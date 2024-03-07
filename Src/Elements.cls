VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Elements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event ValuesUpdated(ByVal values As Collection)

Private PreviousRange As range

Private Type TElements
    values As Collection
    Counter As Integer
    range As range
    format As FormatSettings
    NewFormat As FormatSettings
    HeaderMap As Dictionary
    CodeIds As Dictionary
    NewCodeIds As Dictionary
End Type


Private this As TElements

Public Property Get CodeIds() As Dictionary
    Set CodeIds = this.CodeIds
End Property

Public Property Set CodeIds(ByVal RHS As Dictionary)
    Set this.CodeIds = RHS
End Property

Public Property Get NewCodeIds() As Dictionary
    Set NewCodeIds = this.NewCodeIds
End Property

Public Property Set NewCodeIds(ByVal RHS As Dictionary)
    Set this.NewCodeIds = RHS
End Property

Private Sub Class_Initialize()
    With this
        Set .values = New Collection
        Set .format = New FormatSettings
        Set .NewFormat = New FormatSettings
        .Counter = .values.Count
        
        With .format
            .BgColor = RGB(249, 244, 245)
            .FontColor = vbBlack
            .FontName = "Bahnschrift SemiLight"
            .FontSize = 14
        End With
        With .NewFormat
            .BgColor = RGB(249, 244, 245)
            .FontColor = RGB(255, 87, 20)
            .FontName = "Bahnschrift SemiLight"
            .FontSize = 14
        End With
        Set .HeaderMap = New Dictionary
        Set .CodeIds = New Dictionary
        Set .NewCodeIds = New Dictionary
        
        With .HeaderMap
            .Add "Cant", "Qty"
            .Add "Codigo", "Code"
            .Add "Nombre", "Name"
            .Add "Adicional", "Description"
            .Add "Proveedor", "Supplier"
            .Add "Costo?", "Cost"
            .Add "xMayor?", "WholeSalePrice"
            .Add "Precio", "SellingPrice"
            .Add "Cant*Caja", "BoxQty"
            .Add "Precio*Caja", "BoxPrice"
        End With
    End With
End Sub

Public Sub Update(ByVal ws As Worksheet, ByVal Headers As Variant)

' ClearContents and Format in previous Range
    If Not PreviousRange Is Nothing Then
        With PreviousRange
            .ClearContents
            .ClearFormats
            .Interior.Color = RGB(17, 21, 28)
        End With
    End If
    
        
'Show Elemts.values and Special Formating
    Call ApplyFormat(this.range, this.format)
    
    With this
        If .values.Count = 0 Or UBound(Headers) = 0 Then
            Exit Sub
        End If
        ReDim ElemArr(1 To .values.Count, 0 To UBound(Headers))
        Dim elem As Object
    
        Dim row, column As Integer: row = 1
        For Each elem In .values
            
            For column = 0 To UBound(Headers)
                
                ElemArr(row, column) = CallByName(elem, HeaderMap(Headers(column)), VbGet)
                
            Next column
            row = row + 1
            
            Dim ElemRange As range
            Set ElemRange = range.Rows(row - 1) 'ws.range(Cells(range.row, range.column), _
                                     Cells(range.row, range.Columns.Count))
            If elem.IsNew Then
                Call ApplyFormat(ElemRange, .NewFormat)
            End If
        Next elem
        this.range.value = ElemArr
        .range.EntireRow.RowHeight = 40
    End With
    
    Set PreviousRange = this.range
End Sub

Public Function MissingId(Optional ByVal NewItems As Variant) As Integer

End Function


Public Property Get HeaderMap() As Object
    Set HeaderMap = this.HeaderMap
End Property

Public Property Set HeaderMap(ByVal RHS As Object)
    Set this.HeaderMap = RHS
End Property

Public Property Get values() As Collection
    Set values = this.values
End Property

Public Property Set values(ByVal RHS As Collection)
    Set this.values = RHS
    RaiseEvent ValuesUpdated(this.values)
End Property

Public Property Get Counter() As Integer
    Counter = this.Counter
End Property

Public Property Let Counter(ByVal RHS As Integer)
    this.Counter = RHS
End Property

Public Property Get range() As range
    Set range = this.range
End Property

Public Property Set range(ByVal RHS As range)
    Set this.range = RHS
End Property

Public Property Get format() As FormatSettings
    Set format = this.format
End Property

Public Property Set format(ByVal RHS As FormatSettings)
    Set this.format = RHS
End Property

Public Property Get NewFormat() As FormatSettings
    Set NewFormat = this.NewFormat
End Property

Public Property Set NewFormat(ByVal RHS As FormatSettings)
    Set NewFormat = RHS
    End Property




