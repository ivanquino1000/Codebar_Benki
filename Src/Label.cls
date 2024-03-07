VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Label"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public Enum Models
    Default
    Reduced
    Full
End Enum

Private Type TLabel
    row             As Integer
    column          As Integer
    Model           As Model
    Instances       As Integer
    ws              As Worksheet
    product         As item
    Bussiness       As Company
End Type
Private this        As TLabel

Public Sub Class_Initialize()
    With this
        .row = 2
        .column = 2
        Set .Model = New Model
        .Instances = 0
        Set .ws = ThisWorkbook.Sheets("LabelSheet")
        Set .product = New item
        Set .Bussiness = New Company
    End With
End Sub

Public Sub build(ByVal item As item, ByVal Bussiness As Company)
    With this
        .row = 2
        .column = 2
        Set .Model = Models.Default
        .Instances = 0
        .ws = thiswowkbook.Sheets("LabelSheet")
        Set .product = item
        Set .Bussiness = Bussiness
    End With
End Sub

'TODO: ON EMPTY MANDATORY ITEMS.METADATA => EXIT AND SHOW AN WARNING MESSAGE

Public Sub ToRange()
    Dim LabelRange  As range
    Dim cell        As range

    With this
        'Define Label Range in Starting Position and Model Dimensions
        'Copy ModelGridMap to LabelRange
        Set LabelRange = .ws.Cells(.row, .column).Resize(.Model.RowSize, .Model.ColumnSize)
        .Model.CreateGrid .product.BoxPrice, .product.WholeSalePrice, .product.BoxQty, .product.Code
        LabelRange = .Model.GridMap
        
        For Each cell In LabelRange
            Dim key As String
            Dim keyVal As Variant
            Dim origin As String

            key = cell.Value2
KeyChecker:
            'Check CellVal and ItemObj Associated
            If key = "" Then
                GoTo ToNextCell
            End If
            If .Model.KeyInModel(.product, .Bussiness, key, keyVal) Then
                If keyVal = "" Then
                    If .Model.FallBack.Exists(key) Then
                        key = .Model.FallBack(key)
                        GoTo KeyChecker
                    End If
                    GoTo ToNextCell
                End If
                
            End If

            GoTo ToResizeCell
ToResizeCell:

            Dim Direction As String
            Dim Places As Integer
            With .Model
                If Not .SetResizeInst(key, Direction, Places) Then
                    GoTo ToFormatCell
                End If

            End With


            Redimension cell, Direction, Places
ToFormatCell:
            Dim format As FormatSettings

            With .Model
                If Not .SetFormat(key, format) Then
                    GoTo ToNextCell
                End If

            End With

            ApplyFormat cell, format
            cell.value = keyVal

ToNextCell:
        Next cell
    End With

    LabelRange.Columns.AutoFit
    'LabelRange.Rows.AutoFit
End Sub

Private Function Redimension(ByRef cell As range, ByVal Direction As String, ByVal Places As Integer)
    Select Case Direction
        Case "R"
            Set cell = cell.Resize(1, 1 + Places)
        Case "L"
            Set cell = cell.Offset(0, -Places).Resize(1, Places + 1)
        Case "U"
            Set cell = cell.Offset(-Places, 0).Resize(Places + 1, 1)
        Case "D"
            Set cell = cell.Resize(1 + Places, 1)
        Case Else
    End Select
    cell.Merge
End Function


Public Property Get row() As Integer
    row = this.row
End Property

Public Property Let row(ByVal RHS As Integer)
    this.row = RHS
End Property

Public Property Get column() As Integer
    column = this.column
End Property

Public Property Let column(ByVal RHS As Integer)
    this.column = RHS
End Property

Public Property Get Model() As Model
    Set Model = this.Model
End Property

Public Property Set Model(ByVal RHS As Model)
    Set this.Model = RHS
End Property

Public Property Get Instances() As Integer
    Instances = this.Instances
End Property

Public Property Let Instances(ByVal RHS As Integer)
    this.Instances = RHS
End Property

Public Property Get ws() As Worksheet
    Set ws = this.ws
End Property

Public Property Set ws(ByVal RHS As Worksheet)
    Set this.ws = RHS
End Property

Public Property Get product() As item
    Set product = this.product
End Property

Public Property Set product(ByVal RHS As item)
    Set this.product = RHS
End Property


Public Property Get Bussiness() As Company
    Set Bussiness = this.Bussiness
End Property

Public Property Set Bussiness(ByVal RHS As Company)
    Set this.Bussiness = RHS
End Property


