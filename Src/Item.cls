VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event elementAdded(ByVal element_Id As String, ByVal element_Code As String, ByVal element_Name As String, ByVal element_price As Double)



Private Type TItem
    Id              As Variant
    BoxQty          As Variant
    Qty             As Variant
    BoxPrice        As Variant
    Cost            As Variant
    SellingPrice    As Variant
    WholeSalePrice  As Variant
    Code            As String
    Description     As String
    Name            As String
    Supplier        As String
    IsNew           As Boolean
End Type

Private this        As TItem



Private Sub Class_Initialize()
    With this
        '.Id = 0
        '   .Qty = 1
        .IsNew = False
        .Code = "A000"
        .Name = "baseItemName Base desc"
        '     .Description = "Empty Desc"
        '    .Supplier = "Los Incas"
        '   .Cost = 9.99
        '  .WholeSalePrice = 999.99
        .SellingPrice = 9999.99
        '    .BoxQty = 24
        '   .BoxPrice = 99.99
    End With
End Sub
Public Sub Create(ByVal ICode As String, ByVal IName As String, ByVal IDescription As String, _
        ByVal ISupplier As String, ByVal ICost As String, ByVal IWholeSalePrice As String, _
        ByVal IPrice _
        )
    With this
        .Code = ICode
        .Name = IName
        .Description = IDescription
        .Supplier = ISupplier
        .Cost = ICost
        .WholeSalePrice = IWholeSalePrice
        .SellingPrice = IPrice
    End With

End Sub


Public Property Get Id() As Variant
    If IsObject(this.Id) Then
        Set Id = this.Id
    Else
        Id = this.Id
    End If
End Property

Public Property Let Id(ByVal RHS As Variant)
    this.Id = RHS
End Property

Public Property Set Id(ByVal RHS As Variant)
    Set this.Id = RHS
End Property

Public Property Get BoxQty() As Variant
    If IsObject(this.BoxQty) Then
        Set BoxQty = this.BoxQty
    Else
        BoxQty = this.BoxQty
    End If
End Property

Public Property Let BoxQty(ByVal RHS As Variant)
    this.BoxQty = RHS
End Property

Public Property Set BoxQty(ByVal RHS As Variant)
    Set this.BoxQty = RHS
End Property

Public Property Get Qty() As Variant
    If IsObject(this.Qty) Then
        Set Qty = this.Qty
    Else
        Qty = this.Qty
    End If
End Property

Public Property Let Qty(ByVal RHS As Variant)
    this.Qty = RHS
End Property

Public Property Set Qty(ByVal RHS As Variant)
    Set this.Qty = RHS
End Property

Public Property Get BoxPrice() As Variant
    If IsObject(this.BoxPrice) Then
        Set BoxPrice = this.BoxPrice
    Else
        BoxPrice = this.BoxPrice
    End If
End Property

Public Property Let BoxPrice(ByVal RHS As Variant)
    this.BoxPrice = RHS
End Property

Public Property Set BoxPrice(ByVal RHS As Variant)
    Set this.BoxPrice = RHS
End Property

Public Property Get Cost() As Variant
    If IsObject(this.Cost) Then
        Set Cost = this.Cost
    Else
        Cost = this.Cost
    End If
End Property

Public Property Let Cost(ByVal RHS As Variant)
    this.Cost = RHS
End Property

Public Property Set Cost(ByVal RHS As Variant)
    Set this.Cost = RHS
End Property

Public Property Get SellingPrice() As Variant
    If IsObject(this.SellingPrice) Then
        Set SellingPrice = this.SellingPrice
    Else
        SellingPrice = this.SellingPrice
    End If
End Property

Public Property Let SellingPrice(ByVal RHS As Variant)
    this.SellingPrice = RHS
End Property

Public Property Set SellingPrice(ByVal RHS As Variant)
    Set this.SellingPrice = RHS
End Property

Public Property Get WholeSalePrice() As Variant
    If IsObject(this.WholeSalePrice) Then
        Set WholeSalePrice = this.WholeSalePrice
    Else
        WholeSalePrice = this.WholeSalePrice
    End If
End Property

Public Property Let WholeSalePrice(ByVal RHS As Variant)
    this.WholeSalePrice = RHS
End Property

Public Property Set WholeSalePrice(ByVal RHS As Variant)
    Set this.WholeSalePrice = RHS
End Property

Public Property Get Code() As String
    Code = this.Code
End Property

Public Property Let Code(ByVal RHS As String)
    this.Code = RHS
End Property

Public Property Get Description() As String
    Description = this.Description
End Property

Public Property Let Description(ByVal RHS As String)
    this.Description = RHS
End Property

Public Property Get Name() As String
    Name = this.Name
End Property

Public Property Let Name(ByVal RHS As String)
    this.Name = RHS
End Property

Public Property Get Supplier() As String
    Supplier = this.Supplier
End Property

Public Property Let Supplier(ByVal RHS As String)
    this.Supplier = RHS
End Property

Public Property Get IsNew() As Boolean
    IsNew = this.IsNew
End Property

Public Property Let IsNew(ByVal RHS As Boolean)
    this.IsNew = RHS
End Property
