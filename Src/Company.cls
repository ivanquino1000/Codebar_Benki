VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Company"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TCompany
    Address As String
    RazonSocial As String
    Ruc As String
    Email As String
    Phone As String
    facebook As String

End Type

Private this As TCompany


Private Sub Class_Initialize()
    With this
        .Address = "Octavio Munoz Najar 202"
        .RazonSocial = "SORAZA"
        .Ruc = "20608875426"
        .Email = "sorazaimport30@gmail.com"
        .Phone = "958190331"
        .facebook = "tinyurl.com/soraza"
    End With
End Sub

Public Sub init(ByVal IRazonSocial As String, ByVal IAddress As String, ByVal IRuc As String)
    With this
        .Address = IAddress
        .RazonSocial = IRazonSocial
        .Ruc = IRuc
    End With
End Sub

Public Property Get Address() As String
    Address = this.Address
End Property

Public Property Let Address(ByVal RHS As String)
    this.Address = RHS
End Property

Public Property Get Ruc() As Long
    Ruc = this.Ruc
End Property

Public Property Let Ruc(ByVal RHS As Long)
    this.Ruc = RHS
End Property

Public Property Get RazonSocial() As String
    RazonSocial = this.RazonSocial
End Property

Public Property Let RazonSocial(ByVal RHS As String)
    this.RazonSocial = RHS
End Property

