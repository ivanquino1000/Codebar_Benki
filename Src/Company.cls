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
    Address         As String
    RazonSocial     As String
    Ruc             As String
    Email           As String
    Phone           As String
    facebook        As String
    CodeTag         As String
    RevenuePercentage As Double
End Type

Private this        As TCompany


Private Sub Class_Initialize()
    With this
        .RevenuePercentage = 0.35
        .Address = "Teniente Ferre 104 Miraflores"
        .RazonSocial = "RANTY"
        .Ruc = "20609169053"
        .Email = "sorazaimport30@gmail.com"
        .Phone = "973182574"
        .facebook = "tinyurl.com/benki"
        .CodeTag = "B"
    End With
End Sub

Public Property Get CodeTag() As String
    CodeTag = this.CodeTag
End Property

Public Property Let CodeTag(ByVal RHS As String)
    this.CodeTag = RHS
End Property

Public Property Get Phone() As String
    Phone = this.Phone
End Property

Public Property Let Phone(ByVal RHS As String)
    this.Phone = RHS
End Property

Public Property Get facebook() As String
    facebook = this.facebook
End Property

Public Property Let facebook(ByVal RHS As String)
    this.facebook = RHS
End Property

Public Property Get Email() As String
    Email = this.Email
End Property

Public Property Let Email(ByVal RHS As String)
    this.Email = RHS
End Property

Public Property Get Address() As String
    Address = this.Address
End Property

Public Property Let Address(ByVal RHS As String)
    this.Address = RHS
End Property

Public Property Get Ruc() As String
    Ruc = this.Ruc
End Property

Public Property Let Ruc(ByVal RHS As String)
    this.Ruc = RHS
End Property

Public Property Get RazonSocial() As String
    RazonSocial = this.RazonSocial
End Property

Public Property Let RazonSocial(ByVal RHS As String)
    this.RazonSocial = RHS
End Property


Public Property Get RevenuePercentage() As Double
    RevenuePercentage = this.RevenuePercentage
End Property

Public Property Let RevenuePercentage(ByVal RHS As Double)
    this.RevenuePercentage = RHS
End Property
