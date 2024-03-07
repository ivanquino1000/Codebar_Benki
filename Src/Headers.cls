VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Headers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event HeadersUpdated(ByVal Headers As Variant)

Private Type THeaders
    values          As Variant
    Counter         As Integer
    range           As range
    format          As FormatSettings
End Type

Private this        As THeaders

Private Sub Class_Initialize()
    With this
        .values = Array("Head 1", "Head 2", "Head 3")
        Set .format = New FormatSettings
        .Counter = 1
        With .format
            .BgColor = RGB(214, 104, 83)
            .FontSize = 16
            .FontColor = vbWhite
            .FontName = "Bahnschrift SemiLight"
        End With
    End With
End Sub

Public Sub Update(ByVal ws As Worksheet)
    On Error GoTo err
    Dim Header      As Variant
    Dim i           As Long
    For Each Header In this.values

        ws.Cells( _
                this.range.row, this.range.column + i _
                ).value = Header

        i = i + 1
    Next Header
    Call ApplyFormat(this.range, this.format)

    Exit Sub

err:
    Debug.Print "UpdateHeaders.Err "; err.Description
End Sub

Public Property Get values() As Variant
    If IsObject(this.values) Then
        Set values = this.values
    Else
        values = this.values
    End If
End Property

Public Property Let values(ByVal RHS As Variant)
    this.values = RHS
    RaiseEvent HeadersUpdated(this.values)
End Property

Public Property Set values(ByVal RHS As Variant)
    Set this.values = RHS
    RaiseEvent HeadersUpdated(this.values)
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




