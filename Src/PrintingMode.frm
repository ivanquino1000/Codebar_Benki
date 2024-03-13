VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PrintingMode 
   Caption         =   "UserForm1"
   ClientHeight    =   2625
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4200
   OleObjectBlob   =   "PrintingMode.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PrintingMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'Public selectedOption As String
'   Corte Manual _
    Impresion Ilimitada _
    Cancel


Public Sub UserForm_Initialize()
    ' Populate ComboBox with options
    
    selectedOption = "Cancel"
    ComboBox1.AddItem "Corte Manual"
    ComboBox1.AddItem "Impresion Ilimitada"
    ' Set ControlTipText for each ComboBox item
    'ComboBox1.ItemData(0) = "Pausa despues de imprimir un item para cortar."
    'ComboBox1.ItemData(1) = "Imprime la lista Completa sin pausas."
End Sub

Private Sub CommandButton1_Click()
    ' Logic for when the user clicks the OK button
    If ComboBox1.ListIndex <> -1 Then
        If ComboBox1.value = "Corte Manual" Then
            selectedOption = ComboBox1.value
        ElseIf ComboBox1.value = "Impresion Ilimitada" Then
            selectedOption = ComboBox1.value
        End If
        Unload Me ' Close the UserForm
    Else
        MsgBox "Por Favor Selecciona una Opcion."
    End If
End Sub

Private Sub CommandButton2_Click()
    ' Logic for when the user clicks the Cancel button
    selectedOption = "Cancel"
    Unload Me ' Close the UserForm
End Sub
