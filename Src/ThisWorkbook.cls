VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'ON MAIN SHEET REMAP OVERRIDE DEFAULT KEYS AND BUTTONS
Const IntroductoryMessage As String = "Welcome to Codebar Printer You will first need to know " & _
        vbNewLine & "       PASOS A SEGUIR ANTES DE INICIAR:        " & _
        vbNewLine & "- Descargar de internet la base de datos" & _
        vbNewLine & "- Teclas Especiales " & vbNewLine & _
        vbNewLine & "      @ENTER NUMERICO: borra barra de busqueda" & _
        vbNewLine & "      @TEXTO *?*: Datos Opcionales de Relleno" & vbNewLine & _
        vbNewLine & "- Al finalizar enviar y subir los productos creados en esta sesion" & _
        vbNewLine & "-IvanQuino1000"

Const REQUIRED_DATA_DELETION_ERROR As String = "VALORES OBLIGATORIOS " & vbNewLine & _
        "   - No pueden Sen Elimindos Solo Actualizados " & vbNewLine

Const INMUTABLE_DATA_ERROR As String = "VALORES INMUNTABLES " & vbNewLine & _
        "   - No pueden Sen Cambiados " & vbNewLine

Const MISSING_VALUES As String = "VALORES INCOMPLETOS" & vbNewLine & _
        "   -Alguno de los Elemntos no se Encuentran Completos" & vbNewLine
Const VALIDATION_MESSAGE As String = "CONTINUAR?..." & vbNewLine & _
        "Seguro que Desea Ejecutar Esta Accion?" & vbNewLine

Const EXPORT_ERROR_MESSAGE As String = "ERROR AL EXPORTAR PRODUCTOS" & vbNewLine & _
        "   - Verifique Alguno de los siguientes" & vbNewLine

Const SUPPLIER_REQUEST_MSG As String = "INGRESAR PROVEEDOR" & vbNewLine & _
        "   - Indica el proveedor a utilizarse durante la Session (OPCIONAL)" & vbNewLine
Dim InitTime, EndTime As Double


Private WithEvents listTable As DynamicTable
Attribute listTable.VB_VarHelpID = -1
Private WithEvents searchTable As StaticTable
Attribute searchTable.VB_VarHelpID = -1
Private WithEvents resultsTable As StaticTable
Attribute resultsTable.VB_VarHelpID = -1
Private WithEvents databaseTable As StaticTable
Attribute databaseTable.VB_VarHelpID = -1

Private SessionBussiness As company
Private SessionSupplier As String


Private Sub Workbook_Open()

    With Application
        .EnableAutoComplete = False
        .ScreenUpdating = False
        .EnableEvents = False
        
        .OnKey "{ENTER}", "ClearSearchBar"

        ' ! Sheets Configuration - Zoom, BackgroundColor
        Dim sheet   As Worksheet

        For Each sheet In Me.Sheets
            With sheet.Cells
                .ClearContents
                .ClearFormats
                .Interior.Color = RGB(17, 21, 28)
            End With
        Next sheet

        With ThisWorkbook.Worksheets
            '.Activate
            ActiveWindow.Zoom = 60
        End With

        'Set & Create Tables
        InitTime = Timer

        Set SessionBussiness = New company
        SessionSupplier = InputBox(SUPPLIER_REQUEST_MSG)

        Set listTable = New DynamicTable
        Set databaseTable = New StaticTable
        Set searchTable = New StaticTable
        Set resultsTable = New StaticTable

        listTable.init 5, 2, ThisWorkbook.Sheets("MainSheet"), _
                Array( _
                "Cant", "Codigo", "Nombre", _
                "Adicional", "Proveedor", "Cant*Caja", "Precio*Caja", _
                "Costo?", "xMayor?", "Precio")
        databaseTable.init 1, 1, ThisWorkbook.Sheets("DataSheet"), _
                Array( _
                "Codigo", "Nombre", "Precio" _
                )
        searchTable.init 2, 13, ThisWorkbook.Sheets("MainSheet"), _
                Array( _
                "Codigo", "Nombre", "Precio" _
                )
        resultsTable.init 5, 13, ThisWorkbook.Sheets("MainSheet"), _
                Array( _
                "Codigo", "Nombre", "Precio" _
                )

        EndTime = Timer
        Debug.Print "UITableBuild- Performance:", EndTime - InitTime & " sec"
        WebApp_Upload (ThisWorkbook.Path)

        databaseTable.GetData

        ThisWorkbook.Sheets("MainSheet").Select

        .ScreenUpdating = True
        .EnableEvents = True
    End With
End Sub


Public Sub ListShowPrintPrev()

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        With Me.Sheets("LabelSheet")
            .Cells.ClearContents
            .Cells.ClearFormats
            .Parent.Windows(1).Zoom = 90
        End With

        listTable.ShowPrintPreview
        .ScreenUpdating = True
        .EnableEvents = True
    End With
End Sub

Public Sub ListPrintTest()

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        With Me.Sheets("LabelSheet")
            .Cells.ClearContents
            .Cells.ClearFormats
            .Parent.Windows(1).Zoom = 80
        End With

        listTable.PrintTest

        .ScreenUpdating = True
        .EnableEvents = True
    End With

End Sub

Public Sub ListPrintAll()

    If MsgBox(VALIDATION_MESSAGE & "IMPRIMIR TODOS LOS PRODUCTOS?", Buttons:=vbYesNo) = vbNo Then
        Exit Sub
    End If

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        With Me.Sheets("LabelSheet")
            .Cells.ClearContents
            .Cells.ClearFormats
            .Parent.Windows(1).Zoom = 80
        End With

        listTable.PrintList

        .ScreenUpdating = True
        .EnableEvents = True
    End With

End Sub

Public Sub ClearTable()
    With Application
        .EnableEvents = False
        .ScreenUpdating = False

        If MsgBox(VALIDATION_MESSAGE & "Limpiar Lista de Elementos", vbYesNo) = vbNo Then
            Exit Sub
        End If
        '  ! Elements In Table - Non Empty
        With listTable
            If .ElementsComp.values.Count < 1 Then
                Debug.Print "Clear Table - No Elements Context"
                Exit Sub
            End If
            '  ! Copy New Elements to Session Collection
            Dim NewSessionItem As Object
            For Each NewSessionItem In .ElementsComp.values

                With NewSessionItem
                    ' ! No New Item
                    If .IsNew = False Then
                        GoTo NextItem
                    End If
                    ' ! Copy To Session Col
                    listTable.SesionNewItems.Add NewSessionItem
                End With

NextItem:
            Next NewSessionItem

            Set .ElementsComp.values = New Collection
            .Update

        End With
        .EnableEvents = True
        .ScreenUpdating = True
    End With
End Sub

Private Function CreateCode(Optional ByVal code As String) As String

    Dim keyCode, CompanyId, ProductId As String
    Dim SerialNumber As Long

    ' Check if Code is provided
    If Not IsMissing(code) And code <> "" Then
        keyCode = ExtractLetter(code)
    Else
        ' If Code is missing, generate CompanyId and ProductId
        CompanyId = SessionBussiness.CodeTag
        ProductId = RandomLetter
        keyCode = CompanyId & ProductId
        GoTo SNGenerator
    End If

    ' Split KeyCode into CompanyId and ProductId
    If Len(keyCode) > 1 Then
        CompanyId = Mid(keyCode, 1, 1)
        ProductId = Mid(keyCode, 2, 1)
    Else
        CompanyId = SessionBussiness.CodeTag
        ProductId = Mid(keyCode, 1, 1)
        keyCode = CompanyId & ProductId    '###############################    PRODUCTION CHANGE   ##############
    End If
SNGenerator:
    '  KeyCode Does not exist in Recicle List -> SKIP Recicle Process
    If Not databaseTable.MissingElementsIds.Exists(keyCode) Then
        GoTo CheckListNewItems
    End If
    
    Dim MissingIds  As Variant
    
    MissingIds = databaseTable.MissingElementsIds.item(keyCode)
    
    '  None ID available to Recicle on MissingElementsIDs -> SKIP Recicle Process
    If IsEmpty(MissingIds) Then
        GoTo CheckListNewItems
    End If

    Dim RemainingElement As Variant
    ' Get corresponding id available after comparing with session IDs
    RemainingElement = FindFirstNonIncludedElement(MissingIds, listTable.ElementsComp.NewCodeIds.item(keyCode))

    '  ! Error in Iteration
    If RemainingElement = Null Then
        Debug.Print "FATAL ERRROR COMPARE ARRAYS"
        Exit Function
    End If

    '  ! All Missing In Use
    If RemainingElement = "" Then
        GoTo DB_OR_SESSION_LAST
    End If


    ' ! Remaining found
    If RemainingElement <> "" Then
        SerialNumber = RemainingElement
        CreateCode = CodeBuilder(keyCode, CInt(SerialNumber))
        UpdateListNewProductsIds keyCode, SerialNumber
        Exit Function
    End If

CheckListNewItems:
    ' KeyCode has been used in the current Session for a new product
    If Not listTable.ElementsComp.NewCodeIds.Exists(keyCode) Then
        GoTo LAST_ELEMENT_CHECKER
    End If

    Dim NewCodeIds  As Variant
    NewCodeIds = listTable.ElementsComp.NewCodeIds.item(keyCode)

    ' ! Session IDs Null
    If IsEmpty(NewCodeIds) Then
LAST_ELEMENT_CHECKER:
        '  ! Data Base Last Element  - Don't Exist
        If Not databaseTable.LastElementId.Exists(keyCode) Then
            SerialNumber = 1
            CreateCode = CodeBuilder(keyCode, CInt(SerialNumber))
            UpdateListNewProductsIds keyCode, SerialNumber
            Exit Function
        End If
        '  ! Data Base Last Element - Empty
        If IsEmpty(databaseTable.LastElementId.item(keyCode)) Then
            SerialNumber = 1
            CreateCode = CodeBuilder(keyCode, CInt(SerialNumber))
            UpdateListNewProductsIds keyCode, SerialNumber
            Exit Function
        End If
        '  ! Last Element Available
        SerialNumber = databaseTable.LastElementId.item(keyCode) + 1
        CreateCode = CodeBuilder(keyCode, CInt(SerialNumber))
        UpdateListNewProductsIds keyCode, SerialNumber
        Exit Function
    End If

DB_OR_SESSION_LAST:
    '  ! Data Base Last Element  - Don't Exist
    If Not databaseTable.LastElementId.Exists(keyCode) Then
        SerialNumber = WorksheetFunction.Max(listTable.ElementsComp.NewCodeIds(keyCode)) + 1
        CreateCode = CodeBuilder(keyCode, CInt(SerialNumber))
        UpdateListNewProductsIds keyCode, SerialNumber
        Exit Function
    End If
    '  ! Data Base Last Element - Empty
    If IsEmpty(databaseTable.LastElementId.item(keyCode)) Then
        SerialNumber = WorksheetFunction.Max(listTable.ElementsComp.NewCodeIds(keyCode)) + 1
        CreateCode = CodeBuilder(keyCode, CInt(SerialNumber))
        UpdateListNewProductsIds keyCode, SerialNumber
        Exit Function
    End If
    '  ! Highest ID - Session OR DataBase
    Dim SessionID, DataBaseID As Variant


    SessionID = WorksheetFunction.Max(listTable.ElementsComp.NewCodeIds.item(keyCode))
    DataBaseID = databaseTable.LastElementId.item(keyCode)

    If SessionID > DataBaseID Then
        SerialNumber = SessionID + 1
    Else
        SerialNumber = DataBaseID + 1
    End If

    CreateCode = CodeBuilder(keyCode, CInt(SerialNumber))
    UpdateListNewProductsIds keyCode, SerialNumber


End Function

'Takes the Item Name _
 - can be in two formats:(CODE-NAME/sDESCCRIPTION) /(NAME/sDESCCRIPTION) _
 - Return NAME AND DESCTIPTION IF MATCHED Other wise empty strings
Private Function SplitDescription(ByVal itemName As String, ByRef Name As String, ByRef desc As String)

    If itemName = "" Then
        Name = "Empty Item Name"
        desc = "Empty Item Description"
        Exit Function
    End If

    Dim pattern     As String
    Dim regex       As Object
    Dim matches     As Object
    Dim match       As Object

    ' Define the pattern
    ' ! Group 3: NAME
    ' ! Group 4: DESCRIPTION
    pattern = "(([A-Za-z]{1,3}\d{1,3}))?(\s*-\s*)?(\b\w+\b)\s?(.*)"

    ' Create a regular expression object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = True
        .IgnoreCase = False
        .MultiLine = True
        .pattern = pattern
    End With

    ' Check if the input string matches the pattern
    If regex.test(itemName) Then
        ' Extract the product name from the matched pattern
        Set matches = regex.Execute(itemName)
        If matches.Count > 0 Then
            ' Get the first match (group 3 contains the product name)
            Set match = matches(0)
            Name = UCase(match.SubMatches(3))
            If match.SubMatches.Count > 3 Then
                desc = UCase(match.SubMatches(4))
            End If
        Else
            ' No match found
            Name = "No Match"
            desc = "No Match"
        End If
    Else
        ' Input string doesn't match the pattern
        Name = "Unvalid Format"
        desc = "Unvalid Format"
    End If

End Function

Private Sub UpdateListNewProductsIds(ByVal key As Variant, ByVal ID As Variant)
    Dim CodeIdHolder As Variant

    If IsEmpty(listTable.ElementsComp.NewCodeIds.item(key)) Then
        CodeIdHolder = Array()
    Else
        CodeIdHolder = listTable.ElementsComp.NewCodeIds.item(key)
    End If

    PushToVariant CodeIdHolder, ID

    If Not listTable.ElementsComp.NewCodeIds.Exists(key) Then
        listTable.ElementsComp.NewCodeIds.Add key, CodeIdHolder
    End If

    listTable.ElementsComp.NewCodeIds.item(key) = CodeIdHolder
End Sub




Private Sub Workbook_SheetBeforeDoubleClick(ByVal sh As Object, ByVal Target As Range, Cancel As Boolean)

    If sh.Name = "MainSheet" Then

        '@ResultsTable - AddElement

        If Not Intersect(Target, resultsTable.ElementsRange) Is Nothing Then

            Cancel = True

            With Application
                .EnableEvents = False
                .ScreenUpdating = False


                Dim rowId As Integer

                Dim CodeRange As Range
                Dim NameRange As Range
                Dim PriceRange As Range

                Dim product As New item

                rowId = Target.row - resultsTable.row
                ' ! Support_Ranges_References
                With resultsTable.ElementsRange
                    Set CodeRange = .Columns(1)
                    Set NameRange = .Columns(2)
                    Set PriceRange = .Columns(3)
                End With

                ' ! Results Table - Missing Data Checker
                If CodeRange.Cells(rowId).Value2 = "" Or _
                        NameRange.Cells(rowId).Value2 = "" Or _
                        PriceRange.Cells(rowId).Value2 = "" Then
                    MsgBox MISSING_VALUES & " $Tabla de Resultados: (CODIGO,NOMBRE O PRECIO )NO ENCONTRADOS "
                    GoTo InvalidValues
                End If

                ' ! Add Product to List Table

                With product

                    Dim Name As String
                    Dim desc As String
                    '   Set Unmutable Item Elements
                    .Supplier = SessionSupplier

                    '   OPTION - Create / Select
                    If Not Intersect(Target, CodeRange) Is Nothing Then

                        '  On Create New Item Option
                        Dim UserInputName As String, UserInputPrice As Double
                        '   - Display Inputs Form
                        UserInputName = RequestNewItemName
                        UserInputPrice = RequestNewItemPrice
                        
                        ' Get the name and description from User Input
                        SplitDescription UserInputName, Name, desc

                        '   - Display Price Input Form
                        .code = CreateCode(CodeRange.Cells(rowId).Value2)

                        .Name = .code & "-" & Name
                        .Description = desc

                        .SellingPrice = UserInputPrice

                        .IsNew = True

                    Else

                        '  On Select Current Item

                        ' Get the name and description from Results Table
                        SplitDescription NameRange.Cells(rowId).Value2, Name, desc
                        
                        .code = CodeRange.Cells(rowId).Value2
                        .Name = .code & "-" & Name
                        .Description = desc
                        .code = CodeRange.Cells(rowId).Value2

                        .SellingPrice = PriceRange.Cells(rowId).Value2

                    End If
                    
                End With

                listTable.ElementsComp.values.Add product
                listTable.Update
InvalidValues:
                .EnableEvents = True
                .ScreenUpdating = True
            End With
        End If

        '@Seach_Table / CREATE_CODE_RANDOM
        If Not Intersect(Target, searchTable.ElementsRange) Is Nothing Then
            Cancel = True
            With Application
                .EnableEvents = False
                .ScreenUpdating = False

                '@Seach_Input_Checker
                With searchTable
                    ' ! Empty_UI_Input - Name / Price
                    If IsEmpty(.ElementsRange(, 2)) Or .ElementsRange(, 2) = "" Or _
                            IsEmpty(.ElementsRange(, 3)) Or .ElementsRange(, 3) = "" Then
                        MsgBox MISSING_VALUES & "   - Nombre " & vbNewLine & "  - Precio"
                        GoTo SearchInvalidCase
                    End If
                    ' ! User_Validation
                    If MsgBox(VALIDATION_MESSAGE & "Crear Producto Nuevo", vbYesNo) = vbNo Then
                        GoTo SearchInvalidCase
                    End If

                    ' ! Item_Data_Filling
                    Dim searchProduct As New item

                    With searchProduct
                        .Name = ClearSearchInput(searchTable.ElementsRange.Columns(2).Value2)
                        .SellingPrice = ClearSearchInput(searchTable.ElementsRange.Columns(3).Value2)
                        .code = CreateCode
                        .Supplier = SessionSupplier
                        .IsNew = True
                    End With

                End With

                listTable.ElementsComp.values.Add searchProduct
                listTable.Update

SearchInvalidCase:
                .EnableEvents = True
                .ScreenUpdating = True
            End With

        End If

        '@ListTable / Delete
        If Not Intersect(Target, listTable.ElementsComp.Range) Is Nothing Then
            Cancel = True
            With Application
                .EnableEvents = False
                .ScreenUpdating = False
                If listTable.ElementsComp.values.Count = 0 Then
                    GoTo InvalidCase
                End If
                Dim UserValidation As Long
                UserValidation = MsgBox(VALIDATION_MESSAGE & "Eliminar Producto de la Lista de Impresion", vbYesNo)
                If UserValidation = vbNo Then
                    GoTo InvalidCase
                End If
                Dim ItemId As Integer

                With listTable
                    ItemId = Target.row - .row
                    With .ElementsComp

                        If .values.item(ItemId).IsNew = True Then

                            Dim keyCode, CodeId, code As Variant
                            code = .values.item(ItemId).code
                            keyCode = ExtractLetter(code)
                            CodeId = CInt(ExtractNumber(code))
                            .NewCodeIds.item(keyCode) = RemoveFromArr(.NewCodeIds.item(keyCode), CodeId)

                        End If


                        .values.Remove ItemId

                    End With


                    .Update
                End With
InvalidCase:
                .EnableEvents = True
                .ScreenUpdating = True
            End With

        End If

    End If

End Sub

Private Function RequestNewItemPrice() As Double
    Dim Price        As String
    Price = InputBox("Inserte el Precio del Item:", "NUEVO ITEM - PRECIO")
    
    If Not IsNumeric(Price) Then
        MsgBox "El precio debe ser numerico /n Intente Denuevo "
        RequestNewItemPrice = RequestNewItemPrice()
        Exit Function
    End If
    
    If Price <= 0 Then
        MsgBox "El precio debe ser Positivo /n Intente Denuevo "
        RequestNewItemPrice
        Exit Function
    End If

    RequestNewItemPrice = CDbl(Price)
End Function

Private Function RequestNewItemName() As String
    
    Dim Name        As String
    Name = InputBox("Inserte el Nombre del Item: ", "NUEVO ITEM - NOMBRE")
    If Name = "" Then
        MsgBox "El Nombre no Puede estar Vacio /n Intente Denuevo "
        RequestNewItemName = RequestNewItemName()
        Exit Function
    End If

    RequestNewItemName = Name
End Function

Private Function CalculateRatio(ByVal Cost As Variant, ByVal Price As Variant, ByVal BatchTag As String) As Double
    Dim ratio       As Double
    Select Case BatchTag
        Case "Whole"
            ratio = ((Price / Cost) - 1) * 0.5
        Case Else
            ratio = ((Price / Cost) - 1) * 0.25
    End Select

    If ratio < SessionBussiness.RevenuePercentage Then
        CalculateRatio = -1
        Exit Function
    End If
    CalculateRatio = ratio + 1
End Function

'TODO: ExactMatch - AutoInsertKey - Highlight Option

Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal Target As Range)
    Dim cell        As Range

    '@UI - MainSheet
    If sh.Name = "MainSheet" Then

        '@ListTable - Update Values

        If Not Intersect(Target, listTable.ElementsComp.Range) Is Nothing Then
            With Application
                .EnableEvents = False
                .ScreenUpdating = False

                With listTable

                    '@No Elemnts To Update
                    If listTable.ElementsComp.values.Count = 0 Then
                        GoTo EmptyElements
                    End If

                    Dim product As Range

                    Dim RemovalOrderIndexes As Variant: RemovalOrderIndexes = Array()

                    For Each product In Target.Rows
                        Dim CompleteProduct As Range
                        Set CompleteProduct = sh.Cells(product.row, .column).Resize(, .HeadersComp.Range.Columns.Count)
                        Dim EMPTY_STATE As Boolean: EMPTY_STATE = True

                        Dim i, j As Long

                        For i = LBound(CompleteProduct.Value2, 1) To UBound(CompleteProduct.Value2, 1)
                            For j = LBound(CompleteProduct.Value2, 2) To UBound(CompleteProduct.Value2, 2)
                                If Not IsEmpty(CompleteProduct.Value2(i, j)) Then
                                    EMPTY_STATE = False
                                    GoTo HasElements
                                End If
                            Next j
                        Next i

                        If EMPTY_STATE = True Then

                            PushToVariant RemovalOrderIndexes, product.row - .row

                            GoTo NextItem
                        End If
HasElements:

                        Dim COST_CHANGED As Boolean: COST_CHANGED = False
                        For Each cell In CompleteProduct.Cells

                            Dim prop As String
                            prop = .ElementsComp.HeaderMap(.HeadersComp.Range(, cell.column - .column + 1).value)

                            Dim ElementId As Long: ElementId = cell.row - .row
                            Dim element As Object
                            Set element = .ElementsComp.values.item(ElementId)

                            Select Case prop


                                Case "Qty", "BoxQty", "SellingPrice"

                                    '# Mandatory Data /BREAK AFTER
                                    If prop = "SellingPrice" And cell.Value2 = "" Then
                                        MsgBox REQUIRED_DATA_DELETION_ERROR & "Precio de Venta - Retornara al valor Anterior"
                                        GoTo InvalidValues
                                    End If

                                    If Not IsNumeric(cell.Value2) Then
                                        GoTo InvalidValues
                                    End If

                                    CallByName element, prop, VbLet, cell.Value2
                                Case "WholeSalePrice"

                                    If Not IsNumeric(cell.Value2) Then
                                        GoTo InvalidValues
                                    End If

                                    ' ! WholeSale Price - Deletion Action
                                    If cell.Value2 = "" Then
                                        CallByName element, prop, VbLet, cell.Value2
                                        GoTo InvalidValues
                                    End If
                                    ' ! WholeSale Price - Update Action
                                    If cell.Value2 > CallByName(element, "SellingPrice", VbGet) Then
                                        CallByName element, prop, VbLet, cell.Value2 / 12
                                    Else
                                        CallByName element, prop, VbLet, cell.Value2
                                    End If


                                Case "BoxPrice"

                                    If Not IsNumeric(cell.Value2) Or cell.Value2 = "" Then
                                        GoTo InvalidValues
                                    End If
                                    ' ! Box Price - Deletion Action
                                    If cell.Value2 = "" Then
                                        CallByName element, prop, VbLet, cell.Value2
                                        GoTo InvalidValues
                                    End If
                                    ' ! Box Price - Update Action

                                    Dim boxbatch As Integer: boxbatch = 24
                                    If Not IsEmpty(CallByName(element, "BoxQty", VbGet)) Then
                                        If IsNumeric(CallByName(element, "BoxQty", VbGet)) And CallByName(element, "BoxQty", VbGet) <> "" Then
                                            boxbatch = CallByName(element, "BoxQty", VbGet)
                                        End If
                                    End If

                                    If cell.Value2 > CallByName(element, "SellingPrice", VbGet) Then
                                        CallByName element, prop, VbLet, cell.Value2 / boxbatch
                                    Else
                                        CallByName element, prop, VbLet, cell.Value2
                                    End If

                                Case "Cost"

                                    If Not IsNumeric(cell.Value2) Then
                                        GoTo InvalidValues
                                    End If
                                    ' ! Cost Price - Deletion Action
                                    If cell.Value2 = "" Then
                                        CallByName element, prop, VbLet, cell.Value2
                                        GoTo InvalidValues
                                    End If
                                    ' ! Cost Price - Update Action
                                    If cell.Value2 <> CallByName(element, prop, VbGet) Then
                                        COST_CHANGED = True
                                    End If
                                    CallByName element, prop, VbLet, cell.Value2

                                Case "Name"
                                    If cell.Value2 = "" Then
                                        MsgBox REQUIRED_DATA_DELETION_ERROR & prop & " / Retornara al valor Anterior"
                                        GoTo InvalidValues
                                    End If

                                    'Format Item Name and Description

                                    Dim UName As String, UDesc As String, PrevCode As String

                                    PrevCode = CallByName(element, "Code", VbGet)
                                    SplitDescription cell.Value2, UName, UDesc

                                    'ITEMNAME = CODE-NAME
                                    CallByName element, prop, VbLet, PrevCode & "-" & StrConv(UName, vbProperCase)
                                    CallByName element, "Description", VbLet, StrConv(UDesc, vbProperCase)


                                Case Else
                                    '# Mandatory Data /BREAK AFTER
                                    If prop = "Code" Then
                                        If cell.Value2 = "" Then
                                            MsgBox REQUIRED_DATA_DELETION_ERROR & prop & " / Retornara al valor Anterior"
                                            GoTo InvalidValues
                                        End If
                                        If cell.Value2 <> CallByName(element, "Code", VbGet) Then
                                            MsgBox INMUTABLE_DATA_ERROR & prop & " / Retornara al valor Anterior"
                                            GoTo InvalidValues
                                        End If

                                        GoTo InvalidValues
                                    End If

                                    If prop = "Description" And UDesc <> "" Then
                                        GoTo InvalidValues
                                    End If

                                    CallByName element, prop, VbLet, StrConv(cell.Value2, vbProperCase)
                            End Select
InvalidValues:
                        Next cell

                        ' ! Whole Sale Prices - Auto Completion on Cost Change
                        If COST_CHANGED = True Then

                            Dim WholeAdjutment, BoxAdjustment As Double
                            Dim Update_Product_Cost As Variant
                            Update_Product_Cost = CallByName(element, "Cost", VbGet)

                            WholeAdjutment = CalculateRatio(Update_Product_Cost, CallByName(element, "SellingPrice", VbGet), "Whole")
                            BoxAdjustment = CalculateRatio(Update_Product_Cost, CallByName(element, "SellingPrice", VbGet), "Box")

                            If WholeAdjutment <> -1 Then
                                CallByName element, "WholeSalePrice", VbLet, WorksheetFunction.RoundUp((Update_Product_Cost * WholeAdjutment) / 0.5, 0) * 0.5

                            End If
                            If BoxAdjustment <> -1 Then
                                CallByName element, "BoxPrice", VbLet, WorksheetFunction.RoundUp((Update_Product_Cost * BoxAdjustment) / 0.5, 0) * 0.5 'Update_Product_Cost * BoxAdjustment

                            End If
                            ' ! On Incongruent Prices don'T change User Inputs
                        End If
NextItem:
                    Next product

                    If Not IsEmpty(RemovalOrderIndexes) Then
                        '@Asump: Always in Order
                        For i = LBound(RemovalOrderIndexes) To UBound(RemovalOrderIndexes)

                            ' ! NewCodeIDs - Deletion From Session
                            Dim DeletionItem As Object
                            Set DeletionItem = .ElementsComp.values.item(RemovalOrderIndexes(i) - i)
                            If DeletionItem.IsNew = True Then
                                listTable.ElementsComp.NewCodeIds(ExtractLetter(DeletionItem.code)) = RemoveFromArr(listTable.ElementsComp.NewCodeIds(ExtractLetter(DeletionItem.code)), ExtractNumber(DeletionItem.code))
                            End If

                            .ElementsComp.values.Remove RemovalOrderIndexes(i) - i
                        Next i
                    End If

                    .ElementsComp.Update sh, .HeadersComp.values
                    .Update
EmptyElements:
                End With
                .EnableEvents = True
                .ScreenUpdating = True
            End With

        End If

        '@UI [SearchBar]
        If Not Intersect(Target, searchTable.ElementsRange) Is Nothing Then
            With Application
                .EnableEvents = False
                .ScreenUpdating = False
                Dim item As Object

                If Target.Rows.Count > 1 Then
                    For Each item In Target.Rows
                        GoTo UpdateProperties
                    Next item
                End If

                Set item = Target
UpdateProperties:

                For Each cell In item
                    If IsEmpty(item) Then
                        GoTo FillResults
                    End If

                    Dim Header, UserInput As String
                    Dim HeaderIndex As Integer
                    With searchTable
                        HeaderIndex = cell.column - .column
                        Header = listTable.ElementsComp.HeaderMap( _
                                .HeadersComp.values(HeaderIndex) _
                                )
                        UserInput = cell.value
                    End With

                    If UserInput = "" Then
                        GoTo FillResults
                    End If

                    Select Case Header
                        Case "Code":
                            cell.value = UserInput


                        Case "Name":
                            cell.value = "*" & UserInput

                            '                        Case "SellingPrice":
                            '                            If IsNumeric(UserInput) & UserInput > 30 Then
                            '                                cell.value = "<=" & UserInput
                            '                            Else
                            '                                cell.value = ">=" & UserInput
                            '                            End If
                        Case Else
                            Debug.Print "User Input: Header Out of Index"
                    End Select
                    cell.Offset(, 1).Activate
                Next cell

FillResults:

                resultsTable.ClearPrevRange
                Dim Elements As Integer
                databaseTable.Range.AdvancedFilter Action:=xlFilterCopy, _
                        criteriarange:=searchTable.Range, CopyToRange:=resultsTable.HeadersComp.Range, _
                        Unique:=False

                With resultsTable
                    Elements = .ws.Cells(.ws.Rows.Count, .column).End(xlUp).row - .row
                    .ElementsCounter = Elements
                    .ElementsRange.Sort key1:=.ElementsRange.Columns(UBound(.HeadersComp.values) + 1), _
                            order1:=xlDescending, Header:=xlNo
                End With


                .EnableEvents = True
                .ScreenUpdating = True
            End With
        End If

        '@Headers - Recovery
        If (Not Intersect(Target, resultsTable.HeadersComp.Range) Is Nothing) Or (Not Intersect(Target, searchTable.HeadersComp.Range) Is Nothing) Or (Not Intersect(Target, listTable.HeadersComp.Range) Is Nothing) Then
            With Application
                .EnableEvents = False
                resultsTable.HeadersComp.Update sh
                listTable.HeadersComp.Update sh
                searchTable.HeadersComp.Update sh
                .EnableEvents = True
            End With
        End If

    End If

    '#Data Sheet
    If sh.Name = "DataSheet" Then

        '@Headers - Recovery
        If Not Intersect(Target, databaseTable.HeadersComp.Range) Is Nothing Then
            With Application
                .EnableEvents = False
                databaseTable.HeadersComp.Update sh
                .EnableEvents = True
            End With
        End If

    End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    If listTable Is Nothing Then
        Debug.Print "listTable is not present. Closing workbook normally."
        GoTo CloseWorkbook
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ClearTable

    ' ! Session New Items - Empty
    If Not listTable.SesionNewItems.Count > 0 Then
        Debug.Print "No Elements Created in this Session"
        GoTo SessionBackupStorage
    End If

    Dim ExternalBookLastRow As Integer

    Dim NewItemsCounter As Integer


    ' ! Session New Items - Send To "Items" Import File
    Dim SessionItem As Object
    With listTable

        Dim CodeArr As Variant: CodeArr = Array()
        Dim NameArr As Variant: NameArr = Array()
        Dim PriceArr As Variant: PriceArr = Array()
        Dim CostArr As Variant: CostArr = Array()
        Dim WholesalePriceArr As Variant: WholesalePriceArr = Array()

        For Each SessionItem In .SesionNewItems

            With SessionItem
                ' ! Item Info Array Population
                PushToVariant CodeArr, .code
                PushToVariant NameArr, .Name
                PushToVariant PriceArr, .SellingPrice
                PushToVariant CostArr, .Cost
                PushToVariant WholesalePriceArr, .WholeSalePrice

                NewItemsCounter = NewItemsCounter + 1
            End With

NextItem:
        Next SessionItem



        Dim NewItemsBook As Workbook
        Dim OdooItemsBook As Workbook
        ' ! New Items Book  - Setting


        If Not ValidBookPath(.NewItemsBookPath) Then
            ' ! Destiny items Path  -  Invalid
            MsgBox EXPORT_ERROR_MESSAGE & "   - Libro Items No Encontrado Enviando A: " & vbNewLine & _
                    .NewItemsBookPath & " --> " & .SessionNewItemsPath

            ' ! Local Session items Path - TRY
            If Not ValidBookPath(.SessionNewItemsPath) Then
                ' ! Local Session items - Invalid  Exit Procedure
                If MsgBox(EXPORT_ERROR_MESSAGE & "   - Libro Respaldo Items No Encontrado " & vbNewLine & _
                        "SALIR?, NO SE GUARDARAN LOS DATOS DE LA SESSION", vbYesNo) = vbNo Then
                    Cancel = True
                    Exit Sub
                End If

                ' ! Local Session items - Available
                Set NewItemsBook = Workbooks.Open(.SessionNewItemsPath)
            End If
        Else
            ' ! Destiny items Path  -  Available
            Set NewItemsBook = Workbooks.Open(.NewItemsBookPath)
        End If


        'ODOO Excel_ITEMS_path_Verification
        If Not ValidBookPath(.OdooBookPath) Then
            ' ! Destiny items Path  -  Invalid
            MsgBox EXPORT_ERROR_MESSAGE & "   - Libro Odoo_Items No Encontrado & vbNewLine & ERROR: AL EXPORTAR ODOO ITEMS"
        Else
            ' ! Destiny items Path  -  Available
            Set OdooItemsBook = Workbooks.Open(.OdooBookPath)
        End If




        ' ! Mandatory Destiny Columns References - New Items - ARCA DIGITAL

        '    ! Name - A2
        '    ! Code - B
        '    ! Unit - NIU - E
        '    ! Exchange Type - PEN - F
        '    ! Selling Price - G
        '    ! IGV SELL - 10 - H
        '    ! IGV_INCLUDED- SI - I
        '    ! IGV PURCHASE- 10 - k
        '    ! STOCK- 1- L
        '    ! STOCK_MIN- 0 - M
        '    ! DESCRIPTION - P
        '    ! BARCODE- T

        With NewItemsBook.Worksheets(1)
            ' ! Get Headers
            Dim ItemsExportHeader As Variant
            ItemsExportHeader = .Rows(1).Value2


            ' ! Clear Contents
            .Cells.ClearContents
            .Cells.ClearFormats

            ' ! Return Headers to Export File
            .Rows(1).Value2 = ItemsExportHeader

            ' ! CONST Address Population
            .Range(.Columns("E").Cells(2), .Columns("E").Cells(2 + NewItemsCounter - 1)) = "NIU"
            .Range(.Columns("F").Cells(2), .Columns("F").Cells(2 + NewItemsCounter - 1)) = "PEN"
            .Range(.Columns("H").Cells(2), .Columns("H").Cells(2 + NewItemsCounter - 1)) = 10
            .Range(.Columns("I").Cells(2), .Columns("I").Cells(2 + NewItemsCounter - 1)) = "SI"
            .Range(.Columns("K").Cells(2), .Columns("K").Cells(2 + NewItemsCounter - 1)) = 10
            .Range(.Columns("L").Cells(2), .Columns("L").Cells(2 + NewItemsCounter - 1)) = 1
            .Range(.Columns("M").Cells(2), .Columns("M").Cells(2 + NewItemsCounter - 1)) = 0
            ' ! VARIABLE Address Population
            .Range(.Columns("A").Cells(2), .Columns("A").Cells(2 + NewItemsCounter - 1)).Value2 = Application.Transpose(NameArr)
            .Range(.Columns("B").Cells(2), .Columns("B").Cells(2 + NewItemsCounter - 1)).Value2 = Application.Transpose(CodeArr)
            .Range(.Columns("G").Cells(2), .Columns("G").Cells(2 + NewItemsCounter - 1)).Value2 = Application.Transpose(PriceArr)
            .Range(.Columns("P").Cells(2), .Columns("P").Cells(2 + NewItemsCounter - 1)).Value2 = Application.Transpose(NameArr)
            .Range(.Columns("T").Cells(2), .Columns("T").Cells(2 + NewItemsCounter - 1)).Value2 = Application.Transpose(CodeArr)

        End With

        ' ! ODOO - Mandatory Destiny Columns References - New Items

        '    ! External ID - A2
        '    ! Name- B2
        '    ! Almacenable- Almacenable - C2
        '    ! Internal Reference - D2
        '    ! Barcode- E2
        '    ! Sales Price- F2
        '    ! Cost- G2
        '    ! Precio por Mayor- H2
        '    ! Disponible en pdv - I2

        With OdooItemsBook.Worksheets(1)
            ' ! Get Headers
            Dim OdooItemsExportHeaders As Variant
            OdooItemsExportHeaders = .Rows(1).Value2


            ' ! Clear Contents
            .Cells.ClearContents
            .Cells.ClearFormats

            ' ! Return Headers to Export File
            .Rows(1).Value2 = OdooItemsExportHeaders

            ' ! CONST Address Population
            .Range(.Columns("C").Cells(2), .Columns("C").Cells(2 + NewItemsCounter - 1)) = "Almacenable"
            .Range(.Columns("I").Cells(2), .Columns("I").Cells(2 + NewItemsCounter - 1)) = "True"
            ' ! VARIABLE Address Population
            .Range(.Columns("A").Cells(2), .Columns("A").Cells(2 + NewItemsCounter - 1)).Value2 = Application.Transpose(CodeArr)
            .Range(.Columns("B").Cells(2), .Columns("B").Cells(2 + NewItemsCounter - 1)).Value2 = Application.Transpose(NameArr)
            .Range(.Columns("D").Cells(2), .Columns("D").Cells(2 + NewItemsCounter - 1)).Value2 = Application.Transpose(CodeArr)
            .Range(.Columns("E").Cells(2), .Columns("E").Cells(2 + NewItemsCounter - 1)).Value2 = Application.Transpose(CodeArr)
            .Range(.Columns("F").Cells(2), .Columns("F").Cells(2 + NewItemsCounter - 1)).Value2 = Application.Transpose(PriceArr)
            .Range(.Columns("G").Cells(2), .Columns("G").Cells(2 + NewItemsCounter - 1)).Value2 = Application.Transpose(CostArr)
            .Range(.Columns("H").Cells(2), .Columns("H").Cells(2 + NewItemsCounter - 1)).Value2 = Application.Transpose(WholesalePriceArr)


        End With

    End With

    NewItemsBook.Close True
    OdooItemsBook.Close True

SessionBackupStorage:

    ' ! Session Items - To Session Backup

    Dim DbItemCounter As Integer
    '  !  Database Sheet - Last Row - 1 (header)
    DbItemCounter = databaseTable.ElementsCounter

    ' ! Session Items - Send To "SessionBackup" Local File

    With databaseTable


        Dim DbCodeArr As Variant: DbCodeArr = Application.WorksheetFunction.Transpose(.ElementsRange.Columns(1).Value2)
        Dim DbNameArr As Variant: DbNameArr = Application.WorksheetFunction.Transpose(.ElementsRange.Columns(2).Value2)
        Dim DbPriceArr As Variant: DbPriceArr = Application.WorksheetFunction.Transpose(.ElementsRange.Columns(3).Value2)


        Dim ItemsBook As Workbook

        ' ! New Items Book  - Setting

        If Not ValidBookPath(.SessionBackupPath) Then
            ' ! Destiny Path  -  Invalid
            MsgBox EXPORT_ERROR_MESSAGE & "   - Libro Session No Encontrado Cancelando" & vbNewLine
            Cancel = True
            Exit Sub

        Else
            ' ! Destiny  Path  -  Available
            Set ItemsBook = Workbooks.Open(.SessionBackupPath)
        End If

        ' ! Mandatory Destiny Columns References

        '    ! Code - B
        '    ! Name - E
        '    ! Selling Price - K

        With ItemsBook.Worksheets(1)
            ' ! Get Headers
            Dim DbItemsExportHeader As Variant
            DbItemsExportHeader = .Rows(3).Value2


            ' ! Clear Contents
            .Cells.ClearContents
            .Cells.ClearFormats

            ' ! Return Headers to Export File
            .Rows(3).Value2 = DbItemsExportHeader

            ' ! VARIABLE Address Population

            Dim CompleteDbCodes As Variant: CompleteDbCodes = JoinArrays(DbCodeArr, CodeArr)
            Dim CompleteDbNames As Variant: CompleteDbNames = JoinArrays(DbNameArr, NameArr, False)
            Dim CompleteDbPrices As Variant: CompleteDbPrices = JoinArrays(DbPriceArr, PriceArr, False)


            ' Change Index In case 0 or 1
            Debug.Print ItemsBook.Worksheets(1).Range("B4").Resize(UBound(CompleteDbCodes) - LBound(CompleteDbCodes) + 1, 1).Address

            .Range("B4").Resize(UBound(CompleteDbCodes) - LBound(CompleteDbCodes) + 1, 1).Value2 = Application.Transpose(CompleteDbCodes)
            .Range("E4").Resize(UBound(CompleteDbNames) - LBound(CompleteDbNames) + 1, 1).Value2 = Application.Transpose(CompleteDbNames)
            .Range("K4").Resize(UBound(CompleteDbPrices) - LBound(CompleteDbPrices) + 1, 1).Value2 = Application.Transpose(CompleteDbPrices)

        End With

    End With
    ItemsBook.Close True

    Application.ScreenUpdating = True
    Application.EnableEvents = True
CloseWorkbook:
    On Error GoTo 0
    'Me.Close False
End Sub


