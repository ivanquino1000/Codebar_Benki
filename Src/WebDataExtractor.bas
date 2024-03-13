Attribute VB_Name = "WebDataExtractor"
Option Explicit

Public Function WebApp_Extract(ByVal WorkbookPath As String)


    Dim extractor_Script_Path As String: extractor_Script_Path = GetParentPath(WorkbookPath) & "WebUploader" & "\WebExport.js"
    Dim extractor_result_Path As String: extractor_result_Path = GetParentPath(WorkbookPath) & "WebUploader" & "\WebExportResult.txt"

    Shell "node """ & extractor_Script_Path & """", vbHide


    '   MS To be waited until Continue Execution
    Dim Max_Time_Limit As Integer: Max_Time_Limit = 100
    Dim startTime   As Double: startTime = Timer
    Dim passedTime  As Double
    Dim ExtrationResult As String
    Dim result      As String

    Do
        If dir(extractor_result_Path) <> "" Then
            ' Read the result from the file

            Open extractor_result_Path For Input As #1
            Line Input #1, result
            Close #1
            
            Debug.Print "At " & passedTime & ": The Result is " & result

            ' Check the result
            If result = "Success" Then
                ' Handle success condition
                Toast "Importacion Datos Web ", "EXITOSO"
                Exit Do
            ElseIf result = "Failed" Then
                ' Handle failure condition
                Toast "Importacion Datos Web ", " FALLIDO: " & vbCrLf & " Intentar Nuevamente o Manualmente ", 3
                Exit Do
            End If

        End If
        Application.Wait (Now + TimeValue("0:00:05"))

        passedTime = Timer - startTime

        ' Check if the elapsed time exceeds 1 minute (60 seconds)
        If passedTime >= 40 Then
            MsgBox "Timeout reached. Exiting loop."
            Toast "Importacion Datos Web ", " FALLIDO: \n Intentar Nuevamente o Manualmente ", 3
            Exit Do
        End If
    Loop
End Function
