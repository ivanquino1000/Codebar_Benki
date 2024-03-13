Attribute VB_Name = "WebDataUpload"
Option Explicit

Public Function WebApp_Upload(ByVal WorkbookPath As String)


    Dim uploader_script_path As String: extractor_Script_Path = GetParentPath(WorkbookPath) & "WebUploader" & "\WebUpload.js"
    Dim uploader_result_path As String: extractor_result_Path = GetParentPath(WorkbookPath) & "WebUploader" & "\WebUploadResult.txt"

    Shell "node """ & uploader_script_path & """", vbHide


    '   MS To be waited until Continue Execution
    Dim Max_Time_Limit As Integer: Max_Time_Limit = 100
    Dim startTime   As Double: startTime = Timer
    Dim passedTime  As Double
    Dim ExtrationResult As String
    Dim result      As String
    
    frmProgressBar.Progress passedTime, 40, "Extrayendo Datos de la Web"
    frmProgressBar.Show
    
    Do
        If dir(extractor_result_Path) <> "" Then
            ' Read the result from the file

            Open uploader_result_path For Input As #1
            Line Input #1, result
            Close #1

            Debug.Print "At " & passedTime & ": The Result is " & result

            ' Check the result
            If result = "Success" Then
                ' Handle success condition
                Toast "Importacion Datos Web ", "EXITOSO"
                frmProgressBar.Recaption ("Completado exitosamente")
                Exit Do
            ElseIf result = "Failed" Then
                ' Handle failure condition
                Toast "Importacion Datos Web ", " FALLIDO: " & vbCrLf & " Intentar Nuevamente o Manualmente ", 3
                frmProgressBar.Recaption ("Error Detectado Vuelva a Intentar")
                Exit Do
            End If

        End If
        Application.Wait (Now + TimeValue("0:00:05"))

        passedTime = Timer - startTime

        ' Check if the elapsed time exceeds 1 minute (60 seconds)
        If passedTime >= 40 Then
            Debug.Print "Timeout reached. Exiting loop."
            Toast "Importacion Datos Web ", " FALLIDO: \n Intentar Nuevamente o Manualmente ", 3
            frmProgressBar.Recaption ("Tiempo Exedido Vuelva a intentarlo ")
            Exit Do
        End If
    Loop
    
End Function

